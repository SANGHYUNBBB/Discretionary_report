from __future__ import annotations

import bisect
import os
import re
from dataclasses import dataclass
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Dict, List, Optional, Tuple

from openpyxl import Workbook, load_workbook
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


# ============================================================
# 환경설정
# ============================================================

BASE_DIR = Path.cwd()

INPUT_PATTERNS = [
    "kb26q2*.xlsx",
    "kb26q2*.xlsm",
    "KB26q2*.xlsx",
    "KB26q2*.xlsm",
    "KB거래내역*.xlsx",
    "KB거래내역*.xlsm",
    "kb거래내역*.xlsx",
    "kb거래내역*.xlsm",
]

EXCHANGE_DIR_NAMES = [
    "exchange_rate",
    "Exchange_Rate",
    "EXCHANGE_RATE",
]

OUTPUT_SUFFIX = "_정리.xlsx"

OUTPUT_COLUMNS = [
    "계좌번호",
    "계약자명",
    "PF명",
    "구분",
    "종목명",
    "매매일자",
    "수량",
    "매매단가",
    "매매금액",
    "위탁매매수수료",
    "각종세금",
]

CURRENCY_CODES = {
    "USD",
    "JPY",
    "HKD",
    "CNY",
    "EUR",
    "GBP",
    "AUD",
    "CAD",
    "CHF",
    "SGD",
    "TWD",
}

ZERO = Decimal("0")
HUNDRED = Decimal("100")


# ============================================================
# 환율 데이터 구조
# ============================================================

@dataclass
class ExchangeTable:
    dates_ord: List[int]
    rates: List[Decimal]

    def lookup(self, target_date: date) -> Decimal:
        """
        거래일과 동일한 날짜의 환율을 우선 사용한다.
        동일 날짜가 없으면 가장 가까운 이전 날짜의 환율을 사용한다.
        """
        target_ordinal = target_date.toordinal()

        position = bisect.bisect_right(
            self.dates_ord,
            target_ordinal,
        ) - 1

        if position < 0:
            raise ValueError(
                f"{target_date} 이전에 사용할 수 있는 환율이 없습니다."
            )

        return self.rates[position]


# ============================================================
# 공통 함수
# ============================================================

def clean_text(value) -> str:
    if value is None:
        return ""

    return str(value).strip()


def normalize_transaction(value) -> str:
    """
    거래종류의 띄어쓰기를 제거해서 비교한다.

    예:
    배당금 입금 -> 배당금입금
    해외원천세 출금 -> 해외원천세출금
    """
    return re.sub(
        r"\s+",
        "",
        clean_text(value),
    )


def safe_excel_text(value) -> str:
    return ILLEGAL_CHARACTERS_RE.sub(
        "",
        clean_text(value),
    )


def set_plain_text(cell, value) -> None:
    """
    값이 수식으로 인식되지 않도록 일반 문자열로 저장한다.
    """
    cell.value = safe_excel_text(value)
    cell.data_type = "s"


def to_decimal(value) -> Decimal:
    """
    엑셀 값을 Decimal로 변환한다.
    공란, 숫자가 아닌 값, NaN, 무한대는 0으로 처리한다.
    """
    if value is None or value == "":
        return ZERO

    if isinstance(value, Decimal):
        result = value

    elif isinstance(value, bool):
        result = Decimal(int(value))

    else:
        try:
            result = Decimal(
                str(value).replace(",", "").strip()
            )

        except (
            InvalidOperation,
            AttributeError,
            ValueError,
        ):
            return ZERO

    if not result.is_finite():
        return ZERO

    return result


def parse_date_safe(value) -> Optional[date]:
    if value is None or value == "":
        return None

    if isinstance(value, datetime):
        return value.date()

    if isinstance(value, date):
        return value

    text = clean_text(value)

    for date_format in (
        "%Y/%m/%d",
        "%Y-%m-%d",
        "%Y.%m.%d",
        "%Y%m%d",
    ):
        try:
            return datetime.strptime(
                text,
                date_format,
            ).date()

        except ValueError:
            continue

    return None


def extract_account_info(
    a1_value,
) -> Tuple[str, str]:
    """
    A1 값에서 계좌번호와 계약자명을 분리한다.
    """
    text = clean_text(a1_value)

    account_match = re.search(
        r"(\d+(?:-\d+)+)",
        text,
    )

    if account_match:
        account_number = account_match.group(1)
        holder = text[
            account_match.end():
        ].strip()

        return account_number, holder

    number_match = re.search(
        r"(\d{8,})",
        text,
    )

    account_number = (
        number_match.group(1)
        if number_match
        else ""
    )

    holder = re.sub(
        r"^[\d\-\s]+",
        "",
        text,
    ).strip()

    return account_number, holder


# ============================================================
# 파일 및 폴더 검색
# ============================================================

def find_exchange_dir(
    base_dir: Path,
) -> Optional[Path]:
    for directory_name in EXCHANGE_DIR_NAMES:
        directory_path = base_dir / directory_name

        if (
            directory_path.exists()
            and directory_path.is_dir()
        ):
            return directory_path

    return None


def find_input_files(
    base_dir: Path,
) -> List[Path]:
    """
    KB 원본 파일만 찾는다.
    정리 결과, 임시파일, Excel 잠금 파일은 제외한다.
    """
    files: List[Path] = []

    for pattern in INPUT_PATTERNS:
        files.extend(
            base_dir.glob(pattern)
        )

    result = {
        file_path.resolve()
        for file_path in files
        if not file_path.name.startswith("~$")
        and not file_path.stem.endswith("_정리")
        and "_작성중" not in file_path.stem
    }

    return sorted(result)


# ============================================================
# 환율 처리
# ============================================================

def adjust_fx_rate(
    currency_code: str,
    rate: Decimal,
) -> Decimal:
    """
    JPY 환율은 100엔당 환율이므로
    100으로 나누어 1엔당 환율로 바꾼다.
    """
    currency_code = clean_text(
        currency_code
    ).upper()

    if currency_code == "JPY":
        return rate / HUNDRED

    return rate


def extract_currency_code_from_filename(
    file_path: Path,
) -> str:
    """
    환율 파일명에서 통화코드를 추출한다.

    예:
    USD.xlsx
    USD_환율.xlsx
    환율_USD.xlsx
    """
    stem_upper = (
        file_path.stem.strip().upper()
    )

    match = re.search(
        r"(?<![A-Z])([A-Z]{3})(?![A-Z])",
        stem_upper,
    )

    if match:
        return match.group(1)

    return stem_upper


def load_exchange_rates(
    exchange_dir: Optional[Path],
) -> Dict[str, ExchangeTable]:
    """
    exchange_rate 폴더의 환율 파일을 읽는다.

    환율 파일 구조:
    - 파일명에 통화코드 포함
    - 첫 번째 시트 사용
    - A열: 날짜
    - C열: 환율
    - 10행부터 데이터
    """
    rate_map: Dict[str, ExchangeTable] = {}

    if exchange_dir is None:
        return rate_map

    candidate_files: List[Path] = []

    for extension in (
        "*.xlsx",
        "*.xlsm",
    ):
        candidate_files.extend(
            exchange_dir.glob(extension)
        )

        candidate_files.extend(
            exchange_dir.glob(
                f"**/{extension}"
            )
        )

    unique_files: List[Path] = []
    seen_paths = set()

    for file_path in candidate_files:
        resolved_path = str(
            file_path.resolve()
        )

        if resolved_path in seen_paths:
            continue

        if file_path.name.startswith("~$"):
            continue

        seen_paths.add(resolved_path)
        unique_files.append(file_path)

    for file_path in unique_files:
        currency_code = (
            extract_currency_code_from_filename(
                file_path
            )
        )

        if currency_code not in CURRENCY_CODES:
            continue

        workbook = load_workbook(
            file_path,
            data_only=True,
            read_only=True,
        )

        try:
            worksheet = workbook[
                workbook.sheetnames[0]
            ]

            date_rate_pairs: List[
                Tuple[int, Decimal]
            ] = []

            for row_number in range(
                10,
                worksheet.max_row + 1,
            ):
                exchange_date = parse_date_safe(
                    worksheet[
                        f"A{row_number}"
                    ].value
                )

                exchange_rate = to_decimal(
                    worksheet[
                        f"C{row_number}"
                    ].value
                )

                if (
                    exchange_date is None
                    or exchange_rate == ZERO
                ):
                    continue

                date_rate_pairs.append(
                    (
                        exchange_date.toordinal(),
                        exchange_rate,
                    )
                )

        finally:
            workbook.close()

        if not date_rate_pairs:
            raise ValueError(
                "환율 파일에서 날짜와 환율을 "
                f"찾지 못했습니다: {file_path}"
            )

        date_rate_pairs.sort(
            key=lambda item: item[0]
        )

        rate_map[
            currency_code
        ] = ExchangeTable(
            dates_ord=[
                item[0]
                for item in date_rate_pairs
            ],
            rates=[
                item[1]
                for item in date_rate_pairs
            ],
        )

    return rate_map


def get_effective_rate(
    row: dict,
    fx_tables: Dict[str, ExchangeTable],
) -> Decimal:
    """
    환율 적용 순서:

    1순위:
    원본 KB 파일의 환율 값이 있고 0이 아니면 해당 값 사용

    2순위:
    원본 환율이 공란 또는 0이면 exchange_rate 폴더에서 조회
    - 통화구분과 동일한 환율 파일 사용
    - 거래일과 같은 날짜 우선
    - 같은 날짜가 없으면 가장 가까운 이전 날짜 사용

    JPY:
    원본 환율이든 폴더 환율이든 최종적으로 100으로 나눈다.
    """
    currency_code = clean_text(
        row.get("통화구분")
    ).upper()

    if currency_code in (
        "",
        "KRW",
    ):
        return Decimal("1")

    # 1순위: 원본 환율
    original_rate = to_decimal(
        row.get("환율")
    )

    if original_rate != ZERO:
        return adjust_fx_rate(
            currency_code,
            original_rate,
        )

    # 2순위: exchange_rate 폴더 환율
    trade_date = parse_date_safe(
        row.get("거래일자")
    )

    if trade_date is None:
        raise ValueError(
            "원본 환율이 공란 또는 0이고, "
            "거래일자도 해석할 수 없습니다."
        )

    exchange_table = fx_tables.get(
        currency_code
    )

    if exchange_table is None:
        available_codes = (
            ", ".join(
                sorted(fx_tables.keys())
            )
            if fx_tables
            else "없음"
        )

        raise KeyError(
            f"원본 환율이 공란 또는 0이며, "
            f"통화코드 {currency_code}에 해당하는 "
            "환율 파일도 찾지 못했습니다. "
            f"사용 가능한 통화코드: {available_codes}"
        )

    folder_rate = exchange_table.lookup(
        trade_date
    )

    return adjust_fx_rate(
        currency_code,
        folder_rate,
    )


# ============================================================
# 세금 합계
# ============================================================

def full_tax_sum(
    row: dict,
) -> Decimal:
    """
    농특세/부가세
    + 지방소득세
    + 거래세 등
    + 소득세
    + 양도세
    """
    return (
        to_decimal(
            row.get("농특세/부가세")
        )
        + to_decimal(
            row.get("지방소득세")
        )
        + to_decimal(
            row.get("거래세 등")
        )
        + to_decimal(
            row.get("소득세")
        )
        + to_decimal(
            row.get("양도세")
        )
    )


def foreign_dividend_tax_sum(
    row: dict,
) -> Decimal:
    """
    거래세 등 + 소득세 + 양도세
    """
    return (
        to_decimal(
            row.get("거래세 등")
        )
        + to_decimal(
            row.get("소득세")
        )
        + to_decimal(
            row.get("양도세")
        )
    )


def domestic_dividend_tax_sum(
    row: dict,
) -> Decimal:
    """
    통화구분이 없는 배당금입금 세금:

    거래세 등 + 소득세 + 양도세 + 지방소득세
    """
    return (
        to_decimal(
            row.get("거래세 등")
        )
        + to_decimal(
            row.get("소득세")
        )
        + to_decimal(
            row.get("양도세")
        )
        + to_decimal(
            row.get("지방소득세")
        )
    )


def income_tax_sum(
    row: dict,
) -> Decimal:
    """
    농특세/부가세 + 지방소득세 + 소득세 + 양도세
    """
    return (
        to_decimal(
            row.get("농특세/부가세")
        )
        + to_decimal(
            row.get("지방소득세")
        )
        + to_decimal(
            row.get("소득세")
        )
        + to_decimal(
            row.get("양도세")
        )
    )


# ============================================================
# 직전 거래 외화예수금
# ============================================================

def previous_foreign_cash(
    previous_row: Optional[dict],
) -> Decimal:
    if previous_row is None:
        return ZERO

    return to_decimal(
        previous_row.get("외화예수금")
    )


def foreign_cash_delta(
    row: dict,
    previous_row: Optional[dict],
) -> Decimal:
    """
    이번 외화예수금 - 직전 거래의 외화예수금
    """
    current_cash = to_decimal(
        row.get("외화예수금")
    )

    previous_cash = previous_foreign_cash(
        previous_row
    )

    return current_cash - previous_cash


# ============================================================
# 거래종류별 계산
# ============================================================

def calculate_row(
    row: dict,
    previous_row: Optional[dict],
    fx_tables: Dict[str, ExchangeTable],
) -> Tuple[
    Decimal,
    Decimal,
    Decimal,
    Decimal,
    Decimal,
    str,
]:
    """
    반환 순서:

    수량,
    매매단가,
    매매금액,
    위탁매매수수료,
    각종세금,
    검토메시지
    """
    transaction_original = clean_text(
        row.get("거래종류")
    )

    transaction = normalize_transaction(
        transaction_original
    )

    currency_code = clean_text(
        row.get("통화구분")
    ).upper()

    quantity = to_decimal(
        row.get("수량")
    )

    unit_price = to_decimal(
        row.get("단가")
    )

    trade_amount = to_decimal(
        row.get("거래금액")
    )

    settlement_amount = to_decimal(
        row.get("정산금액")
    )

    foreign_settlement_amount = to_decimal(
        row.get("외화정산금액")
    )

    domestic_fee = to_decimal(
        row.get("수수료")
    )

    foreign_fee = to_decimal(
        row.get("국외수수료")
    )

    transaction_tax = to_decimal(
        row.get("거래세 등")
    )

    # ========================================================
    # 매수 / 매도
    # ========================================================

    if transaction in {
        "매수",
        "매도",
    }:
        exchange_rate = get_effective_rate(
            row,
            fx_tables,
        )

        output_unit_price = (
            unit_price
            * exchange_rate
        )

        return (
            quantity,
            output_unit_price,
            quantity * output_unit_price,
            foreign_fee * exchange_rate,
            transaction_tax * exchange_rate,
            "",
        )

    # ========================================================
    # 주식장내매수 / 주식장내매도
    # ========================================================

    if transaction in {
        "주식장내매수",
        "주식장내매도",
    }:
        return (
            quantity,
            unit_price,
            quantity * unit_price,
            domestic_fee,
            full_tax_sum(row),
            "",
        )

    # ========================================================
    # 외화매수 / 외화매도
    # ========================================================

    if transaction in {
        "외화매수",
        "외화매도",
    }:
        exchange_rate = get_effective_rate(
            row,
            fx_tables,
        )

        return (
            foreign_settlement_amount,
            exchange_rate,
            settlement_amount,
            domestic_fee,
            foreign_dividend_tax_sum(row),
            "",
        )

    # ========================================================
    # 배당금입금
    # ========================================================

    if transaction == "배당금입금":
        if currency_code in (
            "",
            "KRW",
        ):
            return (
                ZERO,
                ZERO,
                trade_amount,
                domestic_fee,
                domestic_dividend_tax_sum(row),
                "",
            )

        exchange_rate = get_effective_rate(
            row,
            fx_tables,
        )

        return (
            foreign_settlement_amount,
            exchange_rate,
            foreign_settlement_amount
            * exchange_rate,
            domestic_fee,
            foreign_dividend_tax_sum(row),
            "",
        )

    # ========================================================
    # 해외원천세 출금
    # ========================================================

    if transaction == "해외원천세출금":
        exchange_rate = get_effective_rate(
            row,
            fx_tables,
        )

        return (
            foreign_settlement_amount,
            exchange_rate,
            ZERO,
            ZERO,
            foreign_settlement_amount
            * exchange_rate,
            "",
        )

    # ========================================================
    # 예탁금이용료 입금
    # ========================================================

    if transaction == "예탁금이용료입금":
        if currency_code in (
            "",
            "KRW",
        ):
            return (
                ZERO,
                ZERO,
                trade_amount,
                ZERO,
                income_tax_sum(row),
                "",
            )

        exchange_rate = get_effective_rate(
            row,
            fx_tables,
        )

        return (
            ZERO,
            exchange_rate,
            foreign_settlement_amount
            * exchange_rate,
            ZERO,
            income_tax_sum(row),
            "",
        )

    # ========================================================
    # 이자소득세추징 출금
    # ========================================================

    if transaction == "이자소득세추징출금":
        return (
            ZERO,
            ZERO,
            ZERO,
            ZERO,
            income_tax_sum(row),
            "",
        )

    # ========================================================
    # 비용충당외화매수 출금
    # ========================================================

    if transaction == "비용충당외화매수출금":
        return (
            ZERO,
            ZERO,
            ZERO,
            settlement_amount,
            ZERO,
            "",
        )

    # ========================================================
    # ADR FEE 출금
    # ========================================================

    if transaction == "ADRFEE출금":
        exchange_rate = get_effective_rate(
            row,
            fx_tables,
        )

        previous_cash = previous_foreign_cash(
            previous_row
        )

        warning_message = ""

        if previous_row is None:
            warning_message = (
                "직전 거래가 없어 ADR FEE 계산 기준 "
                "외화예수금을 0으로 사용했습니다."
            )

        return (
            ZERO,
            ZERO,
            ZERO,
            previous_cash * exchange_rate,
            ZERO,
            warning_message,
        )

    # ========================================================
    # 대체입금
    # ========================================================

    if transaction == "대체입금":
        return (
            ZERO,
            ZERO,
            trade_amount,
            ZERO,
            ZERO,
            "",
        )

    # ========================================================
    # 대체입고
    # ========================================================

    if transaction == "대체입고":
        return (
            quantity,
            ZERO,
            ZERO,
            ZERO,
            ZERO,
            "",
        )

    # ========================================================
    # 외화계좌간대체 입금
    # ========================================================

    if transaction == "외화계좌간대체입금":
        exchange_rate = get_effective_rate(
            row,
            fx_tables,
        )

        cash_difference = foreign_cash_delta(
            row,
            previous_row,
        )

        warning_message = ""

        if previous_row is None:
            warning_message = (
                "직전 거래가 없어 이전 외화예수금을 "
                "0으로 사용했습니다."
            )

        return (
            ZERO,
            ZERO,
            cash_difference * exchange_rate,
            ZERO,
            ZERO,
            warning_message,
        )

    # ========================================================
    # 배당세금추징 출금
    # ========================================================

    if transaction == "배당세금추징출금":
        return (
            ZERO,
            ZERO,
            ZERO,
            domestic_fee,
            full_tax_sum(row),
            "",
        )

    # ========================================================
    # 무상주 입고
    # ========================================================

    if transaction == "무상주입고":
        return (
            quantity,
            ZERO,
            ZERO,
            domestic_fee,
            full_tax_sum(row),
            "",
        )

    # ========================================================
    # 규칙 미지정 거래
    # ========================================================

    fallback_amount = (
        trade_amount
        if trade_amount != ZERO
        else settlement_amount
    )

    return (
        quantity,
        unit_price,
        fallback_amount,
        domestic_fee,
        full_tax_sum(row),
        f"규칙 미지정 거래종류: {transaction_original}",
    )


# ============================================================
# 시트 헤더 처리
# ============================================================

def build_header_map(
    worksheet,
    header_row_index: int,
) -> Dict[str, int]:
    header_map: Dict[str, int] = {}

    for column_number in range(
        1,
        worksheet.max_column + 1,
    ):
        header_name = clean_text(
            worksheet.cell(
                header_row_index,
                column_number,
            ).value
        )

        if (
            header_name
            and header_name not in header_map
        ):
            header_map[
                header_name
            ] = column_number

    return header_map


# ============================================================
# 원본 시트 읽기
# ============================================================

def read_sheet_rows(
    worksheet,
    fx_tables: Dict[str, ExchangeTable],
) -> Tuple[List[List], List[str]]:
    account_number, holder = extract_account_info(
        worksheet["A1"].value
    )

    header_row_index = 3

    header_map = build_header_map(
        worksheet,
        header_row_index,
    )

    required_headers = [
        "거래일자",
        "거래종류",
        "수량",
        "거래금액",
        "정산금액",
        "외화정산금액",
        "거래세 등",
        "소득세",
        "양도세",
        "통화구분",
        "환율",
        "국외수수료",
        "종목명",
        "단가",
        "수수료",
        "농특세/부가세",
        "지방소득세",
        "외화예수금",
    ]

    for header_name in required_headers:
        if header_name not in header_map:
            raise KeyError(
                f"시트 '{worksheet.title}'에서 "
                f"필수 헤더 '{header_name}'를 "
                "찾지 못했습니다."
            )

    output_rows: List[List] = []
    warnings: List[str] = []

    previous_row: Optional[dict] = None

    for row_number in range(
        4,
        worksheet.max_row + 1,
    ):
        transaction_value = worksheet.cell(
            row_number,
            header_map["거래종류"],
        ).value

        trade_date_value = worksheet.cell(
            row_number,
            header_map["거래일자"],
        ).value

        if (
            transaction_value in (
                None,
                "",
            )
            and trade_date_value in (
                None,
                "",
            )
        ):
            continue

        row_data = {
            header_name: worksheet.cell(
                row_number,
                column_number,
            ).value
            for (
                header_name,
                column_number,
            ) in header_map.items()
        }

        try:
            (
                output_quantity,
                output_unit_price,
                output_trade_amount,
                output_fee,
                output_tax,
                warning_message,
            ) = calculate_row(
                row_data,
                previous_row,
                fx_tables,
            )

        except Exception as error:
            warning_message = (
                f"계산 실패("
                f"{clean_text(row_data.get('거래종류'))}"
                f"): {error}"
            )

            output_quantity = ZERO
            output_unit_price = ZERO
            output_trade_amount = ZERO
            output_fee = ZERO
            output_tax = ZERO

        if warning_message:
            warnings.append(
                f"[{worksheet.title} "
                f"R{row_number}] "
                f"{warning_message}"
            )

        transaction_normalized = normalize_transaction(
            row_data.get("거래종류")
        )

        # 외화매수·외화매도는 종목명에 통화구분을 넣는다.
        if transaction_normalized in {
            "외화매수",
            "외화매도",
        }:
            output_stock_name = clean_text(
                row_data.get("통화구분")
            )

        else:
            output_stock_name = clean_text(
                row_data.get("종목명")
            )

        output_rows.append(
            [
                account_number,
                holder,
                "",
                clean_text(
                    row_data.get("거래종류")
                ),
                output_stock_name,
                parse_date_safe(
                    row_data.get("거래일자")
                ),
                float(output_quantity),
                float(output_unit_price),
                float(output_trade_amount),
                float(output_fee),
                float(output_tax),
            ]
        )

        # 빈 행을 제외한 직전 유효 거래 행
        previous_row = row_data

    return output_rows, warnings


# ============================================================
# 결과 파일 서식
# ============================================================

def autosize_columns(
    worksheet,
) -> None:
    for (
        column_number,
        column_cells,
    ) in enumerate(
        worksheet.iter_cols(
            1,
            worksheet.max_column,
        ),
        start=1,
    ):
        maximum_length = 0

        for cell in column_cells:
            if cell.value is None:
                continue

            maximum_length = max(
                maximum_length,
                len(str(cell.value)),
            )

        worksheet.column_dimensions[
            get_column_letter(
                column_number
            )
        ].width = min(
            max(
                maximum_length + 2,
                10,
            ),
            28,
        )


# ============================================================
# 결과 파일 저장
# ============================================================

def save_output(
    output_path: Path,
    rows: List[List],
    warnings: List[str],
) -> None:
    """
    임시파일에 저장한 뒤 정상적으로 열리는지 확인하고
    최종 결과파일로 교체한다.
    """
    temporary_path = output_path.with_name(
        f".{output_path.stem}_작성중.xlsx"
    )

    if temporary_path.exists():
        temporary_path.unlink()

    workbook = Workbook()

    worksheet = workbook.active
    worksheet.title = "정리"

    # 헤더
    for (
        column_number,
        header_name,
    ) in enumerate(
        OUTPUT_COLUMNS,
        start=1,
    ):
        cell = worksheet.cell(
            1,
            column_number,
        )

        set_plain_text(
            cell,
            header_name,
        )

        cell.font = Font(
            bold=True
        )

    # 데이터
    for (
        row_number,
        row_values,
    ) in enumerate(
        rows,
        start=2,
    ):
        for (
            column_number,
            value,
        ) in enumerate(
            row_values,
            start=1,
        ):
            cell = worksheet.cell(
                row_number,
                column_number,
            )

            if column_number <= 5:
                set_plain_text(
                    cell,
                    value,
                )

            else:
                cell.value = value

    worksheet.freeze_panes = "A2"

    worksheet.auto_filter.ref = (
        worksheet.dimensions
    )

    if worksheet.max_row >= 2:
        for row_cells in worksheet.iter_rows(
            min_row=2,
            max_row=worksheet.max_row,
        ):
            row_cells[
                5
            ].number_format = "yyyy-mm-dd"

            for column_index in (
                6,
                7,
                8,
                9,
                10,
            ):
                row_cells[
                    column_index
                ].number_format = "#,##0.00"

    autosize_columns(
        worksheet
    )

    if warnings:
        warning_sheet = workbook.create_sheet(
            "검토필요"
        )

        header_cell = warning_sheet.cell(
            1,
            1,
        )

        set_plain_text(
            header_cell,
            "메시지",
        )

        header_cell.font = Font(
            bold=True
        )

        for (
            row_number,
            warning_message,
        ) in enumerate(
            warnings,
            start=2,
        ):
            set_plain_text(
                warning_sheet.cell(
                    row_number,
                    1,
                ),
                warning_message,
            )

        warning_sheet.column_dimensions[
            "A"
        ].width = 120

    try:
        workbook.save(
            temporary_path
        )

        workbook.close()

        verification_workbook = load_workbook(
            temporary_path,
            data_only=False,
            read_only=True,
        )

        verification_workbook.close()

        try:
            os.replace(
                temporary_path,
                output_path,
            )

        except PermissionError as error:
            raise PermissionError(
                f"'{output_path.name}' 파일이 "
                "Excel에서 열려 있습니다. "
                "파일을 완전히 닫고 다시 실행해 주세요."
            ) from error

    except Exception:
        workbook.close()

        if temporary_path.exists():
            try:
                temporary_path.unlink()

            except OSError:
                pass

        raise


# ============================================================
# 메인 실행
# ============================================================

def main() -> None:
    input_files = find_input_files(
        BASE_DIR
    )

    if not input_files:
        raise FileNotFoundError(
            f"작업폴더({BASE_DIR})에서 "
            "입력 파일을 찾지 못했습니다. "
            f"예상 패턴: {', '.join(INPUT_PATTERNS)}"
        )

    exchange_directory = find_exchange_dir(
        BASE_DIR
    )

    # exchange_rate 폴더가 없어도
    # 원본 환율이 존재하는 행은 처리할 수 있다.
    fx_tables = load_exchange_rates(
        exchange_directory
    )

    for input_file in input_files:
        input_workbook = load_workbook(
            input_file,
            data_only=True,
        )

        try:
            all_rows: List[List] = []
            all_warnings: List[str] = []

            for sheet_name in (
                input_workbook.sheetnames
            ):
                worksheet = input_workbook[
                    sheet_name
                ]

                rows, warnings = read_sheet_rows(
                    worksheet,
                    fx_tables,
                )

                all_rows.extend(
                    rows
                )

                all_warnings.extend(
                    warnings
                )

        finally:
            input_workbook.close()

        output_path = input_file.with_name(
            f"{input_file.stem}{OUTPUT_SUFFIX}"
        )

        save_output(
            output_path,
            all_rows,
            all_warnings,
        )

        print(
            f"완료: {output_path}"
        )

        if all_warnings:
            print(
                "  - 검토필요 건수: "
                f"{len(all_warnings)}"
            )


if __name__ == "__main__":
    main()