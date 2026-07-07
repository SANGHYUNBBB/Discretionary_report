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
    "Samsung26q2*.xlsx",
    "Samsung26q2*.xlsm",
    "삼성거래내역*.xlsx",
    "삼성거래내역*.xlsm",
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

DOMESTIC_STOCK_TYPES = {
    "매수",
    "매도",
    "매수_NXT",
    "매도_NXT",
}

CASH_TRANSFER_TYPES = {
    "이체입금",
    "이체출금",
    "대체입금",
    "대체출금",
    "출금",
}

SUPPORT_BONUS_TYPES = {
    "투자지원금",
    "투자지원금 입금",
}


# ============================================================
# 환율 데이터 구조
# ============================================================

@dataclass
class ExchangeTable:
    dates_ord: List[int]
    rates: List[Decimal]

    def lookup(self, target_date: date) -> Decimal:
        """
        거래일과 같은 날짜의 환율을 우선 사용한다.
        같은 날짜가 없으면 가장 가까운 이전 날짜의 환율을 사용한다.
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


def safe_excel_text(value) -> str:
    return ILLEGAL_CHARACTERS_RE.sub(
        "",
        clean_text(value),
    )


def set_plain_text(cell, value) -> None:
    """
    문자열을 수식이 아닌 일반 텍스트로 저장한다.
    """
    cell.value = safe_excel_text(value)
    cell.data_type = "s"


def to_decimal(value) -> Decimal:
    """
    공란, 비정상 숫자, NaN, 무한대는 0으로 처리한다.
    """
    if value is None or value == "":
        return Decimal("0")

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
            return Decimal("0")

    if not result.is_finite():
        return Decimal("0")

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
        "%Y-%m-%d",
        "%Y.%m.%d",
        "%Y/%m/%d",
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

    else:
        number_match = re.search(
            r"(\d{8,})",
            text,
        )

        account_number = (
            number_match.group(1)
            if number_match
            else ""
        )

    holder_match = re.search(
        r"\]\s*(.+)$",
        text,
    )

    if holder_match:
        holder = holder_match.group(1).strip()

    elif account_match:
        holder = text[
            account_match.end():
        ].strip()

    else:
        holder = ""

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
    100으로 나누어 엔당 환율로 변환한다.
    """
    currency_code = clean_text(
        currency_code
    ).upper()

    if currency_code == "JPY":
        return rate / Decimal("100")

    return rate


def extract_currency_code_from_filename(
    file_path: Path,
) -> str:
    """
    다음과 같은 파일명에서 통화코드를 찾는다.

    USD.xlsx
    USD_환율.xlsx
    환율_USD.xlsx
    """
    stem_upper = (
        file_path.stem.strip().upper()
    )

    currency_match = re.search(
        r"(?<![A-Z])([A-Z]{3})(?![A-Z])",
        stem_upper,
    )

    if currency_match:
        return currency_match.group(1)

    return stem_upper


def load_exchange_rates(
    exchange_dir: Optional[Path],
) -> Dict[str, ExchangeTable]:
    """
    exchange_rate 폴더의 환율 파일을 읽는다.

    - 파일명: 통화코드 포함
    - 첫 번째 시트
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

        if (
            resolved_path in seen_paths
            or file_path.name.startswith("~$")
        ):
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
                    or exchange_rate == Decimal("0")
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


def get_fx_rate(
    row: dict,
    fx_tables: Dict[str, ExchangeTable],
) -> Decimal:
    """
    삼성증권 원본에는 환율 열이 없으므로
    환율은 모두 exchange_rate에서 조회한다.

    1. 통화코드에 해당하는 환율 파일 선택
    2. 거래일과 같은 날짜 사용
    3. 같은 날짜가 없으면 가장 가까운 이전 날짜 사용
    4. JPY는 100으로 나누어 엔당 환율로 변환
    """
    currency_code = clean_text(
        row.get("통화코드")
    ).upper()

    if currency_code in (
        "",
        "KRW",
    ):
        return Decimal("1")

    trade_date = parse_date_safe(
        row.get("거래일자")
    )

    if trade_date is None:
        raise ValueError(
            "환율 조회에 필요한 거래일자를 "
            "해석할 수 없습니다."
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
            f"통화코드 {currency_code}에 해당하는 "
            "환율 파일을 찾지 못했습니다. "
            f"사용 가능한 통화코드: {available_codes}"
        )

    raw_rate = exchange_table.lookup(
        trade_date
    )

    return adjust_fx_rate(
        currency_code,
        raw_rate,
    )


# ============================================================
# 거래명 판별
# ============================================================

def is_foreign_stock_transaction(
    transaction_name: str,
    currency_code: str,
) -> bool:
    """
    미국(NASDAQ)주식매수, 미국(NYSE)주식매도,
    일본(동경)주식매수, 홍콩주식매도 등을 판별한다.
    """
    currency_code = clean_text(
        currency_code
    ).upper()

    if currency_code in (
        "",
        "KRW",
    ):
        return False

    return transaction_name.endswith(
        (
            "주식매수",
            "주식매도",
        )
    )


# ============================================================
# 거래종류별 계산
# ============================================================

def calculate_row(
    row: dict,
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
    반환값:

    수량,
    매매단가,
    매매금액,
    위탁매매수수료,
    각종세금,
    검토메시지
    """
    transaction_name = clean_text(
        row.get("거래명")
    )

    currency_code = clean_text(
        row.get("통화코드")
    ).upper()

    quantity = to_decimal(
        row.get("거래수량")
    )

    unit_price = to_decimal(
        row.get("거래단가")
    )

    trade_amount = to_decimal(
        row.get("거래금액")
    )

    foreign_trade_amount = to_decimal(
        row.get("외화거래금액")
    )

    fee = to_decimal(
        row.get("수수료/Fee")
    )

    tax_fee = to_decimal(
        row.get("제세금/대출이자")
    )

    foreign_fee = to_decimal(
        row.get("외화수수료")
    )

    zero = Decimal("0")

    # ========================================================
    # 외화매수 / 외화매도
    # ========================================================
    #
    # 수량: 거래수량
    # 매매단가: 거래단가
    # 매매금액: 수량 × 매매단가
    #
    # JPY:
    # 매매금액 = 수량 × 매매단가 ÷ 100
    # ========================================================

    if transaction_name in {
        "외화매수",
        "외화매도",
    }:
        output_amount = (
            quantity
            * unit_price
        )

        if currency_code == "JPY":
            output_amount = (
                output_amount
                / Decimal("100")
            )

        return (
            quantity,
            unit_price,
            output_amount,
            zero,
            zero,
            "",
        )

    # ========================================================
    # 외국주식 매수 / 매도
    # ========================================================

    if is_foreign_stock_transaction(
        transaction_name,
        currency_code,
    ):
        exchange_rate = get_fx_rate(
            row,
            fx_tables,
        )

        converted_unit_price = (
            unit_price
            * exchange_rate
        )

        return (
            quantity,
            converted_unit_price,
            quantity * converted_unit_price,
            foreign_fee * exchange_rate,
            tax_fee * exchange_rate,
            "",
        )

    # ========================================================
    # 국내주식 매수 / 매도
    # ========================================================

    if transaction_name in DOMESTIC_STOCK_TYPES:
        return (
            quantity,
            unit_price,
            quantity * unit_price,
            fee,
            tax_fee,
            "",
        )

    # ========================================================
    # 세금출금(해외)
    # ========================================================
    #
    # 수량: 거래수량
    # 매매단가: 거래단가
    # 매매금액: 0
    # 각종세금: 수량 × 거래단가
    #
    # JPY:
    # 각종세금 = 수량 × 거래단가 ÷ 100
    # ========================================================

    if transaction_name == "세금출금(해외)":
        output_tax = (
            quantity
            * unit_price
        )

        if currency_code == "JPY":
            output_tax = (
                output_tax
                / Decimal("100")
            )

        return (
            quantity,
            unit_price,
            zero,
            zero,
            output_tax,
            "",
        )

    # ========================================================
    # 배당금입금
    # ========================================================

    if transaction_name == "배당금입금":
        if currency_code in (
            "",
            "KRW",
        ):
            return (
                zero,
                zero,
                trade_amount,
                zero,
                tax_fee,
                "",
            )

        exchange_rate = get_fx_rate(
            row,
            fx_tables,
        )

        return (
            quantity,
            exchange_rate,
            quantity * exchange_rate,
            zero,
            tax_fee,
            "",
        )

    # ========================================================
    # 이체입금 / 이체출금 / 대체입금 / 대체출금 / 출금
    # ========================================================

    if transaction_name in CASH_TRANSFER_TYPES:
        return (
            zero,
            zero,
            trade_amount,
            zero,
            zero,
            "",
        )

    # ========================================================
    # 투자지원금
    # ========================================================

    if transaction_name in SUPPORT_BONUS_TYPES:
        return (
            quantity,
            unit_price,
            quantity * unit_price,
            zero,
            zero,
            "",
        )

    # ========================================================
    # 이용료입금
    # ========================================================

    if transaction_name == "이용료입금":
        return (
            zero,
            zero,
            trade_amount,
            zero,
            fee + tax_fee,
            "",
        )

    # ========================================================
    # 배당입고
    # ========================================================

    if transaction_name == "배당입고":
        exchange_rate = get_fx_rate(
            row,
            fx_tables,
        )

        converted_unit_price = (
            unit_price
            * exchange_rate
        )

        return (
            quantity,
            converted_unit_price,
            quantity * converted_unit_price,
            zero,
            zero,
            "",
        )

    # ========================================================
    # 수수료입금
    # ========================================================

    if transaction_name == "수수료입금":
        return (
            zero,
            zero,
            zero,
            trade_amount,
            zero,
            "",
        )

    # ========================================================
    # 자문사수수료출금
    # ========================================================

    if transaction_name == "자문사수수료출금":
        return (
            zero,
            zero,
            trade_amount,
            zero,
            zero,
            "",
        )

    # ========================================================
    # 외화이체입금 / 외화이체출금
    # ========================================================

    if transaction_name in {
        "외화이체입금",
        "외화이체출금",
    }:
        exchange_rate = get_fx_rate(
            row,
            fx_tables,
        )

        return (
            quantity,
            exchange_rate,
            quantity * exchange_rate,
            zero,
            zero,
            "",
        )

    # ========================================================
    # 대차출고
    # ========================================================

    if transaction_name == "대차출고":
        return (
            quantity,
            zero,
            trade_amount,
            zero,
            fee + tax_fee,
            "",
        )

    # ========================================================
    # 외화배당세금환급
    # ========================================================

    if transaction_name == "외화배당세금환급":
        exchange_rate = get_fx_rate(
            row,
            fx_tables,
        )

        return (
            quantity,
            exchange_rate,
            quantity * exchange_rate,
            zero,
            fee,
            "",
        )

    # ========================================================
    # 세금출금(국내)
    # ========================================================

    if transaction_name == "세금출금(국내)":
        return (
            zero,
            zero,
            zero,
            zero,
            tax_fee,
            "",
        )

    # ========================================================
    # 이자입금
    # ========================================================

    if transaction_name == "이자입금":
        return (
            zero,
            zero,
            trade_amount,
            fee,
            tax_fee,
            "",
        )

    # ========================================================
    # 타사출고
    # ========================================================

    if transaction_name == "타사출고":
        return (
            quantity,
            unit_price,
            trade_amount,
            fee,
            tax_fee,
            "",
        )

    # ========================================================
    # 청약 / 청약입고
    # ========================================================

    if transaction_name in {
        "청약",
        "청약입고",
    }:
        return (
            quantity,
            unit_price,
            quantity * unit_price,
            fee,
            tax_fee,
            "",
        )

    # ========================================================
    # 오픈이체출금
    # ========================================================

    if transaction_name == "오픈이체출금":
        return (
            zero,
            zero,
            trade_amount,
            fee,
            tax_fee,
            "",
        )

    # ========================================================
    # 외화대체입금 / 외화대체출금
    # ========================================================

    if transaction_name in {
        "외화대체입금",
        "외화대체출금",
    }:
        return (
            quantity,
            unit_price,
            foreign_trade_amount,
            zero,
            zero,
            "",
        )

    # ========================================================
    # 규칙 미지정
    # ========================================================

    warning_message = (
        f"규칙 미지정 거래명: {transaction_name}"
    )

    return (
        quantity,
        unit_price,
        trade_amount,
        fee,
        tax_fee,
        warning_message,
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
    account_number, holder = (
        extract_account_info(
            worksheet["A1"].value
        )
    )

    header_row_index = 3

    header_map = build_header_map(
        worksheet,
        header_row_index,
    )

    required_headers = [
        "거래일자",
        "거래명",
        "종목명",
        "거래수량",
        "거래단가",
        "거래금액",
        "외화거래금액",
        "제세금/대출이자",
        "수수료/Fee",
        "통화코드",
        "외화수수료",
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

    for row_number in range(
        4,
        worksheet.max_row + 1,
    ):
        transaction_name = worksheet.cell(
            row_number,
            header_map["거래명"],
        ).value

        trade_date = worksheet.cell(
            row_number,
            header_map["거래일자"],
        ).value

        if (
            transaction_name in (
                None,
                "",
            )
            and trade_date in (
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
                fx_tables,
            )

        except Exception as error:
            warning_message = (
                f"계산 실패("
                f"{clean_text(row_data.get('거래명'))}"
                f"): {error}"
            )

            output_quantity = to_decimal(
                row_data.get("거래수량")
            )

            output_unit_price = to_decimal(
                row_data.get("거래단가")
            )

            output_trade_amount = to_decimal(
                row_data.get("거래금액")
            )

            output_fee = to_decimal(
                row_data.get("수수료/Fee")
            )

            output_tax = to_decimal(
                row_data.get("제세금/대출이자")
            )

        if warning_message:
            warnings.append(
                f"[{worksheet.title} "
                f"R{row_number}] "
                f"{warning_message}"
            )

        output_rows.append(
            [
                account_number,
                holder,
                "",
                clean_text(
                    row_data.get("거래명")
                ),
                clean_text(
                    row_data.get("종목명")
                ),
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
    임시파일로 저장하고 검증한 뒤
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

    for column_number, header_name in enumerate(
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

    for row_number, row_values in enumerate(
        rows,
        start=2,
    ):
        for column_number, value in enumerate(
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

        for row_number, warning_message in enumerate(
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

    if exchange_directory is None:
        raise FileNotFoundError(
            f"작업폴더({BASE_DIR}) 안에서 "
            "exchange_rate 폴더를 찾지 못했습니다. "
            f"허용 폴더명: "
            f"{', '.join(EXCHANGE_DIR_NAMES)}"
        )

    fx_tables = load_exchange_rates(
        exchange_directory
    )

    if not fx_tables:
        raise ValueError(
            "exchange_rate 폴더에서 사용할 수 있는 "
            "환율 파일을 찾지 못했습니다."
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