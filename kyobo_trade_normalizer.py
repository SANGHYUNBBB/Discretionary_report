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
    "kyobo26q2*.xlsx",
    "kyobo26q2*.xlsm",
    "KYOBO26q2*.xlsx",
    "KYOBO26q2*.xlsm",
    "교보증권*.xlsx",
    "교보증권*.xlsm",
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

FOREIGN_CODES = {
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

KNOWN_CURRENCY_CODES = FOREIGN_CODES | {"KRW"}

SKIP_TYPES = {
    "해외주식매수입고",
    "해외주식매도출고",
    "타사대체입고신청",
}

ZERO = Decimal("0")
HUNDRED = Decimal("100")


# ============================================================
# 환율 테이블
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


def normalize_transaction(value) -> str:
    """거래명 비교를 위해 공백을 제거한다."""
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
    """문자열을 수식이 아닌 일반 텍스트로 저장한다."""
    cell.value = safe_excel_text(value)
    cell.data_type = "s"


def to_decimal(value) -> Decimal:
    """
    값을 Decimal 숫자로 변환한다.

    숫자 문자열 내부의 공백과 쉼표를 모두 제거한다.

    예:
    "96 620"       -> 96620
    "96,620"       -> 96620
    "96\u00a0620" -> 96620
    """
    if value is None or value == "":
        return ZERO

    if isinstance(value, Decimal):
        result = value

    elif isinstance(value, bool):
        result = Decimal(int(value))

    else:
        try:
            cleaned_value = str(value)

            # 일반 공백, 탭, 줄바꿈, NBSP, 전각 공백 등 제거
            cleaned_value = re.sub(r"\s+", "", cleaned_value)

            # 제로폭 공백, BOM, 쉼표 제거
            cleaned_value = (
                cleaned_value
                .replace("\u200b", "")
                .replace("\ufeff", "")
                .replace(",", "")
            )

            result = Decimal(cleaned_value)

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
    교보증권 원본 파일만 찾는다.
    정리 결과, 잠금 파일, 임시파일은 제외한다.
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
    통화구분이 JPY이면 100엔당 환율을
    1엔당 환율로 바꾸기 위해 100으로 나눈다.
    다른 통화는 그대로 사용한다.
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
    다음 형태의 파일명에서 통화코드를 추출한다.

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

    기본 구조:
    - 파일명에 통화코드 포함
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

        if currency_code not in FOREIGN_CODES:
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


def infer_currency(row: dict) -> str:
    """원본 행에서 통화코드를 찾는다."""
    for key in (
        "통화구분",
        "통화코드",
        "통화",
        "외화구분",
    ):
        value = clean_text(
            row.get(key)
        ).upper()

        if value in KNOWN_CURRENCY_CODES:
            return value

    for key in (
        "적요명",
        "종목명",
        "종목명(거래상대명)",
    ):
        text = clean_text(
            row.get(key)
        ).upper()

        for code in FOREIGN_CODES:
            if code in text:
                return code

    return ""


def get_effective_rate(
    row: dict,
    fx_tables: Dict[str, ExchangeTable],
) -> Decimal:
    """
    exchange_rate 폴더에서 거래일 기준 환율을 조회한다.

    1. 통화구분에 해당하는 환율 파일 선택
    2. 거래일과 같은 날짜 우선
    3. 같은 날짜가 없으면 가장 가까운 이전 날짜 사용
    4. JPY이면 환율을 100으로 나눔
    """
    currency_code = infer_currency(
        row
    )

    if currency_code in (
        "",
        "KRW",
    ):
        return Decimal("1")

    trade_date = parse_date_safe(
        row.get("거래일자")
        or row.get("일자")
        or row.get("거래일")
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
# 원본 값 추출
# ============================================================

def first_decimal(
    row: dict,
    keys: Tuple[str, ...],
) -> Decimal:
    for key in keys:
        if key in row:
            return to_decimal(
                row.get(key)
            )

    return ZERO


def quantity_value(row: dict) -> Decimal:
    return first_decimal(
        row,
        (
            "수량",
            "거래수량",
        ),
    )


def unit_price_value(row: dict) -> Decimal:
    return first_decimal(
        row,
        (
            "단가",
            "거래단가",
        ),
    )


def trade_amount_value(row: dict) -> Decimal:
    return first_decimal(
        row,
        (
            "거래금액",
            "금액",
        ),
    )


def fee_value(row: dict) -> Decimal:
    return first_decimal(
        row,
        (
            "수수료",
            "수수료금액",
        ),
    )


def tax_value(row: dict) -> Decimal:
    return first_decimal(
        row,
        (
            "제세금",
            "세금",
            "거래세",
            "세금합계",
        ),
    )


def stock_name_value(row: dict) -> str:
    for key in (
        "종목명(거래상대명)",
        "종목명",
        "거래상대명",
    ):
        if key in row:
            return clean_text(
                row.get(key)
            )

    return ""


# ============================================================
# 계좌정보 및 헤더 검색
# ============================================================

def extract_account_info_from_a1(
    worksheet,
) -> Tuple[str, str]:
    text = clean_text(
        worksheet["A1"].value
    )

    account_match = re.search(
        r"(\d{4}-\d{5}-\d{2})",
        text,
    )

    if account_match:
        account_number = account_match.group(1)
        holder = text[
            account_match.end():
        ].strip()

        return account_number, holder

    fallback_number = re.search(
        r"(\d+(?:-\d+)+|\d{8,})",
        text,
    )

    account_number = (
        fallback_number.group(1)
        if fallback_number
        else ""
    )

    if fallback_number:
        holder = text[
            fallback_number.end():
        ].strip()

    else:
        name_match = re.search(
            r"([가-힣A-Za-z]+)\s*$",
            text,
        )

        holder = (
            name_match.group(1)
            if name_match
            else ""
        )

    return account_number, holder


def find_header_row_and_map(
    worksheet,
) -> Tuple[int, Dict[str, int]]:
    for row_number in range(
        1,
        min(worksheet.max_row, 15) + 1,
    ):
        headers = [
            clean_text(
                worksheet.cell(
                    row_number,
                    column_number,
                ).value
            )
            for column_number in range(
                1,
                worksheet.max_column + 1,
            )
        ]

        if "적요명" in headers:
            header_map = {
                header_name: column_number
                for column_number, header_name
                in enumerate(
                    headers,
                    start=1,
                )
                if header_name
            }

            return row_number, header_map

    raise KeyError(
        f"시트 '{worksheet.title}'에서 "
        "'적요명' 헤더를 찾지 못했습니다."
    )


# ============================================================
# 거래종류별 계산
# ============================================================

def calculate_row(
    row: dict,
    fx_tables: Dict[str, ExchangeTable],
) -> Tuple[
    Optional[List[Decimal]],
    Optional[str],
]:
    transaction_original = clean_text(
        row.get("적요명")
    )

    transaction = normalize_transaction(
        transaction_original
    )

    if transaction in SKIP_TYPES:
        return None, None

    quantity = quantity_value(row)
    unit_price = unit_price_value(row)
    trade_amount = trade_amount_value(row)
    fee = fee_value(row)
    tax = tax_value(row)
    currency_code = infer_currency(row)

    # 해외주식매도대금입금 / 해외주식매수대금출금
    if transaction in {
        "해외주식매도대금입금",
        "해외주식매수대금출금",
    }:
        exchange_rate = get_effective_rate(
            row,
            fx_tables,
        )

        output_unit_price = (
            unit_price
            * exchange_rate
        )

        return [
            quantity,
            output_unit_price,
            quantity * output_unit_price,
            fee * exchange_rate,
            tax * exchange_rate,
        ], ""

    # 배당금입금 - 통화구분이 외화인 경우
    if (
        transaction == "배당금입금"
        and currency_code not in (
            "",
            "KRW",
        )
    ):
        exchange_rate = get_effective_rate(
            row,
            fx_tables,
        )

        return [
            ZERO,
            ZERO,
            trade_amount * exchange_rate,
            fee * exchange_rate,
            tax * exchange_rate,
        ], ""

    # 타사대체입고
    if transaction == "타사대체입고":
        exchange_rate = get_effective_rate(
            row,
            fx_tables,
        )

        output_unit_price = (
            unit_price
            * exchange_rate
        )

        return [
            quantity,
            output_unit_price,
            quantity * output_unit_price,
            fee * exchange_rate,
            tax * exchange_rate,
        ], ""

    # 정기예탁금이용료입금
    if transaction == "정기예탁금이용료입금":
        return [
            ZERO,
            ZERO,
            trade_amount,
            fee,
            tax,
        ], ""

    # 은행이체송금 / 은행이체입금
    if transaction in {
        "은행이체송금",
        "은행이체입금",
    }:
        return [
            ZERO,
            ZERO,
            trade_amount,
            fee,
            tax,
        ], ""

    # 원화입금·출금 환전
    if transaction in {
        "원화입금(환전)",
        "원화출금(환전)",
        "원화출금(증거금환전)",
        "원화입금(증거금환전)",
    }:
        return [
            ZERO,
            ZERO,
            trade_amount,
            fee,
            tax,
        ], ""

    # 외화배당세금환급입금
    if transaction == "외화배당세금환급입금":
        exchange_rate = get_effective_rate(
            row,
            fx_tables,
        )

        return [
            ZERO,
            ZERO,
            trade_amount * exchange_rate,
            fee * exchange_rate,
            tax * exchange_rate,
        ], ""

    # 배당소득세출금
    if transaction == "배당소득세출금":
        return [
            ZERO,
            ZERO,
            ZERO,
            fee,
            tax,
        ], ""

    # 권리정정출금(외화)
    if transaction == "권리정정출금(외화)":
        exchange_rate = get_effective_rate(
            row,
            fx_tables,
        )

        return [
            ZERO,
            ZERO,
            trade_amount * exchange_rate,
            fee * exchange_rate,
            tax * exchange_rate,
        ], ""

    # 외화배당금입금
    if transaction == "외화배당금입금":
        exchange_rate = get_effective_rate(
            row,
            fx_tables,
        )

        return [
            quantity,
            unit_price,
            trade_amount * exchange_rate,
            fee * exchange_rate,
            tax * exchange_rate,
        ], ""

    # 계좌대체입고
    if transaction == "계좌대체입고":
        if currency_code == "":
            return [
                quantity,
                unit_price,
                ZERO,
                fee,
                tax,
            ], ""

        exchange_rate = get_effective_rate(
            row,
            fx_tables,
        )

        return [
            quantity,
            unit_price * exchange_rate,
            ZERO,
            fee * exchange_rate,
            tax * exchange_rate,
        ], ""

    # 계좌간약정대체입금
    if transaction == "계좌간약정대체입금":
        return [
            quantity,
            unit_price,
            trade_amount,
            fee,
            tax,
        ], ""

    # 마감후원화출금(환전)
    if transaction == "마감후원화출금(환전)":
        return [
            quantity,
            unit_price,
            trade_amount,
            fee,
            tax,
        ], ""

    # 마감후외화입금(환전)
    if transaction == "마감후외화입금(환전)":
        exchange_rate = get_effective_rate(
            row,
            fx_tables,
        )

        return [
            quantity,
            unit_price,
            trade_amount * exchange_rate,
            fee * exchange_rate,
            tax * exchange_rate,
        ], ""

    # 환전정산분입금
    if transaction == "환전정산분입금":
        return [
            quantity,
            unit_price,
            trade_amount,
            fee,
            tax,
        ], ""

    # 외화출금(환전)
    if transaction == "외화출금(환전)":
        exchange_rate = get_effective_rate(
            row,
            fx_tables,
        )

        return [
            quantity,
            unit_price,
            trade_amount * exchange_rate,
            fee * exchange_rate,
            tax * exchange_rate,
        ], ""

    # 규칙 미지정
    return [
        quantity,
        unit_price,
        trade_amount,
        fee,
        tax,
    ], (
        f"규칙 미지정 적요명: "
        f"{transaction_original}"
    )


# ============================================================
# 원본 시트 읽기
# ============================================================

def read_sheet_rows(
    worksheet,
    fx_tables: Dict[str, ExchangeTable],
) -> Tuple[List[List], List[str]]:
    account_number, holder = (
        extract_account_info_from_a1(
            worksheet
        )
    )

    (
        header_row_index,
        header_map,
    ) = find_header_row_and_map(
        worksheet
    )

    required_headers = [
        "적요명",
        "거래일자",
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
        header_row_index + 1,
        worksheet.max_row + 1,
    ):
        transaction_value = worksheet.cell(
            row_number,
            header_map["적요명"],
        ).value

        trade_date_value = worksheet.cell(
            row_number,
            header_map["거래일자"],
        ).value

        if (
            clean_text(transaction_value) == ""
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

        transaction_name = clean_text(
            row_data.get("적요명")
        )

        try:
            calculation, warning_message = (
                calculate_row(
                    row_data,
                    fx_tables,
                )
            )

        except Exception as error:
            calculation = [
                ZERO,
                ZERO,
                ZERO,
                ZERO,
                ZERO,
            ]

            warning_message = (
                f"계산 실패({transaction_name}): "
                f"{error}"
            )

        if calculation is None:
            continue

        if warning_message:
            warnings.append(
                f"[{worksheet.title} "
                f"R{row_number}] "
                f"{warning_message}"
            )

        (
            output_quantity,
            output_unit_price,
            output_trade_amount,
            output_fee,
            output_tax,
        ) = calculation

        output_rows.append(
            [
                account_number,
                holder,
                "",
                transaction_name,
                stock_name_value(
                    row_data
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
    임시파일에 먼저 저장하고 정상적으로 열리는지 확인한 뒤
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
