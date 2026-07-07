from __future__ import annotations

import bisect
import os
import re
from dataclasses import dataclass
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Dict, List, Optional, Tuple
from zipfile import ZipFile

from openpyxl import Workbook, load_workbook
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


# ============================================================
# 기본 설정
# ============================================================

BASE_DIR = Path.cwd()

INPUT_PATTERNS = [
    "han26q2.xlsx",
    "HAN26q2.xlsx",
    "한국투자증권*.xlsx",
    "한국투자증권*.xlsm",
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
}


# ============================================================
# 환율 테이블
# ============================================================

@dataclass
class ExchangeTable:
    dates_ord: List[int]
    rates: List[Decimal]

    def lookup(self, target_date: date) -> Decimal:
        """
        거래일과 같거나 거래일보다 이전인 날짜 중
        가장 최근 환율을 반환한다.
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
    """
    엑셀 XML에서 사용할 수 없는 제어문자를 제거한다.
    """
    text = clean_text(value)

    return ILLEGAL_CHARACTERS_RE.sub(
        "",
        text,
    )


def set_plain_text(cell, value) -> None:
    """
    셀 값을 수식이 아닌 일반 문자열로 강제 저장한다.

    값이 =, +, -, @ 등으로 시작하더라도
    수식으로 저장되지 않는다.
    """
    cell.value = safe_excel_text(value)
    cell.data_type = "s"


def to_decimal(value) -> Decimal:
    """
    엑셀 값을 Decimal 숫자로 변환한다.

    공란, None, 숫자가 아닌 값은 0으로 처리한다.
    NaN과 무한대도 0으로 처리한다.
    """
    if value is None or value == "":
        return Decimal("0")

    if isinstance(value, Decimal):
        result = value

    elif isinstance(value, bool):
        result = Decimal(int(value))

    else:
        try:
            text = str(value).replace(",", "").strip()
            result = Decimal(text)

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
    """
    엑셀 날짜 또는 문자열 날짜를 date 형식으로 변환한다.
    """
    if value is None or value == "":
        return None

    if isinstance(value, datetime):
        return value.date()

    if isinstance(value, date):
        return value

    text = clean_text(value)

    date_formats = [
        "%Y/%m/%d",
        "%Y-%m-%d",
        "%Y.%m.%d",
        "%Y%m%d",
    ]

    for date_format in date_formats:
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
    """
    작업 폴더 안에서 exchange_rate 폴더를 찾는다.
    """
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
    작업 대상 원본 엑셀 파일을 찾는다.

    이미 만들어진 정리 파일, 임시파일,
    Excel 잠금 파일은 제외한다.
    """
    files: List[Path] = []

    for pattern in INPUT_PATTERNS:
        files.extend(
            base_dir.glob(pattern)
        )

    result = {
        file_path.resolve()
        for file_path in files
        if not file_path.stem.endswith("_정리")
        and not file_path.name.startswith(".")
        and not file_path.name.startswith("~$")
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
    JPY 환율은 100엔 기준이므로 100으로 나눈다.

    예:
    환율 910 → 실제 적용 환율 9.10
    """
    currency_code = clean_text(
        currency_code
    ).upper()

    if currency_code == "JPY":
        return rate / Decimal("100")

    return rate


def load_exchange_rates(
    exchange_dir: Optional[Path],
) -> Dict[str, ExchangeTable]:
    """
    exchange_rate 폴더의 환율 파일을 읽는다.

    환율 파일 형식:
    - 파일명: USD.xlsx, JPY.xlsx 등
    - 첫 번째 시트 사용
    - A열: 날짜
    - C열: 환율
    - 10행부터 데이터
    """
    rate_map: Dict[str, ExchangeTable] = {}

    # 원본 환율열이 있으면 처리할 수 있으므로
    # exchange_rate 폴더가 없어도 바로 종료하지 않는다.
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

        seen_paths.add(resolved_path)
        unique_files.append(file_path)

    for file_path in unique_files:
        currency_code = (
            file_path.stem.strip().upper()
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

            dates_ord: List[int] = []
            rates: List[Decimal] = []

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

                if exchange_date is None:
                    continue

                if exchange_rate == Decimal("0"):
                    continue

                dates_ord.append(
                    exchange_date.toordinal()
                )

                rates.append(
                    exchange_rate
                )

        finally:
            workbook.close()

        if not dates_ord:
            raise ValueError(
                "환율 파일에서 날짜와 환율을 "
                f"찾지 못했습니다: {file_path}"
            )

        paired_data = sorted(
            zip(
                dates_ord,
                rates,
            ),
            key=lambda item: item[0],
        )

        rate_map[
            currency_code
        ] = ExchangeTable(
            dates_ord=[
                item[0]
                for item in paired_data
            ],
            rates=[
                item[1]
                for item in paired_data
            ],
        )

    return rate_map


def split_tx_and_currency(
    transaction_text: str,
) -> Tuple[str, str]:
    """
    거래종류의 마지막 3글자로 통화코드를 판별한다.

    예:
    해외증권매수USD
    → 해외증권매수, USD

    해외증권매도JPY
    → 해외증권매도, JPY
    """
    transaction_text = clean_text(
        transaction_text
    )

    if len(transaction_text) < 3:
        return transaction_text, ""

    suffix = (
        transaction_text[-3:].upper()
    )

    if suffix in CURRENCY_CODES:
        base_transaction = (
            transaction_text[:-3].strip()
        )

        return (
            base_transaction,
            suffix,
        )

    return transaction_text, ""


def get_effective_rate(
    row: dict,
    fx_tables: Dict[str, ExchangeTable],
) -> Decimal:
    """
    환율 적용 우선순위:

    1순위
    원본 환율열 값이 있고 0이 아니면 원본 환율 적용

    2순위
    원본 환율이 공란 또는 0이면 exchange_rate 폴더 적용

    추가 규칙
    - 거래종류 마지막 3글자로 통화코드 확인
    - JPY는 환율을 100으로 나누어 적용
    - 통화코드가 없으면 원화 거래로 보고 환율 1 적용
    """
    currency_code = clean_text(
        row.get("통화코드")
    ).upper()

    if currency_code in (
        "",
        "KRW",
    ):
        return Decimal("1")

    # --------------------------------------------------------
    # 1순위: 원본 환율열
    # --------------------------------------------------------

    original_rate = to_decimal(
        row.get("환율")
    )

    if original_rate != Decimal("0"):
        return adjust_fx_rate(
            currency_code,
            original_rate,
        )

    # --------------------------------------------------------
    # 2순위: exchange_rate 폴더
    # --------------------------------------------------------

    trade_date = parse_date_safe(
        row.get("거래일")
    )

    if trade_date is None:
        raise ValueError(
            "원본 환율이 공란 또는 0이고 "
            "거래일도 해석할 수 없습니다."
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
            f"원본 환율이 공란 또는 0이며 "
            f"{currency_code} 환율 파일도 없습니다. "
            f"사용 가능한 환율: {available_codes}"
        )

    folder_rate = exchange_table.lookup(
        trade_date
    )

    return adjust_fx_rate(
        currency_code,
        folder_rate,
    )


# ============================================================
# 세금 계산
# ============================================================

def tax_sum_raw(row: dict) -> Decimal:
    """
    각종세금 = 거래세 + 세금 + 부가세
    """
    transaction_tax = to_decimal(
        row.get("거래세")
    )

    tax = to_decimal(
        row.get("세금")
    )

    additional_tax = to_decimal(
        row.get("부가세")
    )

    return (
        transaction_tax
        + tax
        + additional_tax
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
    transaction_type = clean_text(
        row.get("구분기준거래종류")
    )

    transaction_key = (
        transaction_type.casefold()
    )

    quantity = to_decimal(
        row.get("거래수량")
    )

    unit_price = to_decimal(
        row.get("거래단가")
    )

    trade_amount = to_decimal(
        row.get("거래금액")
    )

    fee = to_decimal(
        row.get("수수료")
    )

    raw_tax = tax_sum_raw(row)

    zero = Decimal("0")

    # ========================================================
    # 해외증권매도 / 해외증권매수
    # ========================================================

    if transaction_type in {
        "해외증권매도",
        "해외증권매수",
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
            fee * exchange_rate,
            raw_tax * exchange_rate,
            "",
        )

    # ========================================================
    # 해외증권배당금입금
    # ========================================================

    if transaction_type == "해외증권배당금입금":
        exchange_rate = get_effective_rate(
            row,
            fx_tables,
        )

        return (
            zero,
            zero,
            trade_amount * exchange_rate,
            fee * exchange_rate,
            raw_tax * exchange_rate,
            "",
        )

    # ========================================================
    # 해외주식배당원화세금
    # ========================================================

    if transaction_type == "해외주식배당원화세금":
        return (
            zero,
            zero,
            zero,
            fee,
            raw_tax,
            "",
        )

    # ========================================================
    # 외화매수 / 외화매도 / 자동환전(외화매도)
    # ========================================================

    if transaction_type in {
        "외화매수",
        "외화매도",
        "자동환전(외화매도)",
    }:
        exchange_rate = get_effective_rate(
            row,
            fx_tables,
        )

        return (
            trade_amount,
            exchange_rate,
            trade_amount * exchange_rate,
            fee * exchange_rate,
            raw_tax * exchange_rate,
            "",
        )

    # ========================================================
    # 외화예탁금이용료입금
    # ========================================================

    if transaction_type == "외화예탁금이용료입금":
        exchange_rate = get_effective_rate(
            row,
            fx_tables,
        )

        return (
            trade_amount,
            exchange_rate,
            trade_amount * exchange_rate,
            fee * exchange_rate,
            raw_tax * exchange_rate,
            "",
        )

    # ========================================================
    # 외화예탁금이용료원화세금
    # ========================================================

    if transaction_type == "외화예탁금이용료원화세금":
        return (
            zero,
            zero,
            zero,
            fee,
            raw_tax,
            "",
        )

    # ========================================================
    # 예탁금이용료
    # ========================================================

    if transaction_type == "예탁금이용료":
        return (
            zero,
            zero,
            trade_amount,
            fee,
            raw_tax,
            "",
        )

    # ========================================================
    # 외화실시간직접매도환전
    # HTS자문사외화실시간직접매도환전
    # ========================================================

    if transaction_type in {
        "외화실시간직접매도환전",
        "HTS자문사외화실시간직접매도환전",
    }:
        exchange_rate = get_effective_rate(
            row,
            fx_tables,
        )

        return (
            trade_amount,
            exchange_rate,
            trade_amount * exchange_rate,
            fee * exchange_rate,
            raw_tax * exchange_rate,
            "",
        )

    # ========================================================
    # 외화당사이체입금
    # ========================================================

    if transaction_type == "외화당사이체입금":
        exchange_rate = get_effective_rate(
            row,
            fx_tables,
        )

        return (
            trade_amount,
            exchange_rate,
            trade_amount * exchange_rate,
            fee * exchange_rate,
            raw_tax * exchange_rate,
            "",
        )

    # ========================================================
    # 랩대체계약출금
    # 당사이체출금
    # smart+당사이체출금
    # smart+당사이체입금
    # ========================================================

    if transaction_key in {
        "랩대체계약출금".casefold(),
        "당사이체출금".casefold(),
        "smart+당사이체출금".casefold(),
        "smart+당사이체입금".casefold(),
    }:
        return (
            zero,
            zero,
            trade_amount,
            fee,
            raw_tax,
            "",
        )

    # ========================================================
    # 현지세금환급
    # ========================================================

    if transaction_type == "현지세금환급":
        exchange_rate = get_effective_rate(
            row,
            fx_tables,
        )

        return (
            zero,
            zero,
            trade_amount * exchange_rate,
            zero,
            zero,
            "",
        )

    # ========================================================
    # 현지세금재징수
    # ========================================================

    if transaction_type == "현지세금재징수":
        exchange_rate = get_effective_rate(
            row,
            fx_tables,
        )

        return (
            zero,
            zero,
            zero,
            zero,
            trade_amount * exchange_rate,
            "",
        )

    # ========================================================
    # wrap해지입고
    # ========================================================

    if transaction_key == "wrap해지입고".casefold():
        return (
            quantity,
            unit_price,
            zero,
            zero,
            zero,
            "",
        )

    # ========================================================
    # wrap해지입금
    # ========================================================
    #
    # 수량: 거래수량
    # 매매단가: 거래단가
    # 매매금액: 거래금액
    # 위탁매매수수료: 수수료
    # 각종세금: 거래세 + 세금 + 부가세
    # ========================================================

    if transaction_key == "wrap해지입금".casefold():
        return (
            quantity,
            unit_price,
            trade_amount,
            fee,
            raw_tax,
            "",
        )

    # ========================================================
    # 거래소주식매도 / 코스닥주식매도
    # ========================================================

    if transaction_type in {
        "거래소주식매도",
        "코스닥주식매도",
    }:
        return (
            quantity,
            unit_price,
            trade_amount,
            fee,
            raw_tax,
            "",
        )

    # ========================================================
    # 배당금입금
    # ========================================================

    if transaction_type == "배당금입금":
        return (
            zero,
            zero,
            trade_amount,
            fee,
            raw_tax,
            "",
        )

    # ========================================================
    # 당사이체입금
    # ========================================================

    if transaction_type == "당사이체입금":
        return (
            zero,
            zero,
            trade_amount,
            fee,
            raw_tax,
            "",
        )

    # ========================================================
    # 규칙 미지정 거래
    # ========================================================

    warning_message = (
        f"규칙 미지정 거래종류: {transaction_type}"
    )

    return (
        quantity,
        unit_price,
        trade_amount,
        fee,
        raw_tax,
        warning_message,
    )


# ============================================================
# 원본 엑셀 헤더 찾기
# ============================================================

def build_header_map(
    worksheet,
) -> Dict[str, int]:
    headers = [
        clean_text(
            worksheet.cell(
                1,
                column_number,
            ).value
        )
        for column_number in range(
            1,
            worksheet.max_column + 1,
        )
    ]

    header_map: Dict[str, int] = {}

    for (
        column_number,
        header_name,
    ) in enumerate(
        headers,
        start=1,
    ):
        if not header_name:
            continue

        if header_name not in header_map:
            header_map[
                header_name
            ] = column_number

        # 거래종류 실제 텍스트가
        # '거래종류' 헤더 다음 빈 헤더 열에 있는 형식
        if (
            header_name == "거래종류"
            and column_number
            < worksheet.max_column
            and clean_text(
                worksheet.cell(
                    1,
                    column_number + 1,
                ).value
            ) == ""
        ):
            header_map[
                "거래종류_텍스트"
            ] = column_number + 1

    if "거래종류_텍스트" not in header_map:
        if "거래종류" not in header_map:
            raise KeyError(
                f"시트 '{worksheet.title}'에서 "
                "'거래종류' 헤더를 찾지 못했습니다."
            )

        header_map[
            "거래종류_텍스트"
        ] = header_map["거래종류"]

    return header_map


# ============================================================
# 원본 시트 읽기
# ============================================================

def read_sheet_rows(
    worksheet,
    fx_tables: Dict[str, ExchangeTable],
) -> Tuple[List[List], List[str]]:
    header_map = build_header_map(
        worksheet
    )

    required_headers = [
        "계좌번호",
        "계좌명",
        "거래일",
        "거래종류_텍스트",
        "종목명",
        "거래수량",
        "거래단가",
        "거래금액",
        "수수료",
        "거래세",
        "세금",
        "부가세",
        "환율",
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

    current_account_number = ""
    current_account_holder = ""

    for row_number in range(
        2,
        worksheet.max_row + 1,
    ):
        transaction_raw = clean_text(
            worksheet.cell(
                row_number,
                header_map[
                    "거래종류_텍스트"
                ],
            ).value
        )

        trade_date_value = worksheet.cell(
            row_number,
            header_map["거래일"],
        ).value

        if (
            transaction_raw == ""
            and trade_date_value in (
                None,
                "",
            )
        ):
            continue

        account_number_value = clean_text(
            worksheet.cell(
                row_number,
                header_map["계좌번호"],
            ).value
        )

        account_holder_value = clean_text(
            worksheet.cell(
                row_number,
                header_map["계좌명"],
            ).value
        )

        if account_number_value:
            current_account_number = (
                account_number_value
            )

        if account_holder_value:
            current_account_holder = (
                account_holder_value
            )

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

        (
            base_transaction,
            currency_code,
        ) = split_tx_and_currency(
            transaction_raw
        )

        row_data[
            "통화코드"
        ] = currency_code

        row_data[
            "구분기준거래종류"
        ] = base_transaction

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
                f"계산 실패({base_transaction}): "
                f"{error}"
            )

            output_quantity = Decimal("0")
            output_unit_price = Decimal("0")
            output_trade_amount = Decimal("0")
            output_fee = Decimal("0")
            output_tax = Decimal("0")

        if warning_message:
            warnings.append(
                f"[{worksheet.title} "
                f"R{row_number}] "
                f"{warning_message}"
            )

        output_rows.append(
            [
                current_account_number,
                current_account_holder,
                "",
                base_transaction,
                clean_text(
                    row_data.get("종목명")
                ),
                parse_date_safe(
                    row_data.get("거래일")
                ),
                float(output_quantity),
                float(output_unit_price),
                float(output_trade_amount),
                float(output_fee),
                float(output_tax),
            ]
        )

    return (
        output_rows,
        warnings,
    )


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
# 결과 파일 오류 검사
# ============================================================

def validate_output_file(
    file_path: Path,
) -> None:
    """
    저장한 파일을 다시 열어 파일 구조를 확인한다.

    워크시트 XML에 수식 태그가 들어 있으면
    최종 결과파일로 사용하지 않는다.
    """
    with ZipFile(
        file_path,
        "r",
    ) as archive:
        for member_name in archive.namelist():
            if not (
                member_name.startswith(
                    "xl/worksheets/"
                )
                and member_name.endswith(
                    ".xml"
                )
            ):
                continue

            xml_data = archive.read(
                member_name
            )

            if re.search(
                rb"<f(?:\s|>)",
                xml_data,
            ):
                raise ValueError(
                    "결과파일에 의도하지 않은 "
                    "수식이 발견되었습니다: "
                    f"{member_name}"
                )

    check_workbook = load_workbook(
        file_path,
        data_only=False,
        read_only=True,
    )

    check_workbook.close()


# ============================================================
# 결과 파일 저장
# ============================================================

def save_output(
    output_path: Path,
    rows: List[List],
    warnings: List[str],
) -> None:
    """
    임시파일에 먼저 저장한 뒤 검증하고,
    문제가 없을 때만 최종 파일로 교체한다.
    """
    temp_path = output_path.with_name(
        f".{output_path.stem}_작성중.xlsx"
    )

    if temp_path.exists():
        temp_path.unlink()

    workbook = Workbook()

    worksheet = workbook.active
    worksheet.title = "정리"

    # --------------------------------------------------------
    # 헤더
    # --------------------------------------------------------

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

    # --------------------------------------------------------
    # 데이터
    # A~E: 문자열
    # F: 날짜
    # G~K: 숫자
    # --------------------------------------------------------

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

    # --------------------------------------------------------
    # 검토필요 시트
    # --------------------------------------------------------

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
        # 임시파일 저장
        workbook.save(
            temp_path
        )

        workbook.close()

        # 임시파일 정상 여부 확인
        validate_output_file(
            temp_path
        )

        # 검증된 임시파일을 최종 결과파일로 교체
        try:
            os.replace(
                temp_path,
                output_path,
            )

        except PermissionError as error:
            raise PermissionError(
                f"'{output_path.name}' 파일이 "
                "Excel에서 열려 있습니다. "
                "파일을 완전히 닫은 뒤 "
                "다시 실행해 주세요."
            ) from error

    except Exception:
        workbook.close()

        if temp_path.exists():
            try:
                temp_path.unlink()

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
            "입력 파일을 찾지 못했습니다.\n"
            "예상 파일명: han26q2.xlsx 또는 "
            "한국투자증권*.xlsx"
        )

    exchange_directory = find_exchange_dir(
        BASE_DIR
    )

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

                (
                    rows,
                    warnings,
                ) = read_sheet_rows(
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
            f"{input_file.stem}"
            f"{OUTPUT_SUFFIX}"
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