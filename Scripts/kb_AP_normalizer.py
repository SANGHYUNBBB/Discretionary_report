from __future__ import annotations

import bisect
import os
import re
import sys
from dataclasses import dataclass
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from numbers import Number
from pathlib import Path
from typing import Dict, List, Optional, Tuple

from report_paths import ap_input_dir, ap_output_dir, exchange_dir, settings, single_ap_input

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side


# ============================================================
# 기본 설정
# ============================================================

INPUT_FILE_NAME = "KB_AP 및 잔고.xlsx"
OUTPUT_FILE_NAME = "KB_AP 및 잔고_정리.xlsx"

ALL_SHEET_NAME = "전체_내림차순"
OVER_10_SHEET_NAME = "10%초과"

# None이면 exchange_rate 파일의 가장 최근 날짜 환율을 사용한다.
# 특정 기준일을 강제로 쓰고 싶으면 예: "2026-06-30"
EXCHANGE_LOOKUP_DATE = settings().get("ap_exchange_date")

HEADERS = [
    "상품",
    "계좌번호",
    "계약자명",
    "종목명",
    "투자비중",
    "투자사유및전망",
    "보유수량",
    "취득단가(원화)",
    "매입액원화",
    "시가(분기말기준,원화)",
    "보유잔액(원화)",
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
        기준일과 동일한 날짜의 환율을 우선 사용한다.
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

    def latest(self) -> Decimal:
        """환율 파일에서 가장 최근 날짜의 환율을 사용한다."""
        if not self.rates:
            raise ValueError("환율 데이터가 비어 있습니다.")

        return self.rates[-1]


# ============================================================
# 공통 함수
# ============================================================

def is_number(value):
    """bool을 제외한 숫자인지 확인한다."""
    return isinstance(value, Number) and not isinstance(value, bool)


def clean_text(value) -> str:
    if value is None:
        return ""

    return str(value).strip()


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
            result = Decimal(str(value).replace(",", "").strip())

        except (
            InvalidOperation,
            AttributeError,
            ValueError,
        ):
            return ZERO

    if not result.is_finite():
        return ZERO

    return result


def decimal_to_float_or_none(value: Decimal | None):
    if value is None:
        return None

    return float(value)


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
            return datetime.strptime(text, date_format).date()

        except ValueError:
            continue

    return None


def parse_lookup_date(value) -> Optional[date]:
    if value in (None, ""):
        return None

    parsed_date = parse_date_safe(value)

    if parsed_date is None:
        raise ValueError(
            f"EXCHANGE_LOOKUP_DATE 값을 날짜로 해석하지 못했습니다: {value!r}"
        )

    return parsed_date


def is_cash_record(record):
    """10%초과 시트에서 제외할 예수금 행인지 판별한다."""
    if len(record) < 4:
        return False

    stock_name = str(record[3]).replace(" ", "")

    return "예수금" in stock_name


# ============================================================
# 파일 및 폴더 검색
# ============================================================

def find_input_file(base_dir):
    return single_ap_input("KB")


def find_exchange_dir(base_dir):
    path = exchange_dir()
    return path if path.is_dir() else None


def extract_currency_code_from_filename(file_path: Path) -> str:
    """
    환율 파일명에서 통화코드를 추출한다.

    예:
    USD.xlsx
    USD_환율.xlsx
    환율_USD.xlsx
    """
    stem_upper = file_path.stem.strip().upper()

    match = re.search(
        r"(?<![A-Z])([A-Z]{3})(?![A-Z])",
        stem_upper,
    )

    if match:
        return match.group(1)

    return stem_upper


# ============================================================
# 환율 처리
# ============================================================

def adjust_fx_rate(currency_code: str, rate: Decimal) -> Decimal:
    """
    JPY 환율은 보통 100엔당 환율이므로
    100으로 나누어 1엔당 환율로 바꾼다.
    """
    currency_code = clean_text(currency_code).upper()

    if currency_code == "JPY":
        return rate / HUNDRED

    return rate


def load_xlsx_exchange_file(file_path: Path) -> List[Tuple[int, Decimal]]:
    """
    xlsx/xlsm 환율 파일 읽기.

    파일 구조:
    - 첫 번째 시트 사용
    - A열: 날짜
    - C열: 환율
    - 10행부터 데이터
    """
    workbook = load_workbook(
        file_path,
        data_only=True,
        read_only=True,
    )

    try:
        worksheet = workbook[workbook.sheetnames[0]]
        date_rate_pairs: List[Tuple[int, Decimal]] = []

        for row_number in range(10, worksheet.max_row + 1):
            exchange_date = parse_date_safe(
                worksheet[f"A{row_number}"].value
            )

            exchange_rate = to_decimal(
                worksheet[f"C{row_number}"].value
            )

            if exchange_date is None or exchange_rate == ZERO:
                continue

            date_rate_pairs.append(
                (
                    exchange_date.toordinal(),
                    exchange_rate,
                )
            )

    finally:
        workbook.close()

    return date_rate_pairs


def load_xls_exchange_file(file_path: Path) -> List[Tuple[int, Decimal]]:
    """
    xls 환율 파일 읽기.

    xlrd가 설치되어 있을 때만 동작한다.
    설치가 안 되어 있으면:
        pip install xlrd
    """
    try:
        import xlrd

    except ImportError as error:
        raise ImportError(
            f"{file_path.name} 파일은 .xls 형식입니다. "
            "이 파일을 읽으려면 xlrd가 필요합니다.\n"
            "설치 명령어: pip install xlrd\n"
            "또는 환율 파일을 .xlsx로 저장해 주세요."
        ) from error

    workbook = xlrd.open_workbook(str(file_path))
    worksheet = workbook.sheet_by_index(0)

    date_rate_pairs: List[Tuple[int, Decimal]] = []

    # Excel 기준 10행부터이므로 xlrd zero-based index는 9부터.
    for row_index in range(9, worksheet.nrows):
        raw_date = worksheet.cell_value(row_index, 0)
        raw_rate = worksheet.cell_value(row_index, 2)

        exchange_date = None

        cell_type = worksheet.cell_type(row_index, 0)
        if cell_type == xlrd.XL_CELL_DATE:
            exchange_date = xlrd.xldate_as_datetime(
                raw_date,
                workbook.datemode,
            ).date()
        else:
            exchange_date = parse_date_safe(raw_date)

        exchange_rate = to_decimal(raw_rate)

        if exchange_date is None or exchange_rate == ZERO:
            continue

        date_rate_pairs.append(
            (
                exchange_date.toordinal(),
                exchange_rate,
            )
        )

    return date_rate_pairs


def load_exchange_rates(exchange_dir: Optional[Path]) -> Dict[str, ExchangeTable]:
    """
    exchange_rate 폴더의 환율 파일을 읽는다.

    중요:
    - 같은 통화에 .xlsx/.xlsm 파일과 .xls 파일이 같이 있으면 .xlsx/.xlsm을 우선 사용한다.
    - .xls 파일이 실제 구형 Excel 파일이 아니면 xlrd에서 오류가 날 수 있으므로,
      같은 통화의 .xlsx/.xlsm 대체 파일이 있으면 해당 .xls는 건너뛴다.
    - 대체 파일이 없는데 .xls 읽기에 실패하면 어떤 파일이 문제인지 알려준다.
    """
    rate_map: Dict[str, ExchangeTable] = {}

    if exchange_dir is None:
        return rate_map

    candidate_files: List[Path] = []

    for extension in (
        "*.xlsx",
        "*.xlsm",
        "*.xls",
    ):
        candidate_files.extend(exchange_dir.glob(extension))
        candidate_files.extend(exchange_dir.glob(f"**/{extension}"))

    unique_files = []
    seen_paths = set()

    for file_path in candidate_files:
        resolved_path = str(file_path.resolve())

        if resolved_path in seen_paths:
            continue

        if file_path.name.startswith("~$"):
            continue

        seen_paths.add(resolved_path)
        unique_files.append(file_path)

    files_by_currency: Dict[str, List[Path]] = {}

    for file_path in unique_files:
        currency_code = extract_currency_code_from_filename(file_path)

        if currency_code not in CURRENCY_CODES:
            continue

        files_by_currency.setdefault(currency_code, []).append(file_path)

    for currency_code, files in sorted(files_by_currency.items()):
        # 같은 통화에 여러 파일이 있으면 xlsx/xlsm을 먼저 사용한다.
        files_sorted = sorted(
            files,
            key=lambda path: (
                0 if path.suffix.lower() in {".xlsx", ".xlsm"} else 1,
                path.name.lower(),
            ),
        )

        last_error = None

        for file_path in files_sorted:
            suffix = file_path.suffix.lower()

            try:
                if suffix in {".xlsx", ".xlsm"}:
                    date_rate_pairs = load_xlsx_exchange_file(file_path)

                elif suffix == ".xls":
                    date_rate_pairs = load_xls_exchange_file(file_path)

                else:
                    continue

                if not date_rate_pairs:
                    raise ValueError(
                        "환율 파일에서 날짜와 환율을 찾지 못했습니다: "
                        f"{file_path}"
                    )

                date_rate_pairs.sort(key=lambda item: item[0])

                rate_map[currency_code] = ExchangeTable(
                    dates_ord=[item[0] for item in date_rate_pairs],
                    rates=[item[1] for item in date_rate_pairs],
                )

                break

            except Exception as error:
                last_error = error
                print(f"[환율파일 건너뜀] {file_path.name}: {error}")
                continue

        if currency_code not in rate_map:
            checked_files = ", ".join(path.name for path in files_sorted)
            raise ValueError(
                f"{currency_code} 환율 파일을 읽지 못했습니다. "
                f"확인한 파일: {checked_files}. "
                f"마지막 오류: {last_error}"
            )

    return rate_map


def get_fx_rate(currency_code: str, fx_tables: Dict[str, ExchangeTable]) -> Decimal:
    """
    KB AP 파일에는 별도 환율 열이 없으므로,
    exchange_rate 폴더에서 환율을 조회한다.

    KRW: 1 사용
    외화: EXCHANGE_LOOKUP_DATE가 있으면 해당일 또는 직전일 사용
          EXCHANGE_LOOKUP_DATE가 None이면 가장 최근 환율 사용
    JPY: 최종 환율을 100으로 나누어 1엔당 환율로 변환
    """
    currency_code = clean_text(currency_code).upper()

    if currency_code in ("", "KRW"):
        return Decimal("1")

    exchange_table = fx_tables.get(currency_code)

    if exchange_table is None:
        available_codes = (
            ", ".join(sorted(fx_tables.keys()))
            if fx_tables
            else "없음"
        )

        raise KeyError(
            f"통화코드 {currency_code}에 해당하는 환율 파일을 찾지 못했습니다. "
            f"exchange_rate 폴더와 파일명을 확인해 주세요. "
            f"사용 가능한 통화코드: {available_codes}"
        )

    lookup_date = parse_lookup_date(EXCHANGE_LOOKUP_DATE)

    if lookup_date is None:
        rate = exchange_table.latest()
    else:
        rate = exchange_table.lookup(lookup_date)

    return adjust_fx_rate(currency_code, rate)


def get_fx_rate_cache(currency_code: str, fx_tables: Dict[str, ExchangeTable], cache: dict) -> Decimal:
    currency_code = clean_text(currency_code).upper() or "KRW"

    if currency_code not in cache:
        cache[currency_code] = get_fx_rate(currency_code, fx_tables)

    return cache[currency_code]


# ============================================================
# KB AP 처리
# ============================================================

def find_net_asset_value(ws):
    """
    1~8행에서 '순자산평가금액'을 찾은 뒤,
    바로 오른쪽 셀의 숫자를 반환한다.
    """
    for row in ws.iter_rows(min_row=1, max_row=min(8, ws.max_row)):
        for cell in row:
            if cell.value == "순자산평가금액":
                value = ws.cell(
                    row=cell.row,
                    column=cell.column + 1,
                ).value

                if is_number(value) and float(value) != 0:
                    return float(value)

    raise ValueError(
        f"'{ws.title}' 시트에서 유효한 순자산평가금액을 찾지 못했습니다."
    )


def find_summary_value(ws, label):
    """
    1~8행에서 특정 요약 항목을 찾은 뒤,
    바로 오른쪽 셀의 숫자를 반환한다.

    예:
    - A2 = T+2 예수금, B2 = 금액
    - A5 = 외화예수금원화환산액, B5 = 금액
    """
    for row in ws.iter_rows(min_row=1, max_row=min(8, ws.max_row)):
        for cell in row:
            if clean_text(cell.value) == label:
                value = ws.cell(
                    row=cell.row,
                    column=cell.column + 1,
                ).value

                if is_number(value):
                    return float(value)

                return 0.0

    return 0.0


def build_cash_records(ws, account_no, contractor, net_asset_value):
    """
    KB 예수금 행을 결과 행으로 만든다.

    원화:
    - T+2 예수금

    외화:
    - 외화예수금원화환산액
    - KB는 통화별 구분 없이 '외화예수금'으로 통합 표시한다.

    예수금 행:
    - 전체_내림차순에는 포함
    - 10%초과 시트에는 제외
    - 취득단가(원화), 시가(분기말기준,원화)는 0
    - 투자비중 = 예수금 보유잔액(원화) / 순자산평가금액
    """
    cash_rows = []

    cash_specs = [
        ("원화예수금", "T+2 예수금"),
        ("외화예수금", "외화예수금원화환산액"),
    ]

    for cash_name, source_label in cash_specs:
        cash_amount = find_summary_value(ws, source_label)

        if not is_number(cash_amount) or float(cash_amount) == 0:
            continue

        cash_amount = float(cash_amount)

        cash_rows.append([
            "",                                 # 상품
            str(account_no),                    # 계좌번호
            str(contractor),                    # 계약자명
            cash_name,                          # 종목명
            cash_amount / net_asset_value,      # 투자비중
            "",                                 # 투자사유및전망
            cash_amount,                        # 보유수량
            0,                                  # 취득단가(원화)
            cash_amount,                        # 매입액원화
            0,                                  # 시가(분기말기준,원화)
            cash_amount,                        # 보유잔액(원화)
        ])

    return cash_rows


def build_header_map(ws, header_row):
    header_map = {}

    for column_no in range(1, ws.max_column + 1):
        header = clean_text(ws.cell(row=header_row, column=column_no).value)

        if header and header not in header_map:
            header_map[header] = column_no

    return header_map


def require_headers(ws, header_map, required_headers):
    missing = [
        header
        for header in required_headers
        if header not in header_map
    ]

    if missing:
        raise ValueError(
            f"'{ws.title}' 시트에서 필수 헤더를 찾지 못했습니다: {missing}"
        )


def extract_records(source_wb, fx_tables):
    """
    각 고객 시트에서 종목 정보를 추출한다.

    원본 기준:
    - A1: 계약자명
    - B1: 계좌번호
    - 9행: 종목 헤더
    - 보유수량: 주문가능수량
    - 현재가: 현재가
    - 평가금액: 평가금액
    - 통화코드: 통화코드
    - 종목명: 종목명
    - 매입금액: 매입금액
    - 매입평균가: 매입평균가

    계산:
    - 종목 투자비중 = 평가금액 / 순자산평가금액
    - 종목 취득단가(원화) = 매입평균가 * 환율
    - 종목 매입액원화 = 매입금액
    - 종목 시가(분기말기준,원화) = 현재가 * 환율
    - 종목 보유잔액(원화) = 평가금액

    예수금:
    - 원화예수금 = T+2 예수금
    - 외화예수금 = 외화예수금원화환산액
    - 예수금 투자비중 = 예수금 보유잔액(원화) / 순자산평가금액
    - 예수금 취득단가(원화), 시가(분기말기준,원화)는 0
    - 예수금은 전체_내림차순에만 포함하고 10%초과 시트에서는 제외
    """
    all_records = []
    fx_cache = {}

    for ws in source_wb.worksheets:
        contractor = ws["A1"].value
        account_no = ws["B1"].value

        if not contractor or not account_no:
            print(
                f"[건너뜀] '{ws.title}' 시트: "
                "A1 계약자명 또는 B1 계좌번호가 비어 있습니다."
            )
            continue

        header_map = build_header_map(ws, 9)

        try:
            require_headers(
                ws,
                header_map,
                [
                    "주문가능수량",
                    "현재가",
                    "평가금액",
                    "통화코드",
                    "종목명",
                    "매입금액",
                    "매입평균가",
                ],
            )
        except ValueError as exc:
            print(f"[건너뜀] {exc}")
            continue

        try:
            net_asset_value = find_net_asset_value(ws)
        except ValueError as exc:
            print(f"[건너뜀] {exc}")
            continue

        quantity_col = header_map["주문가능수량"]
        current_price_col = header_map["현재가"]
        valuation_col = header_map["평가금액"]
        currency_col = header_map["통화코드"]
        stock_name_col = header_map["종목명"]
        buy_amount_col = header_map["매입금액"]
        average_buy_price_col = header_map["매입평균가"]

        customer_records = []

        # 예수금 행 추가:
        # 원화예수금 = T+2 예수금
        # 외화예수금 = 외화예수금원화환산액
        customer_records.extend(
            build_cash_records(
                ws=ws,
                account_no=account_no,
                contractor=contractor,
                net_asset_value=net_asset_value,
            )
        )

        for row_no in range(10, ws.max_row + 1):
            quantity = ws.cell(row=row_no, column=quantity_col).value  # 주문가능수량
            average_buy_price = ws.cell(row=row_no, column=average_buy_price_col).value
            buy_amount = ws.cell(row=row_no, column=buy_amount_col).value
            current_price = ws.cell(row=row_no, column=current_price_col).value
            valuation = ws.cell(row=row_no, column=valuation_col).value
            currency_code = ws.cell(row=row_no, column=currency_col).value
            stock_name = ws.cell(row=row_no, column=stock_name_col).value

            if stock_name in (None, "") and valuation in (None, ""):
                continue

            if stock_name in (None, ""):
                print(
                    f"[건너뜀] '{ws.title}' 시트 {row_no}행: "
                    "종목명이 비어 있습니다."
                )
                continue

            if not is_number(valuation):
                print(
                    f"[건너뜀] '{ws.title}' 시트 {row_no}행: "
                    f"평가금액이 숫자가 아닙니다. 값={valuation!r}"
                )
                continue

            exchange_rate = get_fx_rate_cache(
                currency_code,
                fx_tables,
                fx_cache,
            )

            if not is_number(average_buy_price):
                print(
                    f"[건너뜀] '{ws.title}' 시트 {row_no}행: "
                    f"매입평균가가 숫자가 아닙니다. 값={average_buy_price!r}"
                )
                acquisition_unit_krw = None
            else:
                acquisition_unit_krw = float(
                    Decimal(str(average_buy_price)) * exchange_rate
                )

            if not is_number(current_price):
                print(
                    f"[건너뜀] '{ws.title}' 시트 {row_no}행: "
                    f"현재가가 숫자가 아닙니다. 값={current_price!r}"
                )
                market_price_krw = None
            else:
                market_price_krw = float(
                    Decimal(str(current_price)) * exchange_rate
                )

            customer_records.append([
                "",                                     # 상품
                str(account_no),                        # 계좌번호
                str(contractor),                        # 계약자명
                str(stock_name),                        # 종목명
                float(valuation) / net_asset_value,     # 투자비중
                "",                                     # 투자사유및전망
                quantity if is_number(quantity) else None,
                acquisition_unit_krw,                   # 매입평균가 * 환율
                buy_amount if is_number(buy_amount) else None,
                market_price_krw,                       # 현재가 * 환율
                valuation,                              # 평가금액
            ])

        customer_records.sort(
            key=lambda record: record[4],
            reverse=True,
        )

        all_records.extend(customer_records)

    return all_records


# ============================================================
# 결과 파일 작성
# ============================================================

def write_result_sheet(ws, records):
    """
    표(Table) 객체는 만들지 않는다.

    Excel 복구 경고를 피하기 위해
    일반 범위 + 일반 자동필터만 사용한다.
    """
    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(color="FFFFFF", bold=True)
    thin_side = Side(style="thin", color="D9E2F3")
    border = Border(
        left=thin_side,
        right=thin_side,
        top=thin_side,
        bottom=thin_side,
    )

    for column_no, header in enumerate(HEADERS, start=1):
        cell = ws.cell(row=1, column=column_no, value=header)
        cell.fill = header_fill
        cell.font = header_font
        cell.border = border
        cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True,
        )

    for row_no, record in enumerate(records, start=2):
        for column_no, value in enumerate(record, start=1):
            cell = ws.cell(
                row=row_no,
                column=column_no,
                value=value,
            )

            cell.border = border
            cell.alignment = Alignment(
                vertical="center",
                wrap_text=column_no in {4, 6},
            )

        ws.cell(row=row_no, column=5).number_format = "0.00%"          # 투자비중
        ws.cell(row=row_no, column=7).number_format = "#,##0.########" # 보유수량
        ws.cell(row=row_no, column=8).number_format = "#,##0.00"       # 취득단가
        ws.cell(row=row_no, column=9).number_format = "#,##0"          # 매입액원화
        ws.cell(row=row_no, column=10).number_format = "#,##0.00"      # 시가 = 현재가 * 환율
        ws.cell(row=row_no, column=11).number_format = "#,##0"         # 보유잔액

    last_row = max(1, len(records) + 1)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:K{last_row}"
    ws.sheet_view.showGridLines = False
    ws.row_dimensions[1].height = 36

    column_widths = {
        "A": 14,
        "B": 20,
        "C": 16,
        "D": 42,
        "E": 14,
        "F": 36,
        "G": 14,
        "H": 18,
        "I": 18,
        "J": 22,
        "K": 18,
    }

    for column_letter, width in column_widths.items():
        ws.column_dimensions[column_letter].width = width


def create_output_workbook(records, output_path: Path):
    """
    하나의 파일에 두 시트를 생성한다.

    - 전체_내림차순
    - 10%초과: 투자비중 > 10%
    """
    output_wb = Workbook()
    output_wb.remove(output_wb.active)

    over_10_records = [
        record for record in records
        if record[4] > 0.10 and not is_cash_record(record)
    ]

    all_ws = output_wb.create_sheet(ALL_SHEET_NAME)
    write_result_sheet(all_ws, records)

    over_10_ws = output_wb.create_sheet(OVER_10_SHEET_NAME)
    write_result_sheet(over_10_ws, over_10_records)

    temporary_path = output_path.with_name(
        f".{output_path.stem}_작성중.xlsx"
    )

    if temporary_path.exists():
        temporary_path.unlink()

    try:
        output_wb.save(temporary_path)

        verification_workbook = load_workbook(
            temporary_path,
            data_only=False,
            read_only=True,
        )
        verification_workbook.close()

        try:
            os.replace(temporary_path, output_path)

        except PermissionError as error:
            raise PermissionError(
                f"'{output_path.name}' 파일이 Excel에서 열려 있습니다. "
                "결과 파일을 닫고 다시 실행해 주세요."
            ) from error

    except Exception:
        if temporary_path.exists():
            try:
                temporary_path.unlink()
            except OSError:
                pass
        raise

    return len(over_10_records)


def main():
    base_dir = ap_input_dir("KB")
    output_dir = ap_output_dir("KB")

    # 사용법
    # python kb_AP_normalizer.py
    # python kb_AP_normalizer.py "KB_AP 및 잔고.xlsx"
    # python kb_AP_normalizer.py "원본.xlsx" "결과.xlsx"
    if len(sys.argv) >= 2:
        requested_input = Path(sys.argv[1])
        input_path = (
            requested_input
            if requested_input.is_absolute()
            else base_dir / requested_input
        )
    else:
        input_path = find_input_file(base_dir)

    if not input_path.exists():
        raise FileNotFoundError(
            f"입력 파일을 찾지 못했습니다: {input_path}"
        )

    if len(sys.argv) >= 3:
        requested_output = Path(sys.argv[2])
        output_path = (
            requested_output
            if requested_output.is_absolute()
            else output_dir / requested_output
        )
    else:
        output_path = output_dir / OUTPUT_FILE_NAME

    exchange_dir = find_exchange_dir(base_dir)
    fx_tables = load_exchange_rates(exchange_dir)

    source_wb = load_workbook(
        input_path,
        data_only=True,
        read_only=False,
    )

    try:
        records = extract_records(source_wb, fx_tables)
    finally:
        source_wb.close()

    over_10_count = create_output_workbook(
        records,
        output_path,
    )

    print(f"원본 파일: {input_path.name}")
    print(f"환율 폴더: {exchange_dir if exchange_dir else '찾지 못함'}")
    print(f"사용 가능 환율: {', '.join(sorted(fx_tables.keys())) if fx_tables else '없음'}")
    print(f"고객 시트 수: {len(source_wb.worksheets)}개")
    cash_count = sum(1 for record in records if is_cash_record(record))

    print(f"전체 행 수: {len(records)}개")
    print(f"예수금 행 수: {cash_count}개")
    print(f"10% 초과 종목 수: {over_10_count}개")
    print(f"완료 파일: {output_path}")


if __name__ == "__main__":
    main()