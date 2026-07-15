from pathlib import Path
from collections import OrderedDict, Counter, defaultdict
from numbers import Number
import re
import sys

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side


INPUT_FILE_NAME = "삼성_AP 및 잔고.xlsx"
OUTPUT_FILE_NAME = "삼성_AP 및 잔고_정리.xlsx"
SUMMARY_SHEET_NAME = "10%초과_모음"

OUTPUT_HEADERS = [
    "상품",
    "계좌번호",
    "계약자명",
    "종목명",
    "투자비중",
    "투자사유및전망",
]

INVALID_SHEET_CHARS = re.compile(r'[\[\]:*?/\\]')


def is_number(value):
    return isinstance(value, Number) and not isinstance(value, bool)


def is_cash_or_deposit(stock_name):
    """
    10% 초과 통합 시트에서만 제외할 현금성 종목 판별.

    제외 예시:
    - 현금잔고(예수금)
    - USD(외화예수금)
    - CNY(외화예수금)
    - CASH
    """
    normalized = str(stock_name).replace(" ", "").lower()

    return (
        "현금" in normalized
        or "예수금" in normalized
        or "cash" in normalized
    )


def find_input_file(base_dir: Path) -> Path:
    preferred = base_dir / INPUT_FILE_NAME
    if preferred.exists():
        return preferred

    candidates = sorted(
        path
        for path in base_dir.glob("삼성_AP 및 잔고*.xlsx")
        if "계좌별정리" not in path.stem
        and "_정리" not in path.stem
        and not path.name.startswith("~$")
    )

    if not candidates:
        raise FileNotFoundError(
            f"'{INPUT_FILE_NAME}' 파일을 찾지 못했습니다.\n"
            f"확인한 폴더: {base_dir}"
        )

    return candidates[0]


def find_source_sheet(workbook):
    required = {"계좌번호", "계좌명", "종목명", "평가금액"}

    for ws in workbook.worksheets:
        headers = {
            str(cell.value).strip()
            for cell in ws[1]
            if cell.value is not None
        }

        if required.issubset(headers):
            return ws

    raise ValueError(
        "계좌번호, 계좌명, 종목명, 평가금액 헤더가 있는 시트를 찾지 못했습니다."
    )


def build_header_map(ws):
    header_map = {}

    for cell in ws[1]:
        if cell.value is not None:
            header_map[str(cell.value).strip()] = cell.column

    required = ["계좌번호", "계좌명", "종목명", "평가금액"]
    missing = [header for header in required if header not in header_map]

    if missing:
        raise ValueError(f"필수 헤더를 찾지 못했습니다: {missing}")

    return header_map


def group_account_rows(ws, header_map):
    account_col = header_map["계좌번호"]
    name_col = header_map["계좌명"]
    stock_col = header_map["종목명"]
    valuation_col = header_map["평가금액"]

    groups = OrderedDict()

    for row_no in range(2, ws.max_row + 1):
        account_no = ws.cell(row=row_no, column=account_col).value
        account_name = ws.cell(row=row_no, column=name_col).value
        stock_name = ws.cell(row=row_no, column=stock_col).value
        valuation = ws.cell(row=row_no, column=valuation_col).value

        if account_no in (None, "") or account_name in (None, ""):
            continue
        if stock_name in (None, ""):
            continue
        if not is_number(valuation):
            print(
                f"[건너뜀] {row_no}행: 평가금액이 숫자가 아닙니다. "
                f"값={valuation!r}"
            )
            continue

        key = (
            str(account_no).strip(),
            str(account_name).strip(),
        )

        if key not in groups:
            groups[key] = OrderedDict()

        normalized_stock_name = str(stock_name).strip()
        groups[key][normalized_stock_name] = (
            groups[key].get(normalized_stock_name, 0.0)
            + float(valuation)
        )

    return groups


def safe_sheet_name(raw_name, used_names):
    cleaned = INVALID_SHEET_CHARS.sub("_", str(raw_name))
    cleaned = cleaned.strip().strip("'")

    if not cleaned:
        cleaned = "계좌"

    cleaned = cleaned[:31]
    candidate = cleaned
    suffix_no = 2

    while candidate in used_names:
        suffix = f"_{suffix_no}"
        candidate = f"{cleaned[:31-len(suffix)]}{suffix}"
        suffix_no += 1

    used_names.add(candidate)
    return candidate


def build_sheet_names(groups):
    name_counts = Counter(
        account_name
        for _, account_name in groups.keys()
    )
    name_running = defaultdict(int)
    used_names = {SUMMARY_SHEET_NAME}
    result = OrderedDict()

    for account_no, account_name in groups.keys():
        if name_counts[account_name] > 1:
            name_running[account_name] += 1
            raw_name = f"{account_name}{name_running[account_name]}"
        else:
            raw_name = account_name

        result[(account_no, account_name)] = safe_sheet_name(
            raw_name,
            used_names,
        )

    return result


def build_account_rows(groups):
    """
    개별 고객 시트:
    - 현금과 예수금도 그대로 포함
    - 투자비중 계산의 분모에도 포함

    10%초과_모음 시트:
    - 투자비중이 10%를 초과하는 종목만 포함
    - 종목명에 현금, 예수금, CASH가 들어간 행은 제외
    """
    account_rows = OrderedDict()
    summary_rows = []
    excluded_cash_count = 0

    for key, stock_totals in groups.items():
        account_no, account_name = key
        total_account_value = sum(stock_totals.values())

        if total_account_value == 0:
            raise ValueError(
                f"{account_name}({account_no}) 계좌의 "
                "총 평가금액이 0이라 투자비중을 계산할 수 없습니다."
            )

        rows = []

        for stock_name, valuation in stock_totals.items():
            weight = valuation / total_account_value

            rows.append([
                "",
                account_no,
                account_name,
                stock_name,
                weight,
                "",
            ])

        rows.sort(key=lambda row: row[4], reverse=True)
        account_rows[key] = rows

        for row in rows:
            if row[4] <= 0.10:
                continue

            if is_cash_or_deposit(row[3]):
                excluded_cash_count += 1
                continue

            summary_rows.append(row)

    return account_rows, summary_rows, excluded_cash_count


def write_result_sheet(ws, rows):
    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(color="FFFFFF", bold=True)
    thin = Side(style="thin", color="D9E2F3")
    border = Border(
        left=thin,
        right=thin,
        top=thin,
        bottom=thin,
    )

    for column_no, header in enumerate(OUTPUT_HEADERS, start=1):
        cell = ws.cell(row=1, column=column_no, value=header)
        cell.fill = header_fill
        cell.font = header_font
        cell.border = border
        cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
        )

    for row_no, row_values in enumerate(rows, start=2):
        for column_no, value in enumerate(row_values, start=1):
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

        ws.cell(row=row_no, column=5).number_format = "0.00%"

    last_row = max(1, len(rows) + 1)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:F{last_row}"
    ws.sheet_view.showGridLines = False

    column_widths = {
        "A": 14,
        "B": 20,
        "C": 16,
        "D": 38,
        "E": 14,
        "F": 34,
    }

    for column_letter, width in column_widths.items():
        ws.column_dimensions[column_letter].width = width

    ws.row_dimensions[1].height = 24


def create_output(groups, output_path):
    output_wb = Workbook()
    output_wb.remove(output_wb.active)

    sheet_names = build_sheet_names(groups)
    account_rows, summary_rows, excluded_cash_count = build_account_rows(groups)

    # 맨 앞의 최종 요약 시트
    summary_ws = output_wb.create_sheet(SUMMARY_SHEET_NAME)
    write_result_sheet(summary_ws, summary_rows)

    # 고객별 시트는 기존과 동일
    for key, rows in account_rows.items():
        ws = output_wb.create_sheet(sheet_names[key])
        write_result_sheet(ws, rows)

    output_wb.save(output_path)

    return (
        len(account_rows),
        len(summary_rows),
        excluded_cash_count,
    )


def main():
    base_dir = Path(__file__).resolve().parent

    # 사용법
    # python samsung_ap_normalizer_openpyxl_with_10pct_summary_no_cash.py
    # python samsung_ap_normalizer_openpyxl_with_10pct_summary_no_cash.py "삼성_AP 및 잔고.xlsx"
    # python samsung_ap_normalizer_openpyxl_with_10pct_summary_no_cash.py "원본.xlsx" "결과.xlsx"

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
            else base_dir / requested_output
        )
    else:
        output_path = base_dir / OUTPUT_FILE_NAME

    source_wb = load_workbook(
        input_path,
        data_only=True,
        read_only=False,
    )

    source_ws = find_source_sheet(source_wb)
    header_map = build_header_map(source_ws)
    groups = group_account_rows(source_ws, header_map)

    if not groups:
        raise ValueError("처리할 계좌 데이터가 없습니다.")

    (
        account_count,
        summary_count,
        excluded_cash_count,
    ) = create_output(groups, output_path)

    print(f"원본 파일: {input_path.name}")
    print(f"원본 시트: {source_ws.title}")
    print(f"생성 계좌 시트 수: {account_count}개")
    print(f"10% 초과 통합 행 수: {summary_count}개")
    print(f"요약 시트에서 제외한 현금·예수금 행 수: {excluded_cash_count}개")
    print(f"완료 파일: {output_path}")


if __name__ == "__main__":
    main()
