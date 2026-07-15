from __future__ import annotations

import math
import re
import unicodedata
from datetime import date, datetime
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Sequence, Set, Tuple

from openpyxl import Workbook, load_workbook
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


# ============================================================
# 기본 설정
# ============================================================

# 이 .py 파일이 있는 폴더를 작업폴더로 사용한다.
# 따라서 폴더 전체를 다른 위치로 옮겨도 정상 작동한다.
BASE_DIR = Path(__file__).resolve().parent

OUTPUT_FILE_NAME = "SumTradeList.xlsx"

# 출력 시트명, 입력 파일 기본 이름
SOURCE_CONFIG: Sequence[Tuple[str, str]] = (
    ("삼성", "samsung26q2_정리"),
    ("메리츠", "Meritz26q2_정리"),
    ("KB", "kb26q2_정리"),
    ("교보", "kyobo26q2_정리"),
    ("한국투자", "han26q2_정리"),
)

CONTRACT_FILE_STEMS = {
    "contractlist",
    "ㅊontractlist",  # 파일명이 오타로 저장된 경우도 허용
}

ACCOUNT_HEADER_ALIASES = {
    "계좌번호",
    "계좌 번호",
    "증권계좌번호",
    "증권 계좌번호",
    "위탁계좌번호",
    "위탁 계좌번호",
}

CONTRACT_STATUS_HEADER_ALIASES = {
    "구분",
    "계약상태",
    "계약 상태",
    "상태",
    "계약요청상태",
    "계약 요청 상태",
}

# ContractList.xlsx에서 상품명을 읽을 때 사용할 헤더명
PRODUCT_HEADER_ALIASES = {
    "상품명",
    "상품 명",
    "상품",
    "계약상품명",
    "계약 상품명",
    "일임상품명",
    "일임 상품명",
    "일임상품",
    "포트폴리오명",
    "포트폴리오 명",
    "PF명",
    "PF 명",
}

# 정리 파일에서 상품명을 써 넣을 PF명 열
PF_HEADER_ALIASES = {
    "PF명",
    "PF 명",
    "PF",
    "포트폴리오명",
    "포트폴리오 명",
}

CONTRACT_NAME_HEADER_ALIASES = {
    "계약자명",
    "계약자 명",
    "고객명",
    "고객 명",
    "계좌명",
    "계좌 명",
    "성명",
    "이름",
}

# 해당 단어가 계약상태에 포함되면 살아있는 계약에서 제외한다.
INACTIVE_STATUS_KEYWORDS = {
    "해지",
    "취소",
    "만료",
    "종료",
    "반려",
    "거부",
    "철회",
}

HEADER_SCAN_ROWS = 30


# ============================================================
# 문자열 / 계좌번호 처리
# ============================================================

def clean_text(value) -> str:
    if value is None:
        return ""

    text = unicodedata.normalize("NFKC", str(value))
    text = ILLEGAL_CHARACTERS_RE.sub("", text)
    text = text.replace("\u00A0", " ").replace("\u202F", " ")
    return text.strip()


def normalize_header(value) -> str:
    """헤더 비교를 위해 공백·줄바꿈을 제거한다."""
    text = clean_text(value)
    return re.sub(r"\s+", "", text).casefold()


def normalized_aliases(values: Iterable[str]) -> Set[str]:
    return {normalize_header(value) for value in values}


ACCOUNT_HEADER_KEYS = normalized_aliases(ACCOUNT_HEADER_ALIASES)
STATUS_HEADER_KEYS = normalized_aliases(CONTRACT_STATUS_HEADER_ALIASES)
PRODUCT_HEADER_KEYS = normalized_aliases(PRODUCT_HEADER_ALIASES)
PF_HEADER_KEYS = normalized_aliases(PF_HEADER_ALIASES)
CONTRACT_NAME_HEADER_KEYS = normalized_aliases(CONTRACT_NAME_HEADER_ALIASES)


def normalize_account(value) -> str:
    """
    계좌번호 비교용 키 생성.

    예:
      7164134371-01  -> 716413437101
      1027 84981 01 -> 10278498101
      '001-20-003707 -> 00120003707
    """
    if value is None:
        return ""

    if isinstance(value, bool):
        return ""

    if isinstance(value, int):
        text = str(value)
    elif isinstance(value, float):
        if not math.isfinite(value):
            return ""
        text = str(int(value)) if value.is_integer() else format(value, ".15g")
    else:
        text = clean_text(value)
        if text.startswith("'"):
            text = text[1:]

        # 문자열로 저장된 12345.0 형태 보정
        if re.fullmatch(r"[+-]?\d+\.0+", text):
            text = text.split(".", 1)[0]

    text = unicodedata.normalize("NFKC", text).upper()
    # 하이픈, 공백 등 표시문자만 제거하고 영문/숫자는 유지
    return re.sub(r"[^0-9A-Z]", "", text)


def normalize_name(value) -> str:
    """이름 비교용 키. 공백·대소문자·전각/반각 차이를 제거한다."""
    text = unicodedata.normalize("NFKC", clean_text(value)).casefold()
    return re.sub(r"\s+", "", text)


def is_inactive_contract(status_value) -> bool:
    status = re.sub(r"\s+", "", clean_text(status_value)).casefold()
    if not status:
        return False
    return any(keyword.casefold() in status for keyword in INACTIVE_STATUS_KEYWORDS)


# ============================================================
# 파일 찾기
# ============================================================

def workbook_candidates() -> List[Path]:
    result: List[Path] = []
    for path in BASE_DIR.iterdir():
        if not path.is_file():
            continue
        if path.name.startswith("~$"):
            continue
        if path.suffix.lower() not in {".xlsx", ".xlsm"}:
            continue
        result.append(path)
    return result


def find_contract_file(files: Sequence[Path]) -> Path:
    exact_matches = [
        path
        for path in files
        if path.stem.casefold() in CONTRACT_FILE_STEMS
    ]

    if exact_matches:
        return sorted(exact_matches, key=lambda p: p.stat().st_mtime, reverse=True)[0]

    partial_matches = [
        path
        for path in files
        if "contractlist" in path.stem.casefold()
    ]

    if partial_matches:
        return sorted(partial_matches, key=lambda p: p.stat().st_mtime, reverse=True)[0]

    raise FileNotFoundError(
        f"작업폴더에서 ContractList.xlsx를 찾지 못했습니다: {BASE_DIR}"
    )


def find_source_file(files: Sequence[Path], base_stem: str) -> Path:
    target = base_stem.casefold()

    exact_matches = [path for path in files if path.stem.casefold() == target]
    if exact_matches:
        return sorted(exact_matches, key=lambda p: p.stat().st_mtime, reverse=True)[0]

    # (1), _복사본 등이 붙은 경우도 허용한다.
    prefix_matches = [
        path
        for path in files
        if path.stem.casefold().startswith(target)
    ]

    if prefix_matches:
        return sorted(prefix_matches, key=lambda p: p.stat().st_mtime, reverse=True)[0]

    raise FileNotFoundError(
        f"'{base_stem}.xlsx' 파일을 작업폴더에서 찾지 못했습니다."
    )


# ============================================================
# 헤더 찾기
# ============================================================

def find_header_row_and_columns(
    worksheet,
    required_header_keys: Set[str],
    max_scan_rows: int = HEADER_SCAN_ROWS,
) -> Optional[Tuple[int, Dict[str, int]]]:
    """
    required_header_keys 중 하나 이상이 존재하는 헤더행을 찾는다.
    반환값: (헤더행 번호, 정규화된 헤더명 -> 열번호)
    """
    last_scan_row = min(worksheet.max_row, max_scan_rows)

    for row_number in range(1, last_scan_row + 1):
        header_map: Dict[str, int] = {}

        for column_number in range(1, worksheet.max_column + 1):
            key = normalize_header(
                worksheet.cell(row=row_number, column=column_number).value
            )
            if key and key not in header_map:
                header_map[key] = column_number

        if any(key in header_map for key in required_header_keys):
            return row_number, header_map

    return None


def find_column(header_map: Dict[str, int], aliases: Set[str]) -> Optional[int]:
    for alias in aliases:
        if alias in header_map:
            return header_map[alias]
    return None


def last_nonempty_header_column(worksheet, header_row: int) -> int:
    last_column = 0
    for column_number in range(1, worksheet.max_column + 1):
        if clean_text(worksheet.cell(header_row, column_number).value):
            last_column = column_number
    return last_column


# ============================================================
# 살아있는 계약 목록 읽기
# ============================================================

def infer_product_from_sheet_name(sheet_name: str) -> str:
    """
    ContractList에 상품명 열이 없는 경우 시트명을 상품명으로 사용한다.
    단, 'Sheet1', '통합', '목록' 같은 일반 시트명은 상품명으로 쓰지 않는다.
    """
    name = clean_text(sheet_name)
    generic_names = {
        "sheet",
        "sheet1",
        "시트",
        "시트1",
        "통합",
        "목록",
        "contractlist",
        "계약목록",
        "계약리스트",
    }
    if normalize_header(name) in {normalize_header(x) for x in generic_names}:
        return ""
    return name


def load_active_contracts(contract_file: Path) -> Dict[Tuple[str, str], str]:
    """
    ContractList.xlsx의 A/B/C열을 직접 읽는다.

      A열: 상품명
      B열: 이름
      C열: 계좌번호

    반환값:
      {(정규화 계좌번호, 정규화 이름): 상품명}

    중요:
    - A열 상품명이 병합셀 또는 그룹별 1회만 입력된 경우를 고려해
      위쪽의 마지막 상품명을 아래 행으로 자동 이어받는다.
    - 같은 계좌번호+이름 조합이 여러 번 나오더라도,
      빈 상품명이 기존의 정상 상품명을 덮어쓰지 못하게 한다.
    """
    active_contracts: Dict[Tuple[str, str], str] = {}

    workbook = load_workbook(contract_file, data_only=True, read_only=False)
    try:
        for worksheet in workbook.worksheets:
            last_product_name = ""

            for row_number in range(1, worksheet.max_row + 1):
                raw_product = worksheet.cell(row_number, 1).value
                raw_name = worksheet.cell(row_number, 2).value
                raw_account = worksheet.cell(row_number, 3).value

                # A/B/C 헤더행 제외
                if (
                    normalize_header(raw_name) in CONTRACT_NAME_HEADER_KEYS
                    and normalize_header(raw_account) in ACCOUNT_HEADER_KEYS
                ):
                    continue

                current_product_name = clean_text(raw_product)

                # 병합셀/반복 생략 구조 대응: A열이 비어 있으면 직전 상품명 사용
                if current_product_name:
                    last_product_name = current_product_name
                product_name = current_product_name or last_product_name

                name_key = normalize_name(raw_name)
                account_key = normalize_account(raw_account)

                if not account_key or not name_key:
                    continue

                contract_key = (account_key, name_key)

                # 동일 키가 반복될 때 빈 상품명으로 정상값을 덮어쓰지 않는다.
                if product_name:
                    active_contracts[contract_key] = product_name
                elif contract_key not in active_contracts:
                    active_contracts[contract_key] = ""
    finally:
        workbook.close()

    if not active_contracts:
        raise ValueError(
            f"{contract_file.name}의 A/B/C열에서 활성 계약을 찾지 못했습니다. "
            "A=상품명, B=이름, C=계좌번호인지 확인해 주세요."
        )

    return active_contracts


# ============================================================
# 거래내역 필터링
# ============================================================

def read_filtered_rows(
    source_file: Path,
    active_contracts: Dict[Tuple[str, str], str],
) -> Tuple[List, List[List], Set[Tuple[str, str]]]:
    """
    반환값:
      headers        : 출력 헤더
      filtered_rows  : 살아있는 계약에 해당하는 거래행
      matched        : 실제로 매칭된 (계좌번호, 이름) 키

    ContractList.xlsx의 A열 상품명을 정리 파일의 PF명 열에 넣는다.
    원본 정리 파일에 PF명 열이 없으면 계약자명 다음에 새로 만든다.
    """
    headers: Optional[List] = None
    filtered_rows: List[List] = []
    matched_accounts: Set[Tuple[str, str]] = set()

    workbook = load_workbook(source_file, data_only=True, read_only=False)
    try:
        for worksheet in workbook.worksheets:
            found = find_header_row_and_columns(
                worksheet,
                ACCOUNT_HEADER_KEYS,
            )
            if found is None:
                # 검토필요 같은 보조 시트는 자동으로 건너뛴다.
                continue

            header_row, header_map = found
            account_column = find_column(header_map, ACCOUNT_HEADER_KEYS)
            source_pf_column = find_column(header_map, PF_HEADER_KEYS)
            contract_name_column = find_column(
                header_map,
                CONTRACT_NAME_HEADER_KEYS,
            )

            if account_column is None:
                continue
            if contract_name_column is None:
                raise ValueError(
                    f"{source_file.name}의 '{worksheet.title}' 시트에서 "
                    "계약자명/이름 열을 찾지 못했습니다."
                )

            last_column = last_nonempty_header_column(worksheet, header_row)
            if last_column == 0:
                continue

            source_headers = [
                worksheet.cell(header_row, column_number).value
                for column_number in range(1, last_column + 1)
            ]

            # PF명 열이 없는 파일도 동일한 출력 구조로 맞춘다.
            if source_pf_column is not None:
                current_headers = list(source_headers)
                pf_output_index = source_pf_column - 1
                insert_pf_column = False
            else:
                current_headers = list(source_headers)
                if contract_name_column is not None:
                    pf_output_index = contract_name_column
                else:
                    pf_output_index = account_column
                current_headers.insert(pf_output_index, "PF명")
                insert_pf_column = True

            if headers is None:
                headers = current_headers
            elif [normalize_header(x) for x in headers] != [
                normalize_header(x) for x in current_headers
            ]:
                raise ValueError(
                    f"{source_file.name}의 시트별 헤더 구성이 서로 다릅니다. "
                    f"문제 시트: {worksheet.title}"
                )

            for row_number in range(header_row + 1, worksheet.max_row + 1):
                raw_account = worksheet.cell(row_number, account_column).value
                raw_name = worksheet.cell(row_number, contract_name_column).value
                account_key = normalize_account(raw_account)
                name_key = normalize_name(raw_name)
                contract_key = (account_key, name_key)

                if (
                    not account_key
                    or not name_key
                    or contract_key not in active_contracts
                ):
                    continue

                row_values = [
                    worksheet.cell(row_number, column_number).value
                    for column_number in range(1, last_column + 1)
                ]

                # 계좌번호가 숫자로 저장돼도 결과에서는 일반 텍스트로 저장한다.
                row_values[account_column - 1] = clean_text(raw_account)

                product_name = clean_text(active_contracts.get(contract_key, ""))
                if not product_name:
                    raise ValueError(
                        "ContractList.xlsx에서 "
                        f"상품명을 찾지 못했습니다: 이름={clean_text(raw_name)}, "
                        f"계좌번호={clean_text(raw_account)}. "
                        "A열 상품명이 병합되었거나 비어 있는지 확인해 주세요."
                    )

                if insert_pf_column:
                    row_values.insert(pf_output_index, product_name)
                else:
                    row_values[pf_output_index] = product_name

                filtered_rows.append(row_values)
                matched_accounts.add(contract_key)
    finally:
        workbook.close()

    if headers is None:
        raise ValueError(
            f"{source_file.name}에서 '계좌번호' 헤더가 있는 거래내역 시트를 찾지 못했습니다."
        )

    return headers, filtered_rows, matched_accounts


# ============================================================
# 결과 엑셀 작성
# ============================================================

def apply_sheet_format(worksheet, headers: Sequence, row_count: int) -> None:
    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(color="FFFFFF", bold=True)
    thin_gray = Side(style="thin", color="D9E2F3")
    border = Border(bottom=thin_gray)

    worksheet.freeze_panes = "A2"

    if headers:
        worksheet.auto_filter.ref = (
            f"A1:{get_column_letter(len(headers))}{max(row_count + 1, 1)}"
        )

    for cell in worksheet[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    header_keys = [normalize_header(value) for value in headers]

    for column_number, header_key in enumerate(header_keys, start=1):
        column_letter = get_column_letter(column_number)

        if header_key in ACCOUNT_HEADER_KEYS:
            worksheet.column_dimensions[column_letter].width = 20
            for row_number in range(2, worksheet.max_row + 1):
                worksheet.cell(row_number, column_number).number_format = "@"

        elif "일자" in header_key or "날짜" in header_key:
            worksheet.column_dimensions[column_letter].width = 13
            for row_number in range(2, worksheet.max_row + 1):
                worksheet.cell(row_number, column_number).number_format = "yyyy-mm-dd"

        elif any(
            word in header_key
            for word in ("수량", "단가", "금액", "수수료", "세금")
        ):
            worksheet.column_dimensions[column_letter].width = 16
            for row_number in range(2, worksheet.max_row + 1):
                worksheet.cell(row_number, column_number).number_format = "#,##0.00"

        else:
            max_length = len(clean_text(headers[column_number - 1]))
            for row_number in range(2, worksheet.max_row + 1):
                value = worksheet.cell(row_number, column_number).value
                max_length = max(max_length, len(clean_text(value)))
            worksheet.column_dimensions[column_letter].width = min(
                max(max_length + 2, 10),
                30,
            )

    for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row):
        for cell in row:
            cell.border = border
            cell.alignment = Alignment(vertical="center")


def build_output(
    output_path: Path,
    source_results: Sequence[Tuple[str, List, List[List]]],
) -> None:
    workbook = Workbook()
    default_sheet = workbook.active
    workbook.remove(default_sheet)

    for sheet_name, headers, rows in source_results:
        worksheet = workbook.create_sheet(sheet_name)
        worksheet.append(list(headers))

        for row_values in rows:
            worksheet.append(row_values)

        apply_sheet_format(
            worksheet,
            headers,
            len(rows),
        )

    temporary_path = output_path.with_name(
        f".{output_path.stem}_작성중.xlsx"
    )

    if temporary_path.exists():
        temporary_path.unlink()

    try:
        workbook.save(temporary_path)
        workbook.close()
        temporary_path.replace(output_path)
    except PermissionError as exc:
        workbook.close()
        if temporary_path.exists():
            temporary_path.unlink()
        raise PermissionError(
            f"'{output_path.name}' 파일이 Excel에서 열려 있습니다. "
            "파일을 완전히 닫은 뒤 다시 실행해 주세요."
        ) from exc
    except Exception:
        workbook.close()
        if temporary_path.exists():
            temporary_path.unlink()
        raise


# ============================================================
# 메인 실행
# ============================================================

def main() -> None:
    files = workbook_candidates()

    contract_file = find_contract_file(files)
    active_contracts = load_active_contracts(contract_file)

    print(f"계약파일: {contract_file.name}")
    print(f"살아있는 계약 수(계좌번호+이름): {len(active_contracts):,}개")

    missing_product_count = sum(
        1 for product_name in active_contracts.values() if not product_name
    )
    if missing_product_count:
        print(
            f"주의: 상품명을 찾지 못한 활성계좌가 "
            f"{missing_product_count:,}개 있습니다. 해당 PF명은 빈칸으로 남습니다."
        )

    source_results: List[Tuple[str, List, List[List]]] = []
    all_matched_accounts: Set[Tuple[str, str]] = set()

    for sheet_name, base_stem in SOURCE_CONFIG:
        source_file = find_source_file(files, base_stem)
        headers, rows, matched_accounts = read_filtered_rows(
            source_file,
            active_contracts,
        )

        source_results.append((sheet_name, headers, rows))
        all_matched_accounts.update(matched_accounts)

        print(
            f"[{sheet_name}] {source_file.name} -> "
            f"거래 {len(rows):,}건 / 계약 {len(matched_accounts):,}개"
        )

    output_path = BASE_DIR / OUTPUT_FILE_NAME
    build_output(output_path, source_results)

    unmatched_count = len(set(active_contracts) - all_matched_accounts)

    print("-" * 60)
    print(f"완료: {output_path}")
    print(f"5개 증권사 파일에서 매칭되지 않은 활성계약: {unmatched_count:,}개")


if __name__ == "__main__":
    main()
