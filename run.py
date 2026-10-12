"""일임보고서 데이터 정리 실행 메뉴. Python 3.10 이상."""
from pathlib import Path
import argparse
import os
import subprocess
import sys
from datetime import datetime

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / "Scripts"))
from report_paths import period, settings, update_period, trade_files, single_ap_input, contract_input

STAGES = {
    "contracts": ["contract_report_list.py"],
    "trades": ["samsung_trade_normalizer.py", "han_trade_normalizer.py", "kb_trade_normalizer.py",
               "kyobo_trade_normalizer.py", "meritz_trade_normalizer_v2.py"],
    "ap": ["samsung_AP_normalizer.py", "han_AP_normalizer.py", "kb_AP_normalizer.py", "kyobo_AP_normalizer.py"],
    "sum-trades": ["SumTradeList.py"],
    "sum-ap": ["SumAP.py"],
}
LABELS = {"contracts": "계약 대상 추출", "trades": "거래내역 정리", "ap": "AP 정리",
          "sum-trades": "거래내역 취합", "sum-ap": "AP 취합", "all": "전체 실행"}


def preflight(stage):
    errors = []
    for package in ("openpyxl", "xlrd"):
        try:
            __import__(package)
        except ImportError:
            errors.append(f"{package} 미설치: python -m pip install -r requirements.txt")
    try:
        period()
        if stage in ("all", "contracts") and not contract_input().exists():
            errors.append("Input/계약/ContractList.xlsx가 없습니다.")
        if stage in ("all", "trades"):
            for broker in ("삼성", "한투", "KB", "교보", "메리츠"):
                if not trade_files(broker): errors.append(f"Input/거래내역/{broker}에 원본이 없습니다.")
        if stage in ("all", "ap"):
            for broker in ("삼성", "한투", "KB", "교보"):
                single_ap_input(broker)
    except (ValueError, KeyError, OSError) as error:
        errors.append(str(error))
    if errors:
        raise ValueError("\n".join(errors))


def run(stage):
    preflight(stage)
    steps = list(STAGES) if stage == "all" else [stage]
    log_dir = ROOT / "Output" / "Logs"
    log_dir.mkdir(parents=True, exist_ok=True)
    log_path = log_dir / f"실행_{datetime.now():%Y%m%d_%H%M%S_%f}.txt"
    environment = dict(os.environ, PYTHONIOENCODING="utf-8", PYTHONUNBUFFERED="1")
    with log_path.open("w", encoding="utf-8") as log:
        log.write(f"기간: {period()}\nAP 환율 기준: {settings().get('ap_exchange_date') or '파일 내 최신일'}\n")
        for step in steps:
            print(f"\n[{LABELS[step]}]", flush=True)
            for script in STAGES[step]:
                log.write(f"\n[{script}]\n")
                process = subprocess.Popen([sys.executable, str(ROOT / "Scripts" / script)],
                                           cwd=ROOT, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                                           text=True, encoding="utf-8", errors="replace", env=environment)
                for line in process.stdout:
                    print(line, end="", flush=True)
                    log.write(line)
                if process.wait() != 0:
                    raise RuntimeError(f"{script} 실행이 중단되었습니다. 로그: {log_path}")
    print(f"\n실행 완료. 결과: {ROOT / 'Output'}\n실행 기록: {log_path}")
    print("검토필요 시트와 미매칭_검토 시트를 확인하세요. 메리츠 AP는 기존과 같이 수작업 대상입니다.")


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("stage", nargs="?", choices=[*STAGES, "all", "check"])
    args = parser.parse_args()
    if args.stage == "check":
        preflight("all")
        print("필수 파일과 설치 확인 완료. 보고서 기간:", *period())
        return
    if args.stage:
        run(args.stage)
        return
    print("일임보고서 데이터 정리")
    start, end = period()
    print(f"현재 기간: {start} ~ {end}")
    print("1 계약 대상 추출\n2 거래내역 정리\n3 AP 정리\n4 거래내역 취합\n5 AP 취합\n6 전체 실행\n7 기간 변경\n0 종료")
    choice = input("번호: ").strip()
    if choice == "0": return
    if choice == "7":
        from datetime import date
        start = date.fromisoformat(input("시작일 YYYY-MM-DD: ").strip())
        end = date.fromisoformat(input("종료일 YYYY-MM-DD: ").strip())
        if start > end: raise ValueError("시작일이 종료일보다 늦습니다.")
        update_period(start, end)
        print("기간을 저장했습니다. 다시 실행하여 6 전체 실행을 선택하세요.")
        return
    stages = {str(i + 1): value for i, value in enumerate([*STAGES, "all"])}
    if choice not in stages: raise ValueError("메뉴 번호를 확인하세요.")
    run(stages[choice])


if __name__ == "__main__":
    try:
        main()
    except Exception as error:
        print(f"\n오류: {error}", file=sys.stderr)
        sys.exit(1)
