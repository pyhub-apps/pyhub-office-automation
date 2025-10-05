"""
Issue #88 마이그레이션 명령어 테스트 스크립트

Windows 환경에서 18개 마이그레이션된 명령어 기능 검증
"""

import json
import os
import subprocess
import sys
from pathlib import Path

# 테스트 결과 저장
test_results = {"total": 0, "passed": 0, "failed": 0, "skipped": 0, "details": []}


def run_command(cmd_list, description):
    """명령어 실행 및 결과 검증"""
    global test_results
    test_results["total"] += 1

    print(f"\n{'='*80}")
    print(f"테스트: {description}")
    print(f"명령어: {' '.join(cmd_list)}")
    print(f"{'='*80}")

    try:
        result = subprocess.run(cmd_list, capture_output=True, text=True, timeout=30, encoding="utf-8")

        if result.returncode == 0:
            # JSON 파싱 시도
            try:
                output = json.loads(result.stdout)
                success = output.get("success", False)

                if success:
                    print(f"[PASSED] {description}")
                    test_results["passed"] += 1
                    test_results["details"].append(
                        {"test": description, "status": "PASSED", "message": output.get("message", "Success")}
                    )
                else:
                    print(f"[FAILED] {description}")
                    print(f"Error: {output.get('error', {}).get('message', 'Unknown error')}")
                    test_results["failed"] += 1
                    test_results["details"].append(
                        {
                            "test": description,
                            "status": "FAILED",
                            "error": output.get("error", {}).get("message", "Unknown error"),
                        }
                    )
            except json.JSONDecodeError:
                # JSON 파싱 실패 - stdout 출력 확인
                if result.stdout:
                    print(f"[PASSED] (non-JSON output): {description}")
                    test_results["passed"] += 1
                else:
                    print(f"[FAILED] No output")
                    test_results["failed"] += 1
        else:
            print(f"[FAILED] {description}")
            print(f"Return code: {result.returncode}")
            print(f"Error: {result.stderr}")
            test_results["failed"] += 1
            test_results["details"].append({"test": description, "status": "FAILED", "error": result.stderr})

    except subprocess.TimeoutExpired:
        print(f"[TIMEOUT] {description}")
        test_results["failed"] += 1
        test_results["details"].append({"test": description, "status": "FAILED", "error": "Timeout (30s)"})

    except Exception as e:
        print(f"[ERROR] {description}")
        print(f"Exception: {str(e)}")
        test_results["failed"] += 1
        test_results["details"].append({"test": description, "status": "ERROR", "error": str(e)})


def setup_test_environment():
    """테스트 환경 준비"""
    print("\n" + "=" * 80)
    print("테스트 환경 준비")
    print("=" * 80)

    # Python 경로 확인
    python_exe = r"C:\Users\pyhub\AppData\Local\Programs\Python\Python313\python.EXE"
    if not Path(python_exe).exists():
        print(f"[ERROR] Python not found at {python_exe}")
        print("[INFO] Using system python instead")
        python_exe = sys.executable

    print(f"[OK] Using Python: {python_exe}")

    # 버전 확인
    result = subprocess.run(
        [python_exe, "-m", "pyhub_office_automation.cli.main", "--version"], capture_output=True, text=True, encoding="utf-8"
    )
    print(f"[INFO] Version: {result.stdout.strip()}")

    # 테스트 데이터 디렉토리 생성
    test_data_dir = Path("test-data")
    test_data_dir.mkdir(exist_ok=True)
    print(f"[OK] Test data directory: {test_data_dir}")

    return True


def test_table_commands():
    """Table Commands 테스트 (4개)"""
    print("\n" + "#" * 80)
    print("# Table Commands 테스트 (4개)")
    print("#" * 80)

    python = r"C:\Users\pyhub\AppData\Local\Programs\Python\Python313\python.EXE"

    # Excel 파일이 열려있다고 가정
    # 실제 테스트는 수동으로 Excel 파일 열고 진행

    print("\n[WARNING] Table 명령어 테스트는 Excel 파일이 열려있어야 합니다.")
    print("수동 테스트 권장:")
    print("1. Excel에서 빈 워크북 열기")
    print("2. Sheet1에 A1:D10 범위에 데이터 입력")
    print("3. 아래 명령어 실행:")
    print(f"   {python} -m pyhub_office_automation.cli.main excel table-create --range A1:D10 --table-name TestTable")
    print(f"   {python} -m pyhub_office_automation.cli.main excel table-sort --table-name TestTable --column A --order asc")
    print(f"   {python} -m pyhub_office_automation.cli.main excel table-sort-info --table-name TestTable")
    print(f"   {python} -m pyhub_office_automation.cli.main excel table-sort-clear --table-name TestTable")

    test_results["skipped"] += 4


def test_slicer_commands():
    """Slicer Commands 테스트 (4개)"""
    print("\n" + "#" * 80)
    print("# Slicer Commands 테스트 (4개)")
    print("#" * 80)

    python = r"C:\Users\pyhub\AppData\Local\Programs\Python\Python313\python.EXE"

    print("\n[WARNING] Slicer 명령어 테스트는 피벗테이블이 있는 Excel 파일이 필요합니다.")
    print("수동 테스트 권장:")
    print("1. Excel에서 피벗테이블 포함된 파일 열기")
    print("2. 아래 명령어 실행:")
    print(f"   {python} -m pyhub_office_automation.cli.main excel slicer-add --pivot-table PivotTable1 --field Region")
    print(f"   {python} -m pyhub_office_automation.cli.main excel slicer-list")
    print(
        f"   {python} -m pyhub_office_automation.cli.main excel slicer-position --slicer-name Slicer_Region --left 500 --top 100"
    )
    print(f"   {python} -m pyhub_office_automation.cli.main excel slicer-connect --slicer-name Slicer_Region --action list")

    test_results["skipped"] += 4


def test_pivot_commands():
    """Pivot Commands 테스트 (5개)"""
    print("\n" + "#" * 80)
    print("# Pivot Commands 테스트 (5개)")
    print("#" * 80)

    python = r"C:\Users\pyhub\AppData\Local\Programs\Python\Python313\python.EXE"

    print("\n[WARNING] Pivot 명령어 테스트는 데이터가 있는 Excel 파일이 필요합니다.")
    print("수동 테스트 권장:")
    print("1. Excel에서 데이터 시트 포함된 파일 열기")
    print("2. 아래 명령어 실행:")
    print(f"   {python} -m pyhub_office_automation.cli.main excel pivot-create --source-range A1:D100 --dest-range F1")
    print(f"   {python} -m pyhub_office_automation.cli.main excel pivot-list")
    print(
        f"   {python} -m pyhub_office_automation.cli.main excel pivot-configure --pivot-name PivotTable1 --row-fields Region"
    )
    print(f"   {python} -m pyhub_office_automation.cli.main excel pivot-refresh --pivot-name PivotTable1")
    print(f"   {python} -m pyhub_office_automation.cli.main excel pivot-delete --pivot-name PivotTable1")

    test_results["skipped"] += 5


def test_shape_commands():
    """Shape Commands 테스트 (5개)"""
    print("\n" + "#" * 80)
    print("# Shape Commands 테스트 (5개)")
    print("#" * 80)

    python = r"C:\Users\pyhub\AppData\Local\Programs\Python\Python313\python.EXE"

    print("\n[WARNING] Shape 명령어 테스트는 Excel 파일이 열려있어야 합니다.")
    print("수동 테스트 권장:")
    print("1. Excel에서 빈 워크북 열기")
    print("2. 아래 명령어 실행:")
    print(
        f"   {python} -m pyhub_office_automation.cli.main excel shape-add --shape-type rectangle --left 100 --top 100 --width 200 --height 100"
    )
    print(f"   {python} -m pyhub_office_automation.cli.main excel shape-list")
    print(f"   {python} -m pyhub_office_automation.cli.main excel shape-format --shape-name Rectangle1 --fill-color FF0000")
    print(f"   {python} -m pyhub_office_automation.cli.main excel shape-group --shapes Rectangle1,Oval1 --group-name MyGroup")
    print(f"   {python} -m pyhub_office_automation.cli.main excel shape-delete --shapes Rectangle1")

    test_results["skipped"] += 5


def test_basic_engine_functionality():
    """기본 Engine 기능 테스트 (자동화 가능)"""
    print("\n" + "#" * 80)
    print("# 기본 Engine 기능 테스트")
    print("#" * 80)

    python = r"C:\Users\pyhub\AppData\Local\Programs\Python\Python313\python.EXE"

    # Excel 도움말 명령어 테스트
    run_command([python, "-m", "pyhub_office_automation.cli.main", "excel", "--help"], "Excel 명령어 도움말")

    # 워크북 목록 조회 (Excel이 실행 중이지 않으면 빈 결과)
    run_command([python, "-m", "pyhub_office_automation.cli.main", "excel", "workbook-list"], "열린 워크북 목록 조회")


def print_summary():
    """테스트 결과 요약 출력"""
    print("\n" + "=" * 80)
    print("테스트 결과 요약")
    print("=" * 80)

    print(f"\n총 테스트: {test_results['total']}")
    print(f"[PASSED] 성공: {test_results['passed']}")
    print(f"[FAILED] 실패: {test_results['failed']}")
    print(f"[SKIPPED] 건너뜀: {test_results['skipped']}")

    pass_rate = (test_results["passed"] / test_results["total"] * 100) if test_results["total"] > 0 else 0
    print(f"\n성공률: {pass_rate:.1f}%")

    if test_results["details"]:
        print("\n상세 결과:")
        for detail in test_results["details"]:
            status_icon = "[OK]" if detail["status"] == "PASSED" else "[FAIL]"
            print(f"{status_icon} {detail['test']}: {detail['status']}")
            if "error" in detail:
                print(f"   Error: {detail['error']}")

    print("\n" + "=" * 80)
    print("Issue #88 마이그레이션 검증 완료")
    print("=" * 80)
    print("\n[INFO] 수동 테스트 가이드:")
    print("   - 18개 마이그레이션 명령어는 실제 Excel 파일이 필요합니다")
    print("   - 위에 출력된 수동 테스트 명령어를 Excel 파일과 함께 실행하세요")
    print("   - 각 명령어가 JSON 응답을 반환하고 success: true인지 확인하세요")
    print("\n[CHECKLIST] 테스트 체크리스트:")
    print("   [ ] Table 4개 명령어 테스트")
    print("   [ ] Slicer 4개 명령어 테스트")
    print("   [ ] Pivot 5개 명령어 테스트")
    print("   [ ] Shape 5개 명령어 테스트")


def main():
    """메인 테스트 실행"""
    print("\n" + "=" * 80)
    print("Issue #88 마이그레이션 명령어 테스트")
    print("Windows Engine Layer 기능 검증")
    print("=" * 80)

    # 환경 준비
    if not setup_test_environment():
        print("\n[ERROR] 테스트 환경 준비 실패")
        sys.exit(1)

    # 기본 기능 테스트
    test_basic_engine_functionality()

    # 각 카테고리별 테스트
    test_table_commands()
    test_slicer_commands()
    test_pivot_commands()
    test_shape_commands()

    # 결과 요약
    print_summary()

    # 종료 코드 반환
    if test_results["failed"] > 0:
        sys.exit(1)
    else:
        sys.exit(0)


if __name__ == "__main__":
    main()
