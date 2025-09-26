"""
COM 리소스 관리 테스트 실행기
모든 COM 관련 테스트를 실행하고 결과를 정리하는 스크립트
"""

import os
import subprocess
import sys
import time
from pathlib import Path
from typing import Dict, List, Tuple


def run_test_file(test_file: str) -> Tuple[bool, str, float]:
    """개별 테스트 파일 실행"""
    start_time = time.time()

    try:
        result = subprocess.run(
            [sys.executable, "-m", "pytest", test_file, "-v", "--tb=short"], capture_output=True, text=True, timeout=300
        )  # 5분 타임아웃

        execution_time = time.time() - start_time
        success = result.returncode == 0

        output = result.stdout + result.stderr
        return success, output, execution_time

    except subprocess.TimeoutExpired:
        execution_time = time.time() - start_time
        return False, f"테스트 타임아웃 (5분 초과): {test_file}", execution_time

    except Exception as e:
        execution_time = time.time() - start_time
        return False, f"테스트 실행 실패: {str(e)}", execution_time


def run_all_com_tests():
    """모든 COM 테스트 실행"""
    test_files = [
        "test_com_resource_manager.py",
        "test_utils_timeout.py",
        "test_excel_com_integration.py",
        "test_com_performance_memory.py",
        "test_com_edge_cases.py",
    ]

    test_directory = Path(__file__).parent
    results = {}

    print("=" * 80)
    print("COM 리소스 관리 테스트 실행기")
    print("=" * 80)
    print()

    total_start_time = time.time()

    for test_file in test_files:
        test_path = test_directory / test_file

        if not test_path.exists():
            print(f"[WARNING] 테스트 파일 없음: {test_file}")
            results[test_file] = (False, "파일 없음", 0.0)
            continue

        print(f"[TEST] 실행 중: {test_file}")
        print("-" * 60)

        success, output, execution_time = run_test_file(str(test_path))
        results[test_file] = (success, output, execution_time)

        if success:
            print(f"[SUCCESS] 성공 ({execution_time:.2f}초)")
        else:
            print(f"[FAILED] 실패 ({execution_time:.2f}초)")
            print("에러 출력:")
            print(output)

        print()

    total_execution_time = time.time() - total_start_time

    # 결과 요약
    print("=" * 80)
    print("테스트 결과 요약")
    print("=" * 80)

    successful_tests = []
    failed_tests = []

    for test_file, (success, output, execution_time) in results.items():
        if success:
            successful_tests.append((test_file, execution_time))
        else:
            failed_tests.append((test_file, execution_time))

    print(f"[SUCCESS] 성공한 테스트: {len(successful_tests)}")
    for test_file, execution_time in successful_tests:
        print(f"   - {test_file} ({execution_time:.2f}초)")

    print()
    print(f"[FAILED] 실패한 테스트: {len(failed_tests)}")
    for test_file, execution_time in failed_tests:
        print(f"   - {test_file} ({execution_time:.2f}초)")

    print()
    print(f"[STATS] 전체 실행 시간: {total_execution_time:.2f}초")
    print(f"[STATS] 성공률: {len(successful_tests)}/{len(test_files)} ({len(successful_tests)/len(test_files)*100:.1f}%)")

    return len(failed_tests) == 0


def run_specific_test_category(category: str):
    """특정 카테고리의 테스트만 실행"""
    category_mapping = {
        "unit": ["test_com_resource_manager.py"],
        "timeout": ["test_utils_timeout.py"],
        "integration": ["test_excel_com_integration.py"],
        "performance": ["test_com_performance_memory.py"],
        "edge": ["test_com_edge_cases.py"],
    }

    if category not in category_mapping:
        print(f"[ERROR] 알 수 없는 카테고리: {category}")
        print(f"사용 가능한 카테고리: {', '.join(category_mapping.keys())}")
        return False

    test_files = category_mapping[category]
    test_directory = Path(__file__).parent

    print(f"[TARGET] {category.upper()} 테스트 실행")
    print("=" * 60)

    all_success = True

    for test_file in test_files:
        test_path = test_directory / test_file

        if not test_path.exists():
            print(f"[WARNING] 테스트 파일 없음: {test_file}")
            all_success = False
            continue

        print(f"[TEST] 실행 중: {test_file}")
        success, output, execution_time = run_test_file(str(test_path))

        if success:
            print(f"[SUCCESS] 성공 ({execution_time:.2f}초)")
        else:
            print(f"[FAILED] 실패 ({execution_time:.2f}초)")
            print("상세 출력:")
            print(output)
            all_success = False

        print()

    return all_success


def run_coverage_analysis():
    """코드 커버리지 분석 실행"""
    print("[COVERAGE] 코드 커버리지 분석 실행")
    print("=" * 60)

    try:
        # pytest-cov가 설치되어 있는지 확인
        subprocess.run([sys.executable, "-c", "import pytest_cov"], check=True, capture_output=True)
    except (subprocess.CalledProcessError, ImportError):
        print("[WARNING] pytest-cov가 설치되지 않음. 커버리지 분석을 건너뜁니다.")
        print("설치 명령: pip install pytest-cov")
        return False

    test_directory = Path(__file__).parent
    project_root = test_directory.parent

    coverage_command = [
        sys.executable,
        "-m",
        "pytest",
        str(test_directory / "test_com_*.py"),
        "--cov=" + str(project_root / "pyhub_office_automation" / "excel"),
        "--cov-report=term-missing",
        "--cov-report=html:htmlcov",
        "--cov-fail-under=80",
        "-v",
    ]

    try:
        result = subprocess.run(coverage_command, timeout=600)  # 10분 타임아웃

        if result.returncode == 0:
            print("[SUCCESS] 커버리지 분석 완료")
            print("[REPORT] HTML 리포트: htmlcov/index.html")
            return True
        else:
            print("[FAILED] 커버리지 분석 실패 또는 커버리지 부족")
            return False

    except subprocess.TimeoutExpired:
        print("[TIMEOUT] 커버리지 분석 타임아웃")
        return False


def main():
    """메인 실행 함수"""
    if len(sys.argv) > 1:
        command = sys.argv[1]

        if command == "all":
            success = run_all_com_tests()
            sys.exit(0 if success else 1)

        elif command == "coverage":
            success = run_coverage_analysis()
            sys.exit(0 if success else 1)

        elif command in ["unit", "timeout", "integration", "performance", "edge"]:
            success = run_specific_test_category(command)
            sys.exit(0 if success else 1)

        else:
            print("사용법:")
            print("  python run_com_tests.py all           # 모든 테스트 실행")
            print("  python run_com_tests.py unit          # 단위 테스트만 실행")
            print("  python run_com_tests.py timeout       # 타임아웃 테스트만 실행")
            print("  python run_com_tests.py integration   # 통합 테스트만 실행")
            print("  python run_com_tests.py performance   # 성능 테스트만 실행")
            print("  python run_com_tests.py edge          # 엣지 케이스 테스트만 실행")
            print("  python run_com_tests.py coverage      # 커버리지 분석 실행")
            sys.exit(1)

    else:
        # 기본적으로 모든 테스트 실행
        success = run_all_com_tests()
        sys.exit(0 if success else 1)


if __name__ == "__main__":
    main()
