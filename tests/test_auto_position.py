#!/usr/bin/env python3
"""
자동 배치 기능 테스트 스크립트
새로 추가한 범위 관리 유틸리티 함수들의 기본 동작을 테스트합니다.
"""

import sys

sys.path.insert(0, "/Users/allieus/Apps/pyhub-office-automation")

from pyhub_office_automation.excel.utils import (
    check_range_overlap,
    coords_to_excel_address,
    estimate_pivot_table_size,
    excel_address_to_coords,
    parse_excel_range,
)


def test_excel_address_conversion():
    """Excel 주소 변환 함수 테스트"""
    print("=== Excel 주소 변환 테스트 ===")

    # 좌표 → 주소 변환 테스트
    test_cases = [
        (1, 1, "A1"),
        (1, 26, "Z1"),
        (1, 27, "AA1"),
        (5, 10, "J5"),
        (100, 52, "AZ100"),
    ]

    for row, col, expected in test_cases:
        result = coords_to_excel_address(row, col)
        status = "✅" if result == expected else "❌"
        print(f"{status} ({row}, {col}) → {result} (기대값: {expected})")

    # 주소 → 좌표 변환 테스트
    print("\n주소 → 좌표 변환:")
    for row, col, address in test_cases:
        result_row, result_col = excel_address_to_coords(address)
        status = "✅" if (result_row, result_col) == (row, col) else "❌"
        print(f"{status} {address} → ({result_row}, {result_col}) (기대값: ({row}, {col}))")


def test_range_parsing():
    """범위 파싱 함수 테스트"""
    print("\n=== 범위 파싱 테스트 ===")

    test_cases = [
        ("A1:C10", (1, 1, 10, 3)),
        ("B5:D15", (5, 2, 15, 4)),
        ("Z1:AA5", (1, 26, 5, 27)),
        ("A1", (1, 1, 1, 1)),  # 단일 셀
    ]

    for range_str, expected in test_cases:
        result = parse_excel_range(range_str)
        status = "✅" if result == expected else "❌"
        print(f"{status} {range_str} → {result} (기대값: {expected})")


def test_range_overlap():
    """범위 겹침 검사 함수 테스트"""
    print("\n=== 범위 겹침 검사 테스트 ===")

    test_cases = [
        ("A1:C10", "B5:D15", True),  # 겹침
        ("A1:C10", "D1:F10", False),  # 겹치지 않음
        ("A1:C10", "A1:C10", True),  # 완전히 동일
        ("A1:C10", "B2:B5", True),  # 포함 관계
        ("F1:H10", "A1:E10", False),  # 완전히 분리
    ]

    for range1, range2, expected in test_cases:
        result = check_range_overlap(range1, range2)
        status = "✅" if result == expected else "❌"
        print(f"{status} {range1} vs {range2} → {result} (기대값: {expected})")


def test_pivot_size_estimation():
    """피벗 테이블 크기 추정 함수 테스트"""
    print("\n=== 피벗 테이블 크기 추정 테스트 ===")

    test_cases = [
        ("A1:D100", 3, "100행 4열 데이터"),
        ("A1:K50", 5, "50행 11열 데이터"),
        ("A1:C20", 2, "20행 3열 데이터"),
    ]

    for range_str, field_count, description in test_cases:
        result = estimate_pivot_table_size(range_str, field_count)
        print(f"📊 {description} (필드 {field_count}개) → 예상 크기: {result[0]}열 × {result[1]}행")


def run_all_tests():
    """모든 테스트 실행"""
    print("🧪 자동 배치 기능 유틸리티 함수 테스트 시작\n")

    try:
        test_excel_address_conversion()
        test_range_parsing()
        test_range_overlap()
        test_pivot_size_estimation()

        print("\n✅ 모든 테스트가 완료되었습니다!")
        print("💡 실제 Excel 파일에서의 테스트는 Windows 환경에서 수행하세요.")

    except Exception as e:
        print(f"\n❌ 테스트 중 오류 발생: {str(e)}")
        return False

    return True


if __name__ == "__main__":
    success = run_all_tests()
    sys.exit(0 if success else 1)
