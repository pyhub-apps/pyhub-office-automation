"""
기본 엔진 검증 스크립트

WindowsEngine이 정상적으로 초기화되고 기본 기능이 동작하는지 확인합니다.
"""

import platform


def test_engine_initialization():
    """엔진 초기화 테스트"""
    from pyhub_office_automation.excel.engines import get_engine, get_platform_name

    print(f"Platform: {get_platform_name()}")

    if platform.system() == "Windows":
        engine = get_engine()
        print(f"Engine type: {type(engine).__name__}")
        print(f"Engine initialized: {engine is not None}")

        # 워크북 목록 조회 (Excel이 실행 중이어야 함)
        try:
            workbooks = engine.get_workbooks()
            print(f"Open workbooks: {len(workbooks)}")
            for wb in workbooks:
                print(f"  - {wb.name} ({wb.sheet_count} sheets)")
        except Exception as e:
            print(f"Note: {e}")
            print("  (Excel을 먼저 실행하고 워크북을 열어주세요)")

        return True
    else:
        print("Windows가 아니므로 WindowsEngine 테스트를 건너뜁니다")
        return True


if __name__ == "__main__":
    success = test_engine_initialization()
    print("\n✅ Engine initialization test passed!" if success else "\n❌ Test failed")
