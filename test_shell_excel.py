"""
Excel Shell Mode 실전 테스트 스크립트
테스트용 Excel 파일을 생성하고 기본 명령어를 검증합니다.
"""

import sys
from pathlib import Path

import xlwings as xw

# Windows 터미널 인코딩 문제 해결
if sys.platform == "win32":
    import codecs

    sys.stdout = codecs.getwriter("utf-8")(sys.stdout.detach())
    sys.stderr = codecs.getwriter("utf-8")(sys.stderr.detach())


def create_test_excel():
    """테스트용 Excel 파일 생성"""
    print("=" * 60)
    print("Excel Shell Mode 테스트 파일 생성")
    print("=" * 60)

    try:
        # 새 워크북 생성
        print("\n1. 새 워크북 생성 중...")
        app = xw.App(visible=True)
        wb = app.books.add()

        # 첫 번째 시트 이름 변경 및 데이터 추가
        print("2. Sheet1 설정 중...")
        sheet1 = wb.sheets[0]
        sheet1.name = "TestData"

        # 샘플 데이터 추가
        print("3. 샘플 데이터 추가 중...")
        sheet1.range("A1").value = [["Name", "Age", "City", "Score"]]
        sheet1.range("A2").value = [
            ["Alice", 25, "Seoul", 95],
            ["Bob", 30, "Busan", 87],
            ["Charlie", 35, "Incheon", 92],
            ["David", 28, "Daegu", 88],
            ["Eve", 32, "Gwangju", 91],
        ]

        # Excel Table 생성
        print("4. Excel Table 생성 중...")
        try:
            list_obj = sheet1.api.ListObjects.Add(1, sheet1.range("A1:D6").api, None, 1)
            list_obj.Name = "PeopleTable"
            list_obj.TableStyle = "TableStyleMedium2"
            print("   ✓ Table 생성: PeopleTable")
        except Exception as e:
            print(f"   ⚠ Table 생성 실패: {e}")

        # 두 번째 시트 추가
        print("5. 두 번째 시트 추가 중...")
        sheet2 = wb.sheets.add("SalesData")
        sheet2.range("A1").value = [["Product", "Q1", "Q2", "Q3", "Q4"]]
        sheet2.range("A2").value = [
            ["Product A", 100, 120, 140, 160],
            ["Product B", 80, 90, 100, 110],
            ["Product C", 150, 160, 170, 180],
        ]

        # 두 번째 테이블
        try:
            list_obj2 = sheet2.api.ListObjects.Add(1, sheet2.range("A1:E4").api, None, 1)
            list_obj2.Name = "SalesTable"
            list_obj2.TableStyle = "TableStyleLight9"
            print("   ✓ Table 생성: SalesTable")
        except Exception as e:
            print(f"   ⚠ Table 생성 실패: {e}")

        # 파일 저장
        test_file = Path("test_shell.xlsx").absolute()
        print(f"\n6. 파일 저장 중: {test_file}")
        wb.save(test_file)

        print("\n" + "=" * 60)
        print("✓ 테스트 파일 생성 완료!")
        print("=" * 60)
        print(f"\n파일 경로: {test_file}")
        print(f"시트 수: {len(wb.sheets)}")
        print(f"시트 이름: {[s.name for s in wb.sheets]}")

        print("\n" + "=" * 60)
        print("Excel Shell Mode 테스트 시작")
        print("=" * 60)
        print("\n다음 명령을 실행하세요:")
        print(f'  uv run oa excel shell --file-path "{test_file}"')
        print("\nShell 내부에서 테스트할 명령어:")
        print("  1. show context")
        print("  2. sheets")
        print("  3. workbook-info")
        print("  4. table-list")
        print("  5. use sheet TestData")
        print("  6. range-read --range A1:D3")
        print("  7. use sheet SalesData")
        print("  8. range-read --range A1:E4")
        print("  9. help")
        print("  10. exit")

        print("\n워크북이 열린 상태로 대기 중입니다.")
        print("Excel Shell 테스트 완료 후 이 창에서 Enter를 누르면 파일이 닫힙니다.")
        input("\nPress Enter to close workbook...")

        wb.close()
        app.quit()
        print("✓ 워크북이 닫혔습니다.")

    except Exception as e:
        print(f"\n✗ 에러 발생: {e}")
        import traceback

        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    create_test_excel()
