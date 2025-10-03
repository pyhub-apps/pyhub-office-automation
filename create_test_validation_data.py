"""Create test Excel file with data quality issues for Issue #90 validation testing"""

import sys

import xlwings as xw

# Fix encoding for Windows console
if sys.platform == "win32":
    import io

    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

# Create new workbook
app = xw.App(visible=False)
wb = app.books.add()
sht = wb.sheets[0]
sht.name = "ValidationTest"

# Create test data with various quality issues
test_data = [
    # Headers
    ["이름", "이메일", "나이", "가격", "날짜", "회원ID"],
    # Clean data
    ["홍길동", "hong@example.com", 25, 10000.5, "2024-01-15", "USER001"],
    ["김철수", "kim@example.com", 30, 20000.0, "2024-02-20", "USER002"],
    # NULL values
    ["이영희", None, 28, 15000.0, "2024-03-10", "USER003"],
    [None, "park@example.com", 35, None, "2024-04-05", "USER004"],
    # Empty strings
    ["최민수", "", 32, 18000.0, "2024-05-12", "USER005"],
    # Whitespace-only strings
    ["정수진", "   ", 27, 22000.0, "2024-06-18", "USER006"],
    # Duplicate rows (same as row 2)
    ["홍길동", "hong@example.com", 25, 10000.5, "2024-01-15", "USER001"],
    # Duplicate key column (회원ID)
    ["박영수", "park2@example.com", 29, 17000.0, "2024-07-22", "USER002"],
    # Type mismatches
    ["강민지", "kang@example.com", "스물여덟", 19000.0, "2024-08-15", "USER007"],
    ["송하늘", "song@example.com", 31, "일만원", "2024-09-20", "USER008"],
    ["윤서연", "yoon@example.com", 26, 16000.0, "Invalid Date", "USER009"],
    # More clean data
    ["임동혁", "lim@example.com", 33, 21000.0, "2024-10-10", "USER010"],
    ["조미래", "jo@example.com", 24, 14000.0, "2024-11-05", "USER011"],
]

# Write data to Excel
sht.range("A1").value = test_data

# Auto-fit columns
sht.autofit()

# Save file
import os

save_path = os.path.abspath("test_validation_data.xlsx")
wb.save(save_path)
print(f"✅ Test data created: {save_path}")

# Print data summary
print(f"\nData Summary:")
print(f"- Total rows: {len(test_data) - 1} (excluding header)")
print(f"- Columns: {len(test_data[0])}")
print(f"\nData Quality Issues:")
print(f"- NULL values: 2 cells (이름, 나이)")
print(f"- Empty strings: 1 cell (이메일)")
print(f"- Whitespace-only: 1 cell (이메일)")
print(f"- Duplicate rows: 1 (row 2 = row 8)")
print(f"- Duplicate key (회원ID): 1 (USER002 appears twice)")
print(f"- Type errors:")
print(f"  - 나이: '스물여덟' (should be int)")
print(f"  - 가격: '일만원' (should be float)")
print(f"  - 날짜: 'Invalid Date' (should be date)")

# Close workbook
wb.close()
app.quit()

print("\n✅ Ready for testing!")
print("\nTest commands:")
print("1. All checks:")
print('   uv run oa excel data-validate --file-path "test_validation_data.xlsx" --range "A1:F14" --format text')
print("\n2. NULL check with required columns:")
print(
    '   uv run oa excel data-validate --file-path "test_validation_data.xlsx" --range "A1:F14" --checks null --required-columns "이름,이메일,회원ID"'
)
print("\n3. Duplicate check with key columns:")
print(
    '   uv run oa excel data-validate --file-path "test_validation_data.xlsx" --range "A1:F14" --checks duplicate --key-columns "회원ID"'
)
print("\n4. Type validation:")
print(
    '   uv run oa excel data-validate --file-path "test_validation_data.xlsx" --range "A1:F14" --checks type --column-types "나이:int,가격:float,날짜:date"'
)
print("\n5. JSON output:")
print('   uv run oa excel data-validate --file-path "test_validation_data.xlsx" --range "A1:F14" --format json')
