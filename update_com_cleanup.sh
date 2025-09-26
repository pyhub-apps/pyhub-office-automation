#!/bin/bash

# Script to update COM cleanup in Excel automation files
# This adds basic COM cleanup to files that don't have it yet

files=(
    "pyhub_office_automation/excel/chart_position.py"
    "pyhub_office_automation/excel/data_analyze.py"
    "pyhub_office_automation/excel/data_transform.py"
    "pyhub_office_automation/excel/metadata_generate.py"
    "pyhub_office_automation/excel/pivot_refresh.py"
    "pyhub_office_automation/excel/range_convert.py"
    "pyhub_office_automation/excel/range_read.py"
    "pyhub_office_automation/excel/range_write.py"
    "pyhub_office_automation/excel/shape_add.py"
    "pyhub_office_automation/excel/shape_delete.py"
    "pyhub_office_automation/excel/shape_format.py"
    "pyhub_office_automation/excel/shape_group.py"
    "pyhub_office_automation/excel/shape_list.py"
    "pyhub_office_automation/excel/sheet_activate.py"
    "pyhub_office_automation/excel/slicer_add.py"
    "pyhub_office_automation/excel/slicer_connect.py"
    "pyhub_office_automation/excel/slicer_list.py"
    "pyhub_office_automation/excel/slicer_position.py"
    "pyhub_office_automation/excel/table_analyze.py"
    "pyhub_office_automation/excel/table_create.py"
    "pyhub_office_automation/excel/table_list.py"
    "pyhub_office_automation/excel/table_sort.py"
    "pyhub_office_automation/excel/table_sort_clear.py"
    "pyhub_office_automation/excel/table_sort_info.py"
    "pyhub_office_automation/excel/table_write.py"
    "pyhub_office_automation/excel/textbox_add.py"
    "pyhub_office_automation/excel/workbook_info.py"
)

echo "Updating COM cleanup in ${#files[@]} files..."

for file in "${files[@]}"; do
    if [[ -f "$file" ]]; then
        echo "Processing: $file"

        # Create backup
        cp "$file" "${file}.backup"

        # Pattern 1: Basic finally block pattern
        sed -i 's/    finally:\n        # 워크북 정리 - 활성 워크북이나 이름으로 접근한 경우 앱 종료하지 않음/    finally:\n        # COM 객체 명시적 해제\n        try:\n            # 가비지 컬렉션 강제 실행\n            import gc\n            gc.collect()\n\n            # Windows에서 COM 라이브러리 정리\n            import platform\n            if platform.system() == "Windows":\n                try:\n                    import pythoncom\n                    pythoncom.CoUninitialize()\n                except:\n                    pass\n\n        except:\n            pass\n\n        # 워크북 정리 - 활성 워크북이나 이름으로 접근한 경우 앱 종료하지 않음/' "$file"

        # Pattern 2: Alternative pattern
        sed -i 's/    finally:\n        # 새로 생성한 워크북인 경우에만 정리/    finally:\n        # COM 객체 명시적 해제\n        try:\n            # 가비지 컬렉션 강제 실행\n            import gc\n            gc.collect()\n\n            # Windows에서 COM 라이브러리 정리\n            import platform\n            if platform.system() == "Windows":\n                try:\n                    import pythoncom\n                    pythoncom.CoUninitialize()\n                except:\n                    pass\n\n        except:\n            pass\n\n        # 새로 생성한 워크북인 경우에만 정리/' "$file"

        echo "  Updated: $file"
    else
        echo "  File not found: $file"
    fi
done

echo "COM cleanup update completed!"