# OA CLI 명령 예시 - 대시보드 생성 작업

## SECTION 1: 샘플 데이터 분석 및 차트 제안

### 데이터 준비 및 읽기

```bash
# 새 워크북 생성
oa excel workbook-create --save-path "게임판매대시보드.xlsx" --name "GameSales"

# 샘플 데이터를 포함한 CSV 파일이 있다면 읽어오기
oa excel table-write --file-path "게임판매대시보드.xlsx" --sheet "Data" --data-file "game_sales.csv" --table-name "GameData"

# 또는 직접 데이터 입력
oa excel range-write --file-path "게임판매대시보드.xlsx" --sheet "Data" --range "A1" --data '[["순위","게임명","플랫폼","발행일","장르","퍼블리셔","북미 판매량","유럽 판매량","일본 판매량","기타 판매량","글로벌 판매량"]]'

# 데이터 읽어서 분석용으로 확인
oa excel table-read --file-path "게임판매대시보드.xlsx" --sheet "Data" --output-file "game_data.json"

# 워크북 정보 확인
oa excel workbook-info --file-path "게임판매대시보드.xlsx" --include-sheets --include-properties
```

## SECTION 2: 차트와 피벗테이블 구성 계획

### 피벗테이블용 시트 생성

```bash
# 피벗 시트 추가
oa excel sheet-add --file-path "게임판매대시보드.xlsx" --name "피벗"

# 대시보드 시트 추가
oa excel sheet-add --file-path "게임판매대시보드.xlsx" --name "대시보드"

# 시트 목록 확인
oa excel workbook-info --file-path "게임판매대시보드.xlsx" --include-sheets
```

## SECTION 3: 대시보드 틀 구성

### VBA 매크로 코드를 텍스트 파일로 준비
```bash
# VBA 코드를 텍스트 파일로 저장 (dashboard_layout.vba)
echo 'Sub CreateDashboardLayout()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("대시보드")

    ' 배경색 설정
    ws.Cells.Interior.Color = RGB(&HF2, &HED, &HF3)

    ' 타이틀 바 추가
    Dim titleBar As Shape
    Set titleBar = ws.Shapes.AddShape(msoShapeRectangle, 10, 10, 800, 60)
    titleBar.Fill.ForeColor.RGB = RGB(&H1D, &H24, &H33)

    ' 차트 박스들 추가
    Dim i As Integer
    For i = 1 To 5
        Dim chartBox As Shape
        If i <= 2 Then
            Set chartBox = ws.Shapes.AddShape(msoShapeRoundedRectangle, _
                10 + (i - 1) * 410, 90, 390, 250)
        Else
            Set chartBox = ws.Shapes.AddShape(msoShapeRoundedRectangle, _
                10 + (i - 3) * 270, 360, 260, 250)
        End If
        chartBox.Name = "ChartBox" & i
        chartBox.Fill.ForeColor.RGB = RGB(&HFF, &HFF, &HFF)
    Next i
End Sub' > dashboard_layout.vba

# VBA 매크로 실행 (xlwings를 통해)
oa excel run-macro --file-path "게임판매대시보드.xlsx" --macro-name "CreateDashboardLayout"
```

## SECTION 4: 피벗테이블 생성

### 피벗테이블 생성을 위한 데이터 준비
```bash
# 원본 데이터가 테이블 형식으로 준비되어 있는지 확인
oa excel table-read --file-path "게임판매대시보드.xlsx" --sheet "Data" --output-file "check_data.json"

# VBA 코드로 피벗테이블 생성 (pivot_tables.vba)
echo 'Sub CreatePivotTables()
    Dim wsData As Worksheet, wsPivot As Worksheet
    Set wsData = ThisWorkbook.Sheets("Data")
    Set wsPivot = ThisWorkbook.Sheets("피벗")

    ' 피벗캐시 생성
    Dim pc As PivotCache
    Set pc = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:="GameData")

    ' 첫 번째 피벗테이블: 글로벌 판매량 TOP5
    Dim pt1 As PivotTable
    Set pt1 = pc.CreatePivotTable( _
        TableDestination:=wsPivot.Range("A3"), _
        TableName:="Pivot_GlobalSales_ByGame")

    With pt1
        .PivotFields("게임명").Orientation = xlRowField
        .PivotFields("글로벌 판매량").Orientation = xlDataField
        .DataFields("합계 - 글로벌 판매량").NumberFormat = "#,##0.00"
    End With
End Sub' > pivot_tables.vba

# 매크로 실행
oa excel run-macro --file-path "게임판매대시보드.xlsx" --macro-name "CreatePivotTables"
```

## SECTION 5: 차트 생성

### 차트 생성 준비
```bash
# 피벗테이블이 생성되었는지 확인
oa excel sheet-activate --file-path "게임판매대시보드.xlsx" --sheet "피벗"

# VBA 코드로 차트 생성 (create_charts.vba)
echo 'Sub CreateCharts()
    Dim wsDash As Worksheet, wsPivot As Worksheet
    Set wsDash = ThisWorkbook.Sheets("대시보드")
    Set wsPivot = ThisWorkbook.Sheets("피벗")

    Dim chartBox As Shape
    Dim chartObj As ChartObject
    Dim i As Integer

    For i = 1 To 5
        Set chartBox = wsDash.Shapes("ChartBox" & i)

        ' 차트 생성
        Set chartObj = wsDash.ChartObjects.Add( _
            Left:=chartBox.Left + chartBox.Width * 0.05, _
            Top:=chartBox.Top + chartBox.Height * 0.05, _
            Width:=chartBox.Width * 0.9, _
            Height:=chartBox.Height * 0.9)

        ' 차트 유형 설정
        Select Case i
            Case 1: chartObj.Chart.ChartType = xlColumnClustered
            Case 2: chartObj.Chart.ChartType = xlBarStacked100
            Case 3: chartObj.Chart.ChartType = xlDoughnut
            Case 4: chartObj.Chart.ChartType = xlBarClustered
            Case 5: chartObj.Chart.ChartType = xlXYScatterLines
        End Select
    Next i
End Sub' > create_charts.vba

# 매크로 실행
oa excel run-macro --file-path "게임판매대시보드.xlsx" --macro-name "CreateCharts"
```

## SECTION 6: 슬라이서 추가

### 슬라이서 생성 및 연결
```bash
# 슬라이서 박스 추가를 위한 도형 생성
oa excel shape-add --file-path "게임판매대시보드.xlsx" --sheet "대시보드" --shape-type "Rectangle" --name "SlicerBox" --left 820 --top 90 --width 150 --height 520

# VBA 코드로 슬라이서 추가 (add_slicers.vba)
echo 'Sub AddSlicers()
    Dim wsPivot As Worksheet, wsDash As Worksheet
    Set wsPivot = ThisWorkbook.Sheets("피벗")
    Set wsDash = ThisWorkbook.Sheets("대시보드")

    Dim pt As PivotTable
    Set pt = wsPivot.PivotTables(1)

    Dim slicerBox As Shape
    Set slicerBox = wsDash.Shapes("SlicerBox")

    ' 플랫폼 슬라이서 추가
    Dim sc1 As SlicerCache
    Set sc1 = ThisWorkbook.SlicerCaches.Add2(pt, "플랫폼")
    sc1.Slicers.Add wsDash, , "플랫폼", "플랫폼 필터", _
        slicerBox.Left + 10, slicerBox.Top + 10, 130, 200

    ' 장르 슬라이서 추가
    Dim sc2 As SlicerCache
    Set sc2 = ThisWorkbook.SlicerCaches.Add2(pt, "장르")
    sc2.Slicers.Add wsDash, , "장르", "장르 필터", _
        slicerBox.Left + 10, slicerBox.Top + 220, 130, 200

    ' 모든 피벗테이블 연결
    Dim i As Integer
    For i = 2 To wsPivot.PivotTables.Count
        sc1.PivotTables.AddPivotTable wsPivot.PivotTables(i)
        sc2.PivotTables.AddPivotTable wsPivot.PivotTables(i)
    Next i
End Sub' > add_slicers.vba

# 매크로 실행
oa excel run-macro --file-path "게임판매대시보드.xlsx" --macro-name "AddSlicers"
```

## 워크북 최종 저장 및 확인

```bash
# 워크북 저장
oa excel workbook-save --file-path "게임판매대시보드.xlsx"

# 최종 구조 확인
oa excel workbook-info --file-path "게임판매대시보드.xlsx" --include-sheets --include-properties

# 대시보드 시트 활성화
oa excel sheet-activate --file-path "게임판매대시보드.xlsx" --sheet "대시보드"
```

## 주의사항

1. **VBA 매크로 실행**: 현재 `oa` CLI는 VBA 매크로 직접 실행 기능이 없으므로, `run-macro` 명령은 향후 구현이 필요합니다.

2. **대안 접근법**: VBA 대신 Python/xlwings를 사용하여 동일한 작업을 수행할 수 있습니다:
   - 피벗테이블: pandas의 pivot_table 기능 활용
   - 차트: xlwings의 차트 API 활용
   - 도형: xlwings의 shapes API 활용

3. **데이터 파일 준비**: CSV 또는 JSON 형식으로 샘플 데이터를 미리 준비하면 `table-write` 명령으로 쉽게 가져올 수 있습니다.

4. **연속 작업**: 옵션 없이 활성 워크북 자동 선택으로 한 번 열린 워크북에서 연속적인 작업이 가능합니다:
   ```bash
   oa excel workbook-open --file-path "게임판매대시보드.xlsx"
   oa excel sheet-add --name "피벗"
   oa excel sheet-add --name "대시보드"
   oa excel range-write --sheet "Data" --range "A1" --data '[...]'
   ```

5. **임시 파일 활용**: 대량의 데이터는 임시 파일을 통해 전달:
   ```bash
   # 데이터를 임시 파일로 저장
   echo '[["데이터1", "데이터2"], ["값1", "값2"]]' > temp_data.json
   oa excel range-write --file-path "게임판매대시보드.xlsx" --sheet "Data" --range "A1" --data-file "temp_data.json"
   ```