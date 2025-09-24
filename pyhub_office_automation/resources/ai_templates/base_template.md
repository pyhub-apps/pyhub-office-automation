# Code Assistant Context

## oa : pyhub-office-automation CLI utility

+ `oa` 명령을 통해, 현재 구동 중인 엑셀 프로그램과 통신하며 시트 데이터 읽고 쓰기, 피벗 테이블 생성, 차트 생성 등을 할 수 있어.
    - 엑셀 파일 접근에는 `oa` 프로그램을 사용하고, 한 번에 10개 이상의 많은 엑셀 파일을 읽어야할 때에는 효율성을 위해 python과 python 엑셀 라이브러리를 통해 읽어줘.
    - 엑셀 파일을 열기 전에, 반드시 `oa excel workbook-list` 명령으로 열려진 엑셀파일이 있는 지 확인해줘.
    - 파일을 읽을 수 없다면 유저에게 파일 경로를 꼼꼼하게 확인해보라고 알려줘.
+ **ALWAYS** `oa excel --help` 명령으로 지원 명령을 먼저 확인하고, `oa excel 명령 --help` 명령으로 사용법을 확인한 뒤에 명령을 사용해줘.
+ `oa llm-guide` 명령으로 지침을 조회해줘.
+ `--workbook-name` 인자나 `--file-path` 인자를 지정하지 않으면 활성화된 워크북을 참조하고, `--sheet` 인자를 지정하지 않으면, 활성화된 시트를 참조함.
    - 모든 `oa` 명령에서 명시적으로 `--sheet` 인자로 시트명을 지정하여 읽어오자.

## 핵심 사용 패턴

### 1. 작업 전 상황 파악
```bash
# 현재 열린 워크북 확인
oa excel workbook-list

# 활성 워크북 정보 확인
oa excel workbook-info --include-sheets
```

### 2. 워크북 연결 방법
- **자동 연결**: 옵션 없이 사용하면 활성 워크북 자동 사용 (기본값)
- **파일 경로**: `--file-path "경로/파일명.xlsx"`
- **워크북 이름**: `--workbook-name "파일명.xlsx"`

### 3. 데이터 읽기/쓰기
```bash
# 범위 데이터 읽기
oa excel range-read --sheet "Sheet1" --range "A1:C10"

# 데이터 쓰기
oa excel range-write --sheet "Sheet1" --range "A1" --data '[["Name", "Score"], ["Alice", 95]]'

# 테이블 읽기 (pandas DataFrame으로)
oa excel table-read --sheet "Sheet1" --output-file "data.csv"
```

### 4. 차트 생성
```bash
# 기본 차트 생성
oa excel chart-add --sheet "Sheet1" --data-range "A1:B10" --chart-type "Column" --title "Sales Chart"

# 피벗 차트 생성 (Windows만)
oa excel chart-pivot-create --sheet "Sheet1" --data-range "A1:D100" --rows "Category" --values "Sales"
```

## 에러 방지 워크플로우

1. **항상 workbook-list로 시작**: 현재 상황 파악
2. **명시적 시트 지정**: `--sheet` 옵션 사용
3. **단계별 진행**: 복잡한 작업을 작은 단위로 분할
4. **경로 확인**: 파일 경로는 절대 경로나 정확한 상대 경로 사용