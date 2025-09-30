# Excel Shell Mode Testing Guide

## Phase 2 완료 - Excel 명령어 통합

Excel Shell Mode의 Phase 2가 완료되었습니다. 이제 셸 내에서 Excel 명령어를 실행할 수 있으며, 컨텍스트(워크북/시트)가 자동으로 주입됩니다.

### 구현된 기능

#### 1. 컨텍스트 자동 주입
- 셸에서 설정한 워크북과 시트가 자동으로 명령어에 전달됩니다
- `--workbook-name`, `--sheet` 옵션을 생략 가능

#### 2. 지원 명령어
- `range-read` - 셀 범위 읽기
- `range-write` - 셀 범위 쓰기
- `table-list` - 테이블 목록
- `table-read` - 테이블 데이터 읽기
- `chart-add` - 차트 추가
- `chart-list` - 차트 목록
- `workbook-info` - 워크북 정보

### 테스트 방법

#### 사전 준비
1. Excel 파일 준비 (예: `test.xlsx`)
2. 파일에 데이터와 테이블 추가

#### 테스트 시나리오

```bash
# 1. Excel 파일로 셸 시작
uv run oa excel shell --file-path "C:\path\to\test.xlsx"

# 또는 이미 열린 파일로 시작
uv run oa excel shell --workbook-name "test.xlsx"

# 또는 활성 워크북 사용
uv run oa excel shell
```

#### 셸 내부 테스트

```bash
# 1. 컨텍스트 확인
[Excel: test.xlsx > Sheet1] > show context

# 2. 시트 목록 확인
[Excel: test.xlsx > Sheet1] > sheets

# 3. 시트 전환
[Excel: test.xlsx > Sheet1] > use sheet TestData

# 4. 범위 읽기 (컨텍스트 자동 주입)
[Excel: test.xlsx > TestData] > range-read --range A1:C10
# 실제 실행: range-read --range A1:C10 --workbook-name "test.xlsx" --sheet "TestData"

# 5. 테이블 목록
[Excel: test.xlsx > TestData] > table-list
# 실제 실행: table-list --workbook-name "test.xlsx"

# 6. 워크북 정보
[Excel: test.xlsx > TestData] > workbook-info
# 실제 실행: workbook-info --workbook-name "test.xlsx"

# 7. 도움말
[Excel: test.xlsx > TestData] > help

# 8. 종료
[Excel: test.xlsx > TestData] > exit
```

### 고급 사용법

#### 명시적 인자 우선
셸에서도 명시적으로 옵션을 지정하면 컨텍스트 대신 해당 값이 사용됩니다:

```bash
# 현재 시트가 "TestData"이지만, "Sheet1"에서 읽기
[Excel: test.xlsx > TestData] > range-read --range A1:C10 --sheet Sheet1
```

#### 인용부호 처리
복잡한 인자는 인용부호로 감싸세요:

```bash
# JSON 데이터 쓰기
[Excel: test.xlsx > TestData] > range-write --range A1 --data '[[1, 2, 3], [4, 5, 6]]'

# 공백이 포함된 시트명
[Excel: test.xlsx > Sheet1] > use sheet "Sales Report"
```

#### 워크북 전환
여러 워크북을 열어두고 전환할 수 있습니다:

```bash
# 워크북 목록 확인
[Excel: test.xlsx > Sheet1] > workbooks

# 다른 워크북으로 전환
[Excel: test.xlsx > Sheet1] > use workbook report.xlsx
[Excel: report.xlsx > Summary] >
```

### 자동완성 기능

#### Tab 키로 자동완성
```bash
# 명령어 자동완성
[Excel: test.xlsx > Sheet1] > ran<Tab>
# → range-read, range-write

# 시트명 자동완성
[Excel: test.xlsx > Sheet1] > use sheet <Tab>
# → Sheet1, TestData, Summary 등 현재 워크북의 시트 목록

# 워크북명 자동완성 (Windows 전용)
[Excel: test.xlsx > Sheet1] > use workbook <Tab>
# → test.xlsx, report.xlsx 등 열린 워크북 목록
```

### 히스토리 기능

- **위/아래 화살표**: 이전 명령어 탐색
- **히스토리 파일**: `~/.oa_excel_shell_history`

### 알려진 제한사항

#### 1. Windows 전용 기능
- 워크북 목록 조회 (`workbooks`, `use workbook <Tab>`)
- macOS에서는 활성 워크북만 사용 가능

#### 2. 출력 형식
- Excel 명령어의 출력은 JSON 형식으로 표시됩니다
- `--format` 옵션으로 변경 가능

#### 3. 인터랙티브 프롬프트
- Excel 명령어에서 사용자 입력을 요구하는 경우 작동하지 않을 수 있습니다

### 다음 단계 (Phase 3)

#### 계획된 개선사항
1. **더 많은 Excel 명령어 지원**
   - pivot-create, pivot-list
   - sheet-add, sheet-delete
   - 모든 Excel 서브명령어 통합

2. **향상된 출력 형식**
   - Rich 테이블로 데이터 표시
   - 색상 및 포맷팅 개선

3. **에러 처리 개선**
   - 더 명확한 에러 메시지
   - 명령어 실패 시 컨텍스트 유지

4. **배치 명령어 실행**
   - 스크립트 파일에서 명령어 읽기
   - 파이프라인 지원

### 버그 리포트 및 피드백

발견된 버그나 개선 사항은 GitHub Issue에 보고해주세요:
- Issue #85: https://github.com/pyhub-apps/pyhub-office-automation/issues/85

### 기술 세부사항

#### 컨텍스트 주입 메커니즘
```python
# 사용자 입력: range-read --range A1:C10
# 셸이 자동 변환:
# → range-read --range A1:C10 --workbook-name "test.xlsx" --sheet "TestData"

# 명시적 인자가 있으면 우선 적용:
# 사용자 입력: range-read --range A1:C10 --sheet Sheet1
# → 컨텍스트의 sheet 무시, Sheet1 사용
```

#### 명령어 실행 흐름
1. 사용자 입력 파싱 (`shlex.split`)
2. 명령어 타입 판별 (셸 명령 vs Excel 명령)
3. Excel 명령어인 경우:
   - 컨텍스트 확인
   - 인자에 워크북/시트 옵션이 없으면 자동 주입
   - `CliRunner`로 명령어 실행
4. 결과 출력

---

**작성일**: 2025-10-01
**버전**: Phase 2 완료
**상태**: 테스트 준비 완료
