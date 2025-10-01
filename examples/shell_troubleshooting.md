# Excel Shell Mode - 문제 해결 가이드

Excel Shell Mode 사용 시 자주 발생하는 문제와 해결 방법을 정리했습니다.

## 목차
1. [설치 및 실행 문제](#설치-및-실행-문제)
2. [워크북 연결 문제](#워크북-연결-문제)
3. [컨텍스트 관련 문제](#컨텍스트-관련-문제)
4. [명령어 실행 문제](#명령어-실행-문제)
5. [성능 및 안정성](#성능-및-안정성)
6. [Windows 특수 문제](#windows-특수-문제)

---

## 설치 및 실행 문제

### 문제 1: Shell이 시작되지 않음

**증상**:
```bash
$ oa excel shell
Error: xlwings module not found
```

**원인**: xlwings가 설치되지 않았거나 가상환경 문제

**해결 방법**:
```bash
# 1. xlwings 설치 확인
$ pip list | grep xlwings

# 2. 없다면 설치
$ pip install xlwings

# 3. 또는 전체 패키지 재설치
$ pip install -e .
```

---

### 문제 2: prompt-toolkit 에러

**증상**:
```bash
$ oa excel shell
ModuleNotFoundError: No module named 'prompt_toolkit'
```

**원인**: prompt-toolkit 미설치

**해결 방법**:
```bash
# prompt-toolkit과 click-repl 설치
$ pip install prompt-toolkit>=3.0.50 click-repl>=0.3.0

# 또는
$ pip install -e .
```

---

### 문제 3: Excel이 설치되지 않음

**증상**:
```bash
[Excel: None > None] > workbook-list
Error: Excel application not found
```

**원인**: Excel이 시스템에 설치되지 않았거나 COM 등록 문제 (Windows)

**해결 방법**:
1. Excel이 설치되어 있는지 확인
2. Windows: Excel을 관리자 권한으로 한 번 실행하여 COM 등록
3. macOS: Excel이 Applications 폴더에 있는지 확인

---

## 워크북 연결 문제

### 문제 4: 워크북을 찾을 수 없음

**증상**:
```bash
[Excel: None > None] > use workbook "Sales.xlsx"
Error: Workbook 'Sales.xlsx' not found
```

**원인**:
- 워크북이 열려있지 않음
- 파일명이 정확하지 않음
- 경로가 포함되지 않음

**해결 방법**:
```bash
# 1. 현재 열린 워크북 확인
[Excel: None > None] > workbook-list

# 2. 정확한 파일명 사용 (대소문자 구분)
[Excel: None > None] > use workbook "sales.xlsx"  # 소문자
[Excel: None > None] > use workbook "Sales.xlsx"  # 대문자

# 3. 전체 경로 사용
[Excel: None > None] > use workbook "C:/Data/Sales.xlsx"

# 4. 워크북이 닫혀있다면 먼저 열기
$ oa excel workbook-open --file-path "C:/Data/Sales.xlsx"
# 그 후 Shell 시작
$ oa excel shell --workbook-name "Sales.xlsx"
```

---

### 문제 5: 활성 워크북 선택 실패

**증상**:
```bash
$ oa excel shell
Warning: No active workbook found
[Excel: None > None] >
```

**원인**: Excel에 열린 워크북이 없음

**해결 방법**:
```bash
# 방법 1: Shell 내에서 워크북 열기
[Excel: None > None] > workbook-list  # 열린 파일 확인
[Excel: None > None] > use workbook "파일명.xlsx"

# 방법 2: Shell 시작 시 파일 지정
$ oa excel shell --file-path "C:/Data/Report.xlsx"

# 방법 3: Excel에서 파일을 수동으로 열고 다시 시작
```

---

## 컨텍스트 관련 문제

### 문제 6: 컨텍스트가 자동 주입되지 않음

**증상**:
```bash
[Excel: sales.xlsx > Data] > range-read --range A1:C10
Error: workbook_name is required
```

**원인**: 컨텍스트가 제대로 설정되지 않음

**해결 방법**:
```bash
# 1. 현재 컨텍스트 확인
[Excel: sales.xlsx > Data] > show context

# 2. 워크북이 None이면 설정
[Excel: None > None] > use workbook "sales.xlsx"

# 3. 시트가 None이면 설정
[Excel: sales.xlsx > None] > use sheet "Data"

# 4. 컨텍스트 확인 후 재시도
[Excel: sales.xlsx > Data] > show context
[Excel: sales.xlsx > Data] > range-read --range A1:C10
```

---

### 문제 7: 시트 전환 실패

**증상**:
```bash
[Excel: sales.xlsx > Sheet1] > use sheet "Data"
Error: Sheet 'Data' not found
```

**원인**: 시트 이름이 정확하지 않음

**해결 방법**:
```bash
# 1. 현재 워크북의 모든 시트 확인
[Excel: sales.xlsx > Sheet1] > sheets

# 출력 예시:
# Available sheets in sales.xlsx:
#   1. Sheet1 (Active)
#   2. RawData
#   3. Summary

# 2. 정확한 시트명 사용 (공백, 대소문자 주의)
[Excel: sales.xlsx > Sheet1] > use sheet "RawData"

# 3. 시트명에 공백이 있으면 따옴표 사용
[Excel: sales.xlsx > Sheet1] > use sheet "Sales Data"
```

---

## 명령어 실행 문제

### 문제 8: 명령어를 찾을 수 없음

**증상**:
```bash
[Excel: sales.xlsx > Data] > read-range --range A1:C10
Error: Unknown command: read-range
```

**원인**: 명령어 이름 오타

**해결 방법**:
```bash
# 1. Tab 자동완성 활용
[Excel: sales.xlsx > Data] > ra<TAB>
# → range-read, range-write, range-convert

# 2. help로 명령어 확인
[Excel: sales.xlsx > Data] > help

# 3. 올바른 명령어 사용
[Excel: sales.xlsx > Data] > range-read --range A1:C10
```

---

### 문제 9: 옵션 인자 오류

**증상**:
```bash
[Excel: sales.xlsx > Data] > range-read A1:C10
Error: Missing option '--range'
```

**원인**: 옵션 플래그 누락

**해결 방법**:
```bash
# 올바른 형식: --옵션 값
[Excel: sales.xlsx > Data] > range-read --range A1:C10

# 여러 옵션 사용
[Excel: sales.xlsx > Data] > range-write --range A1 --data '[[1,2,3]]'

# 명령어 도움말 확인
[Excel: sales.xlsx > Data] > range-read --help
```

---

### 문제 10: JSON 데이터 파싱 에러

**증상**:
```bash
[Excel: sales.xlsx > Data] > range-write --range A1 --data [[1,2,3]]
Error: Invalid JSON format
```

**원인**: JSON 형식이 올바르지 않음

**해결 방법**:
```bash
# 1. JSON은 작은따옴표로 감싸기
[Excel: sales.xlsx > Data] > range-write --range A1 --data '[[1,2,3]]'

# 2. 복잡한 JSON은 이스케이프 주의
[Excel: sales.xlsx > Data] > range-write --range A1 --data '[["Name","Age"],["Alice",25]]'

# 3. 큰 데이터는 파일 사용
# Python으로 data.json 생성 후
[Excel: sales.xlsx > Data] > range-write --range A1 --data-file data.json
```

---

## 성능 및 안정성

### 문제 11: Shell이 느림

**증상**: 명령 실행이 느리거나 응답 없음

**원인**:
- 대용량 데이터 처리
- COM 객체 누적
- 메모리 부족

**해결 방법**:
```bash
# 1. 대용량 데이터는 페이징 사용
[Excel: sales.xlsx > Data] > range-read --range A1:A100000 --limit 1000

# 2. Shell 재시작으로 리소스 해제
[Excel: sales.xlsx > Data] > exit
$ oa excel shell

# 3. 작은 범위로 나눠서 처리
[Excel: sales.xlsx > Data] > range-read --range A1:A1000
[Excel: sales.xlsx > Data] > range-read --range A1001:A2000

# 4. table-read에서 샘플링 사용
[Excel: sales.xlsx > Data] > table-read --limit 100 --sample-mode
```

---

### 문제 12: Shell 응답 없음 (Frozen)

**증상**: 명령 입력 후 멈춤

**원인**:
- Excel COM 객체 충돌
- 대용량 데이터 로딩
- 순환 참조

**해결 방법**:
```bash
# 1. Ctrl+C로 명령 취소
^C
[Excel: sales.xlsx > Data] >

# 2. Excel 작업 관리자에서 확인
# Excel 프로세스가 응답 중인지 확인

# 3. Shell 종료 후 Excel 재시작
[Excel: sales.xlsx > Data] > exit
# Excel 종료 → 재시작 → Shell 재시작

# 4. 안전한 워크플로우 사용
# - 작은 범위부터 시작
# - 단계별로 확인
# - 대용량 작업은 일반 CLI 사용
```

---

## Windows 특수 문제

### 문제 13: 한글 파일명 에러

**증상**:
```bash
[Excel: None > None] > use workbook "판매데이터.xlsx"
Error: Workbook not found
```

**원인**:
- 한글 파일명 인코딩 문제
- 경로에 한글 포함

**해결 방법**:
```bash
# 1. 절대 경로 사용
[Excel: None > None] > use workbook "C:/데이터/판매데이터.xlsx"

# 2. 파일명을 영문으로 변경 권장
# "판매데이터.xlsx" → "sales_data.xlsx"

# 3. workbook-list로 정확한 이름 확인
[Excel: None > None] > workbook-list
# 출력에서 정확한 파일명 복사하여 사용
```

---

### 문제 14: cp949 인코딩 에러

**증상**:
```
UnicodeDecodeError: 'cp949' codec can't decode byte 0xe2
```

**원인**: Windows 콘솔 인코딩 문제 (자동화 스크립트에서만 발생)

**해결 방법**:
```bash
# 1. 대화형 Shell에서는 문제없음 (직접 사용)
$ oa excel shell

# 2. Python 스크립트에서 subprocess 사용 시
import subprocess

result = subprocess.run(
    ["oa", "excel", "shell"],
    capture_output=True,
    text=True,
    encoding='utf-8',  # 명시적 UTF-8 인코딩
    errors='replace'   # 에러 무시
)

# 3. PowerShell 사용 시 UTF-8 설정
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
```

---

### 문제 15: COM 객체 에러 (Windows)

**증상**:
```
pywintypes.com_error: (-2147352567, 'Exception occurred', ...)
```

**원인**:
- Excel COM 인터페이스 충돌
- 여러 Python 프로세스에서 동시 Excel 접근

**해결 방법**:
```bash
# 1. Excel 프로세스 확인 및 종료
# 작업 관리자에서 EXCEL.EXE 모두 종료

# 2. Shell 재시작
$ oa excel shell

# 3. 동시 접근 방지
# - 여러 터미널에서 동시에 Shell 실행 금지
# - 한 번에 하나의 Shell만 실행

# 4. Excel을 관리자 권한으로 실행 (COM 재등록)
```

---

## 디버깅 팁

### 1. 상세 로그 활성화
```bash
# 환경 변수 설정 (개발 중)
$ export OA_DEBUG=1
$ oa excel shell
```

### 2. 단계별 검증
```bash
# 각 단계마다 show context로 상태 확인
[Excel: sales.xlsx > Data] > show context
[Excel: sales.xlsx > Data] > range-read --range A1:C10
[Excel: sales.xlsx > Data] > show context  # 다시 확인
```

### 3. 최소 재현 테스트
```bash
# 문제가 발생하는 최소 단위로 축소
$ oa excel shell
[Excel: None > None] > workbook-list
[Excel: None > None] > use workbook "test.xlsx"
[Excel: test.xlsx > None] > use sheet "Sheet1"
[Excel: test.xlsx > Sheet1] > range-read --range A1:A1
# 여기서 에러? → 해당 명령만 일반 CLI로 테스트
```

### 4. 일반 CLI로 비교 테스트
```bash
# Shell에서 안 되면 일반 CLI로 시도
$ oa excel range-read --workbook-name "test.xlsx" --sheet "Sheet1" --range A1:A1
# 이것도 안 되면 Shell 문제가 아님
```

---

## 자주 묻는 질문 (FAQ)

### Q1: Shell Mode와 일반 CLI 중 언제 무엇을 사용해야 하나요?
**A**:
- **Shell Mode**: 3개 이상 연속 작업, 탐색적 분석, 대화형 작업
- **일반 CLI**: 단발성 작업 1-2개, 자동화 스크립트, CI/CD 파이프라인

### Q2: Tab 자동완성이 느려요
**A**:
- 터미널 성능 문제일 수 있습니다
- Windows Terminal, VSCode Terminal 사용 권장
- 레거시 cmd.exe는 느릴 수 있음

### Q3: 명령어 히스토리가 저장되나요?
**A**:
- 예, `~/.oa_excel_shell_history` 파일에 저장됩니다
- 다음 세션에서도 위/아래 화살표로 접근 가능

### Q4: 여러 워크북을 동시에 작업할 수 있나요?
**A**:
- Shell 하나에서는 한 번에 하나의 워크북만 활성화
- `use workbook` 명령으로 전환 가능
- 동시 작업이 필요하면 일반 CLI 사용

### Q5: Shell이 예상치 못하게 종료되었어요
**A**:
- 로그 확인: 에러 메시지가 있었는지 스크롤업하여 확인
- 재시작: `$ oa excel shell`
- 컨텍스트 복원: `use workbook`, `use sheet` 재설정

---

## 지원 받기

문제가 계속되면:

1. **GitHub Issues**: https://github.com/pyhub-apps/pyhub-office-automation/issues
2. **로그 첨부**: 에러 메시지와 실행 단계를 자세히 기록
3. **환경 정보**:
   ```bash
   $ oa info
   $ python --version
   $ pip list | grep -E "(xlwings|prompt-toolkit|click-repl)"
   ```

---

**Happy Shell Automation!** 🚀
