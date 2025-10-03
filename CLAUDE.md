# CLAUDE.md - AI Agent Quick Reference

> **Note**: 이 문서는 AI 에이전트가 빠르게 참조할 수 있는 핵심 가이드입니다. 상세 내용은 [docs/](./docs/) 폴더를 참고하세요.

## 목차
- [프로젝트 개요](#프로젝트-개요)
- [핵심 명령어](#핵심-명령어)
- [AI Agent 워크플로우](#ai-agent-워크플로우)
- [Quick Tips](#quick-tips)
- [상세 문서](#상세-문서)

---

## 프로젝트 개요

### 기본 정보
- **패키지**: `pyhub-office-automation`
- **CLI 명령**: `oa` (office automation)
- **플랫폼**: Windows 10/11 (Python 3.13+)
- **용도**: AI 에이전트 기반 Excel/HWP 자동화

### 아키텍처
```
pyhub_office_automation/
├── excel/          # xlwings Excel 자동화
├── hwp/            # pyhwpx HWP 자동화
├── shell/          # 대화형 Shell 모드
└── cli/            # CLI 진입점
```

### 핵심 의존성
- **xlwings**: Excel 자동화 (Windows COM, macOS AppleScript)
- **pyhwpx**: HWP 자동화 (Windows COM)
- **typer**: CLI 프레임워크
- **pandas**: 데이터 처리
- **prompt-toolkit**: Shell 모드 자동완성

---

## 핵심 명령어

### Excel 명령어 (22개)

**워크북 관리 (4개)**
```bash
oa excel workbook-list           # 열린 워크북 목록 (필수 시작 명령)
oa excel workbook-info           # 활성 워크북 상세 정보
oa excel workbook-open --file-path "file.xlsx"
oa excel workbook-create --save-path "new.xlsx"
```

**시트 관리 (4개)**
```bash
oa excel sheet-activate --sheet "Sheet1"
oa excel sheet-add --name "NewSheet"
oa excel sheet-delete --sheet "OldSheet"
oa excel sheet-rename --old-name "Sheet1" --new-name "Data"
```

**데이터 읽기/쓰기 (2개)**
```bash
oa excel range-read --sheet "Sheet1" --range "A1:C10"
oa excel range-write --sheet "Sheet1" --range "A1" --data '[["Name", "Score"]]'
```

**테이블 (5개)**
```bash
oa excel table-list                  # ⭐ 테이블 구조 + 샘플 데이터 (즉시 분석 가능)
oa excel table-read --output-file "data.csv"
oa excel table-write --data-file "data.csv" --table-name "Sales"
oa excel table-analyze --table-name "Sales"
oa excel metadata-generate
```

**차트 (7개)**
```bash
oa excel chart-add --data-range "A1:B10" --chart-type "Column"
oa excel chart-list
oa excel chart-configure --name "Chart1" --title "New Title"
oa excel chart-position --name "Chart1" --left 100 --top 50
oa excel chart-export --chart-name "Chart1" --output-path "chart.png"
oa excel chart-delete --name "Chart1"
oa excel chart-pivot-create --data-range "A1:D100" --rows "Category"  # Windows only
```

### Shell Mode (연속 작업 3개 이상 시 권장)

```bash
# Excel Shell
oa excel shell                       # 활성 워크북 자동 선택
oa excel shell --file-path "data.xlsx"

# PowerPoint Shell
oa ppt shell --file-path "report.pptx"
```

**Shell 내부 명령어**:
- `help`, `show context`, `use workbook/sheet`, `sheets`, `workbook-info`, `clear`, `exit`

> 📖 **상세 가이드**: [docs/SHELL_USER_GUIDE.md](./docs/SHELL_USER_GUIDE.md)

---

## AI Agent 워크플로우

### 표준 3단계 워크플로우

#### 1️⃣ Context Discovery (상황 파악)
```bash
# 항상 workbook-list로 시작
oa excel workbook-list

# 활성 워크북 구조 확인
oa excel workbook-info

# 테이블 구조 + 샘플 데이터 확인 (즉시 분석 가능)
oa excel table-list
```

#### 2️⃣ Action (작업 수행)
```bash
# 연속 작업 3개 이상 → Shell Mode 사용
oa excel shell

# 단발성 작업 1-2개 → 일반 CLI
oa excel range-read --sheet "Data" --range "A1:C10"
```

#### 3️⃣ Validation (검증)
```bash
# 변경사항 확인
oa excel workbook-info

# 데이터 검증
oa excel range-read --range "A1:A1"  # 헤더 확인
```

### 워크북 연결 방법

```bash
# 옵션 1: 활성 워크북 자동 사용 (기본값, 옵션 없음)
oa excel range-read --range "A1:C10"

# 옵션 2: 파일 경로로 연결
oa excel range-read --file-path "data.xlsx" --range "A1:C10"

# 옵션 3: 열린 워크북 이름으로 연결
oa excel range-read --workbook-name "Sales.xlsx" --range "A1:C10"
```

### Context-Aware 분석 패턴

```bash
# 1. 환경 파악
oa excel workbook-list

# 2. 테이블 구조 즉시 분석 (샘플 데이터 포함)
oa excel table-list
# → Claude가 즉시 차트 제안 가능:
#    "글로벌 판매량 Top 10 막대 차트를 만들어드릴까요?"
#    "지역별 판매량 비교 (북미 vs 유럽)는 어떨까요?"

# 3. 타겟 분석 실행
oa excel chart-add --data-range "GameData[글로벌 판매량]" --chart-type "Column"
```

> 📖 **상세 패턴**: [docs/CLAUDE_CODE_PATTERNS.md](./docs/CLAUDE_CODE_PATTERNS.md)

---

## Quick Tips

### Shell Mode 사용 시점
✅ **사용 권장**:
- 동일 워크북에서 3개 이상 연속 작업
- 탐색적 데이터 분석 (EDA)
- 시트 전환이 빈번한 작업
- Tab 자동완성으로 명령어 입력 속도 10배 향상

❌ **일반 CLI 권장**:
- 단발성 작업 1-2개
- 스크립트/자동화 환경

### 자주하는 실수와 해결법

**❌ 실수 1**: `workbook-list` 없이 바로 작업 시작
```bash
# 나쁜 예
oa excel range-read --range "A1:C10"  # 어느 워크북? 어느 시트?

# 좋은 예
oa excel workbook-list              # 1. 현황 파악
oa excel workbook-info              # 2. 구조 확인
oa excel range-read --sheet "Data" --range "A1:C10"  # 3. 명시적 작업
```

**❌ 실수 2**: `--sheet` 옵션 생략
```bash
# 위험: 활성 시트가 어디인지 모름
oa excel range-read --range "A1:C10"

# 안전: 항상 시트명 명시
oa excel range-read --sheet "RawData" --range "A1:C10"
```

**❌ 실수 3**: Shell Mode를 쓸 곳에 일반 CLI 사용
```bash
# 비효율: 명령어 길이 3배 증가
oa excel range-read --file-path "sales.xlsx" --sheet "Data" --range "A1:C10"
oa excel chart-add --file-path "sales.xlsx" --sheet "Data" --data-range "A1:C10"
oa excel chart-configure --file-path "sales.xlsx" --sheet "Data" --name "Chart1"

# 효율: Shell Mode 사용
oa excel shell --file-path "sales.xlsx"
[Excel: sales.xlsx > None] > use sheet Data
[Excel: sales.xlsx > Data] > range-read --range A1:C10
[Excel: sales.xlsx > Data] > chart-add --data-range A1:C10
[Excel: sales.xlsx > Data] > chart-configure --name Chart1
```

### 차트 선택 가이드

**`chart-add` 권장** (⭐ 기본 선택):
- 간단한 데이터 시각화
- 크로스 플랫폼 호환
- 빠른 생성
- 피벗차트 타임아웃 회피

**`chart-pivot-create` (신중히 사용)**:
- Windows 전용
- `--skip-pivot-link` 옵션 필수
- 대용량 데이터(>1000행) 시 타임아웃 주의

> 📖 **차트 상세 가이드**: [docs/ADVANCED_FEATURES.md](./docs/ADVANCED_FEATURES.md)

### macOS 한글 경로 처리

macOS에서 한글 파일명 자동 NFC 정규화:
```bash
# macOS에서도 한글 파일명 그대로 사용 가능
oa excel range-read --file-path "한글데이터.xlsx" --range "A1:C10"
# → 자동으로 NFD → NFC 변환 처리
```

---

## Python 직접 실행

### Python 환경
```bash
# Python 경로
C:\Users\pyhub\AppData\Local\Programs\Python\Python313\python.EXE

# 패키지 설치
C:\Users\pyhub\AppData\Local\Programs\Python\Python313\python.EXE -m pip install 패키지명
```

### matplotlib 한글 폰트 설정
```python
import matplotlib.pyplot as plt

# 한글 폰트 설정
plt.rcParams['font.family'] = 'Malgun Gothic'
plt.rcParams['axes.unicode_minus'] = False

# 고해상도 설정
plt.rcParams['figure.dpi'] = 300
plt.rcParams['savefig.dpi'] = 300
```

### 대용량 데이터 처리
```python
# 10개 이상 엑셀 파일 → pandas로 직접 처리 (효율적)
import pandas as pd
from pathlib import Path

all_data = []
for file_path in Path().glob("data/*.xlsx"):
    df = pd.read_excel(file_path)
    df['source_file'] = file_path.name
    all_data.append(df)

combined_data = pd.concat(all_data, ignore_index=True)
```

---

## 상세 문서

### 사용자 가이드
- **[Shell Mode 완벽 가이드](./docs/SHELL_USER_GUIDE.md)**
  - Excel Shell / PowerPoint Shell
  - 워크플로우 예제
  - 권장 패턴 5가지

### 고급 기능
- **[Map Chart & 차트 가이드](./docs/ADVANCED_FEATURES.md)**
  - Map Chart 5단계 워크플로우
  - 차트 유형별 예시
  - 피벗테이블 패턴

### Claude Code 특화
- **[분석 패턴 가이드](./docs/CLAUDE_CODE_PATTERNS.md)**
  - 체계적 디버깅 접근
  - 코드 리뷰 체크리스트
  - table-list 즉시 분석 패턴

---

## 개발 정보

### 빌드 스크립트
```powershell
# Windows
.\build_windows.ps1 -BuildType onefile -GenerateMetadata

# macOS/Linux
./build_macos.sh --onefile --metadata
```

### 코드 품질
```powershell
.\lint.ps1          # 전체 검사
.\lint.ps1 -Fix     # 자동 수정
.\lint.ps1 -Quick   # 빠른 검사
```

### 버전 관리 (HeadVer)
```bash
# 표준 버전 태그 생성 (v{major}.{yearweek}.{build})
python scripts/create_version_tag.py --auto-increment

# 특정 빌드 번호
python scripts/create_version_tag.py 19 --message "Fix critical bug"

# 미리보기
python scripts/create_version_tag.py --dry-run --auto-increment
```

---

## 보안 및 데이터 처리

### Privacy Protection
- ⚠️ **중요**: 문서 콘텐츠는 AI 학습에 절대 사용 금지
- 임시 파일 즉시 삭제
- 로컬 전용 처리 (외부 전송 없음)

### File Safety
- 파일 경로 검증 (디렉토리 traversal 방지)
- 프로그램 미설치 시 graceful handling
- 파일 접근 에러 처리

---

## 설정 파일 정보

- **생성 시간**: 2025-09-24 00:05:37
- **패키지 버전**: 9.2539.33
- **생성 명령**: `oa ai-setup claude`

---

**© 2024 pyhub-office-automation** | [GitHub](https://github.com/pyhub-kr/pyhub-office-automation)
