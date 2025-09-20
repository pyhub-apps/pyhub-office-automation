# `PRD.md`

## **PRD: `pyhub-office-automation` - Python 기반 엑셀 및 HWP 자동화 에이전트 스크립트**

### **1. 개요 (Introduction)**

*   **문서 목적:** 본 문서는 Python 기반 엑셀(xlwings) 및 HWP(pyhwpx) 자동화 스크립트 패키지인 `pyhub-office-automation`의 개발 요구사항을 정의하고, 이를 Gemini CLI 등의 AI 에이전트에서 활용하여 비전문가 사용자를 위한 대화형 자동화를 구현하기 위한 지침을 제공합니다.
*   **제품 비전:** Python 및 CLI에 대한 지식이 없는 일반 사용자도 AI 에이전트(주로 Gemini CLI)와의 대화를 통해 복잡한 엑셀 및 HWP 작업을 쉽게 자동화할 수 있도록 지원하는 것을 목표로 합니다. `pyhub-office-automation` 패키지는 사용자의 로컬 환경에서 실행되어 로컬 파일에 직접 접근하며, AI 에이전트가 사용자의 의도를 파악하고, 필요한 스크립트를 호출하여 작업을 수행하며, 결과를 사용자 친화적으로 보고하는 시스템을 구축합니다.
*   **대상 사용자:** Python 개발 경험, CLI 사용 경험이 없는 비전문가 사용자. (AI 에이전트가 중간 매개 역할을 수행하여 사용자와 상호작용)
*   **패키지 이름:** `pyhub-office-automation`
*   **메인 CLI 명령어:** `oa`

### **2. 기능 요구사항 (Functional Requirements)**

#### **2.1. 엑셀 자동화 스크립트 (xlwings 기반)**

*   **설계 원칙:**
    *   `pyhub_office_automation/excel/` 디렉토리 내에 작고 재사용 가능한 모듈형 함수(스크립트)들을 구성합니다. 각 스크립트는 명확한 단일 책임(Single Responsibility)을 가지며 `click` 기반의 CLI 인터페이스를 가집니다.
    *   `oa excel <command>` 형태로 호출됩니다. (예: `oa excel open-workbook`)
    *   xlwings가 제공하는 모든 기능을 활용할 수 있도록 다양한 스크립트를 제공하되, 초기에는 가장 자주 사용될 법한 핵심 기능 위주로 구성합니다.
    *   각 스크립트는 자체적인 버전을 가지며, CLI `--version` 옵션을 통해 조회 가능합니다.
*   **주요 기능 (예시):**
    *   `workbook-open.py`: 특정 경로의 엑셀 파일 열기. (호출 예시: `oa excel workbook-open --file-path "C:\data\sample.xlsx"`)
    *   `workbook-save.py`: 현재 열려있는 엑셀 파일 저장 (다른 이름으로 저장 포함).
    *   `workbook-close.py`: 엑셀 파일 닫기.
    *   `workbook-create.py`: 새 엑셀 파일 생성.
    *   `add-sheet.py`: 새 워크시트 추가.
    *   `delete-sheet.py`: 특정 워크시트 삭제.
    *   `rename-sheet.py`: 워크시트 이름 변경.
    *   `activate-sheet.py`: 특정 워크시트 활성화.
    *   `read-range.py`: 특정 범위의 셀 값 읽기.
    *   `write-range.py`: 특정 범위에 값 쓰기.
    *   `read-table.py`: 특정 범위의 데이터를 테이블(Pandas DataFrame) 형태로 읽기.
    *   `write-table.py`: Pandas DataFrame을 특정 범위에 쓰기.
    *   `find-value.py`: 특정 값을 찾아 해당 셀의 주소 반환.
    *   `set-cell-format.py`: 특정 셀/범위의 서식 설정 (예: 글꼴, 크기, 색상, 정렬).
    *   `set-border.py`: 특정 셀/범위의 테두리 설정.
    *   `auto-fit-columns.py`: 특정 범위의 열 너비 자동 맞춤.
    *   `run-macro.py`: 엑셀 매크로 실행.
    *   `add-chart.py`: 차트 추가 (차트 종류, 데이터 범위 지정).

#### **2.2. HWP 자동화 스크립트 (pyhwpx 기반, Windows COM 활용)**

*   **설계 원칙:**
    *   `pyhub_office_automation/hwp/` 디렉토리 내에 작고 재사용 가능한 모듈형 함수(스크립트)들을 구성합니다. 각 스크립트는 명확한 단일 책임(Single Responsibility)을 가지며 `click` 기반의 CLI 인터페이스를 가집니다.
    *   `oa hwp <command>` 형태로 호출됩니다. (예: `oa hwp open-hwp`)
    *   pyhwpx가 제공하는 모든 기능을 활용할 수 있도록 다양한 스크립트를 제공하되, 초기에는 핵심 기능 위주로 구성합니다.
    *   각 스크립트는 자체적인 버전을 가지며, CLI `--version` 옵션을 통해 조회 가능합니다.
*   **주요 기능 (예시):**
    *   `open-hwp.py`: HWP 파일 열기. (호출 예시: `oa hwp open-hwp --file-path "C:\data\document.hwp"`)
    *   `save-hwp.py`: HWP 파일 저장 (다른 이름으로 저장 포함).
    *   `close-hwp.py`: HWP 파일 닫기.
    *   `create-hwp.py`: 새 HWP 문서 생성.
    *   `insert-text.py`: 현재 커서 위치 또는 특정 위치에 텍스트 삽입.
    *   `replace-text.py`: 문서 내 특정 문자열을 다른 문자열로 대체.
    *   `get-text.py`: 문서 전체 또는 특정 범위의 텍스트 추출.
    *   `set-text-format.py`: 특정 텍스트의 서식 설정 (글꼴, 크기, 색상 등).
    *   `insert-table.py`: 표 삽입 (행/열 수, 셀 너비 등).
    *   `fill-table-data.py`: 특정 표에 데이터 채우기.
    *   `get-table-data.py`: 특정 표의 데이터 추출.
    *   `insert-image.py`: 이미지 삽입 (경로, 크기, 위치 지정).
    *   `insert-page-break.py`: 페이지 나누기 삽입.
    *   `get-page-count.py`: 총 페이지 수 반환.
    *   `merge-documents.py`: 여러 HWP 파일을 하나의 문서로 병합.
    *   `find-replace-dialog.py`: 한글 찾기/바꾸기 대화상자 호출 (필요시).

#### **2.3. AI 에이전트 (주로 Gemini CLI) 연동 기능**

*   **목표:** AI 에이전트가 `oa` 메인 명령어와 각 서브 명령어의 `--help` 및 `--version` 출력 내용을 파싱하여 스크립트의 기능, 입력 파라미터, 사용 예시 및 버전 정보를 학습하고, 이를 기반으로 사용자 요청을 처리하도록 합니다.
*   **입력 파라미터 처리:**
    *   모든 스크립트의 입력은 CLI 옵션(`--option-name value`) 방식으로 전달합니다.
    *   긴 텍스트(예: HWP에 삽입할 내용, 엑셀에 쓸 다량의 데이터)나 파일 경로가 필요한 경우, AI 에이전트가 `tempfile` 모듈을 사용하여 임시 파일을 생성하고 내용을 저장한 후, 해당 임시 파일의 경로를 스크립트 인자로 넘겨줍니다. 스크립트는 작업을 마친 후 해당 임시 파일을 자동으로 제거합니다.
*   **출력값 처리:**
    *   스크립트는 `json` 또는 `yaml` 등 AI 에이전트가 파싱하기 쉬운 구조화된 형태로 결과를 표준 출력(stdout)으로 반환하며, 결과 JSON에는 스크립트의 `version` 정보가 포함되어야 합니다.
    *   **AI의 가공:** AI 에이전트는 스크립트의 Raw 출력값을 받아서 사용자에게 가독성 높고 자연어에 가까운 형태로 가공하여 보고합니다. 오류 발생 시에도 AI가 이를 해석하여 사용자에게 이해하기 쉬운 메시지를 전달합니다.

#### **2.4. CLI 활용 지침 스크립트 (Self-Documentation CLI)**

*   **메인 명령어:** `oa`
*   **목표:** AI 에이전트 또는 사용자가 `pyhub-office-automation` 패키지 자체의 활용 방법, 설치 방법, 주요 스크립트 목록 등을 `oa` 메인 CLI 명령을 통해 조회할 수 있도록 합니다.
*   **설계:** `oa` 메인 CLI 스크립트는 다음과 같은 서브 명령어를 제공합니다.
    *   **`oa info`:** 패키지 전체의 버전 정보, 설치 상태, 주요 의존성 목록을 출력합니다.
    *   **`oa excel list`:** 엑셀 자동화 서브 명령어 목록과 각 명령어의 간략한 설명, 버전 정보를 출력합니다.
    *   **`oa hwp list`:** HWP 자동화 서브 명령어 목록과 각 명령어의 간략한 설명, 버전 정보를 출력합니다.
    *   **`oa get-help <category> <command>`:** 특정 카테고리(excel, hwp)의 명령어에 대한 `--help` 출력 내용을 반환합니다. (예: `oa get-help excel open-workbook`)
    *   **`oa install-guide`:** Python 설치부터 `pip install pyhub-office-automation`까지의 단계별 설치 가이드를 출력합니다. (AI가 이 정보를 활용하여 사용자에게 설치를 안내)
*   **버전 관리:** 이 메인 `oa` CLI 스크립트 자체도 버전 정보를 가집니다.

### **3. 비기능 요구사항 (Non-Functional Requirements)**

*   **성능:**
    *   작은 수의 파일에 대한 작업이며, 대화형 자동화를 목표로 하므로, 각 스크립트의 실행 시간은 수 초 이내를 지향합니다. (대용량 파일 처리는 고려하지 않음)
*   **보안:**
    *   **AI 학습 방지:** 스크립트를 통해 처리되는 문서의 내용은 절대 AI 에이전트의 학습 데이터로 사용되지 않도록 명확한 지침을 AI 에이전트에게 제공해야 합니다.
    *   임시 파일 사용 시, 반드시 사용 후 즉시 제거하여 데이터 잔존을 방지합니다.
*   **사용성:**
    *   스크립트 자체는 CLI로 동작하지만, 최종 사용자는 AI 에이전트와의 자연어 대화를 통해 모든 작업을 수행하므로, 스크립트의 복잡성을 사용자에게 직접 노출시키지 않습니다.
*   **호환성:**
    *   **운영체제:** Windows 10/11 전용. 다른 OS는 고려하지 않습니다.
    *   **Python 버전:** Python 3.13 이상.
    *   **라이브러리:** `xlwings`, `pyhwpx`, `click`, `pandas`, `pathlib`, `tempfile`.

### **4. 기술 스택 (Technical Stack)**

*   **언어:** Python 3.13+
*   **CLI 프레임워크:** `click`
*   **엑셀 라이브러리:** `xlwings` (Windows COM 기반, macOS AppleScript 기반)
*   **HWP 라이브러리:** `pyhwpx` (Windows COM 기반)
*   **데이터 처리:** `pandas`
*   **파일 시스템 유틸리티:** `pathlib`, `tempfile`

### **5. 테스트 및 검증 전략 (Testing & Validation Strategy)**

*   **단위 테스트 (Unit Tests):**
    *   각 스크립트 내의 핵심 함수별로 `pytest`를 활용하여 단위 테스트를 작성합니다.
    *   정상 동작 케이스뿐만 아니라, `Edge Case` (예: 빈 파일, 잘못된 경로, 잘못된 입력 타입, HWP 프로그램 미설치, 엑셀 시트명 오류 등) 에 대한 테스트를 반드시 포함합니다.
*   **CLI 구동 테스트:**
    *   각 스크립트의 CLI 명령어를 직접 구동하여 `--help` 옵션의 정상 작동 및 출력 내용을 검증합니다.
    *   실제 인자를 전달하여 스크립트가 예상대로 동작하는지 검증하고, 오류 메시지가 적절하게 출력되는지 확인합니다.
*   **AI 에이전트 연동 테스트:**
    *   Gemini CLI를 통해 실제 사용자 요청 시나리오를 가정한 테스트를 수행하여, AI가 스크립트를 올바르게 선택하고 인자를 전달하며, 결과를 적절히 해석하여 사용자에게 보고하는지 검증합니다.
*   **테스트 보고:**
    *   각 스크립트 개발 완료 시, 단위 테스트 및 CLI 구동 테스트 결과를 상세히 보고합니다. (테스트 성공/실패 여부, 오류 메시지, Edge 케이스 처리 결과 등)

### **6. 배포 및 활용 가이드라인 (Deployment & Usage Guidelines)**

*   **배포 방식:** **PyPI (Python Package Index)를 통한 `pyhub-office-automation` 패키지 배포**
    *   패키지는 `pyhub-office-automation` 이라는 이름으로 PyPI에 등록됩니다.
    *   사용자는 `pip install pyhub-office-automation` 명령어를 통해 패키지를 설치합니다.
    *   `oa` 메인 CLI 명령어는 `setup.py`의 `entry_points` 설정을 통해 시스템 PATH에 자동으로 등록됩니다.
*   **로컬 환경 요구사항:**
    *   Windows 10/11 운영체제
    *   Python 3.13 이상 설치 (환경 변수 PATH 설정 포함)
    *   한글(HWP) 프로그램 설치 (HWP 자동화 스크립트 사용 시 필수)
*   **AI 에이전트 활용 프로세스 (주로 Gemini CLI):**
    1.  **초기 설정/안내:**
        *   사용자가 `oa` 관련 기능을 요청할 경우, AI는 먼저 사용자 PC에 Python 및 `pyhub-office-automation` 패키지가 설치되어 있는지 확인합니다.
        *   설치되어 있지 않다면, AI는 `oa install-guide` 명령어를 활용하여 사용자에게 Python 설치 링크, 환경 변수 설정 방법, 그리고 `pip install pyhub-office-automation` 명령어를 단계별로 안내합니다.
        *   설치 완료 후, AI는 사용자에게 설치 확인을 요청합니다.
    2.  **스크립트 정보 학습:**
        *   AI 에이전트는 `oa excel list`, `oa hwp list` 명령어를 주기적으로 호출하여 사용 가능한 서브 명령어 목록과 각 명령어의 버전 정보를 업데이트합니다.
        *   각 명령어의 상세한 사용법은 `oa get-help <category> <command>` 명령어를 통해 조회하여 학습합니다.
    3.  **사용자 요청 처리:**
        *   AI는 사용자 요청을 분석하여 어떤 `oa` 서브 명령어가 필요한지 식별합니다. (예: `oa excel open-workbook`)
        *   해당 명령어의 `--help` 정보를 바탕으로 필요한 입력 파라미터를 사용자에게 질문하여 수집합니다.
        *   필요한 파라미터가 모두 수집되면, AI는 `oa excel open-workbook --file-path "..."`와 같은 형태로 CLI 명령어를 구성합니다.
        *   AI는 이 구성된 명령어를 사용자의 로컬 셸 환경에서 실행하는 기능(`execute_shell_command`와 같은 내부 도구 함수)을 통해 실행합니다.
        *   명령어의 JSON 결과값을 받아 사용자 친화적인 자연어로 해석하여 보고합니다.
    4.  **업데이트 안내:** AI는 `oa info` 명령어를 통해 주기적으로 패키지 업데이트 여부를 확인하고, 새로운 버전이 있다면 사용자에게 `pip install --upgrade pyhub-office-automation`을 통해 업데이트를 권장하고 안내합니다.

### **7. 향후 계획 (Future Considerations)**

*   **실행 파일 (Executable) 배포:** 비전문가 사용자를 위한 Python 설치 없는 실행 환경 제공 방안으로 `.exe` 파일 배포를 추후 검토합니다.
*   **다른 AI 에이전트 지원:** Codex CLI, Claude Code 등 다른 AI 환경으로의 확장 고려.

