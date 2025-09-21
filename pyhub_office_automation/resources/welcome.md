# 🎉 pyhub-office-automation에 오신 것을 환영합니다!

이 도구는 **AI 에이전트를 위한 Office 자동화 도구**입니다.
Excel과 HWP 문서를 프로그래밍 방식으로 제어할 수 있습니다.

## 🚀 빠른 시작

### 1. 설치 상태 확인
```bash
oa info
```
패키지 버전과 의존성 상태를 확인합니다.

### 2. Excel 명령어 살펴보기
```bash
oa excel list
```
사용 가능한 모든 Excel 자동화 명령어를 확인합니다.

### 3. 현재 작업 환경 확인
```bash
oa excel workbook-list
```
현재 열려있는 Excel 워크북들을 확인합니다.

### 4. 첫 번째 자동화 시도
```bash
# 새 워크북 생성
oa excel workbook-create --save-path "test.xlsx" --name "첫번째테스트"

# 데이터 쓰기
oa excel range-write --use-active --range "A1" --data '["안녕", "세상아"]'

# 데이터 읽기
oa excel range-read --use-active --range "A1:B1"
```

## 📚 더 많은 도움말

- `oa --help` - 전체 명령어 목록
- `oa excel <command> --help` - 특정 Excel 명령어 도움말
- `oa hwp list` - HWP 명령어 목록 (Windows 전용)
- `oa install-guide` - 상세 설치 가이드
- `oa llm-guide` - AI 에이전트를 위한 사용 지침

## 💡 팁

- 모든 명령어는 JSON 형식으로 결과를 출력합니다
- AI 에이전트가 파싱하기 쉽도록 설계되었습니다
- Windows에서 최고의 성능을 발휘합니다 (Excel, HWP 모두 지원)
- macOS에서는 Excel 기능만 사용 가능합니다

---
*AI 에이전트나 LLM과 함께 사용하는 경우, `oa llm-guide`를 먼저 확인해보세요!*