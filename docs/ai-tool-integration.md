# AI 도구 통합 가이드

## 개요

pyhub-office-automation은 로컬 AI 도구들과 연동하여 복잡한 데이터 분석 및 변환 작업을 지원합니다. 이 문서는 효과적인 AI 도구 활용 패턴을 설명합니다.

## 지원 AI 도구

### 1. Codex CLI
- **용도**: 복잡한 데이터 구조 분석, 변환 알고리즘 제안
- **설치**: `brew install codex` 또는 공식 설치 가이드
- **특징**: 높은 추론력(reasoning effort: high), 코드 생성 특화

### 2. Gemini CLI
- **용도**: 배치 처리, pandas 코드 생성, 패턴 인식
- **설치**: `brew install gemini` 또는 공식 설치 가이드
- **특징**: 프롬프트 모드(-p), 구체적 코드 예제 생성

### 3. Claude CLI
- **용도**: 사용자 친화적 설명, 문서화, 한국어 지원
- **설치**: 로컬 설치 또는 별칭 설정
- **특징**: 자연어 처리, 문맥 이해, 다국어 지원

## 사용 패턴

### 패턴 1: 데이터 분석 워크플로우

```bash
# 1단계: Codex로 고급 분석
codex exec "다음 Excel 데이터 구조를 분석하고 피벗테이블 적합성을 평가:
- 헤더 구조: A1:L1 (제품명, 1월, 2월, ..., 12월)
- 데이터 범위: A2:L50 (49개 제품의 월별 매출)
- 문제점과 변환 방법 제안"

# 2단계: Gemini로 구체적 코드 생성
gemini -p "교차표 형식의 Excel 데이터를 pandas로 unpivot하는 코드:
- 입력: 제품(행) x 월(열) 형식
- 출력: 제품, 월, 매출액 3열 형식
- 결측치 및 0값 처리 포함"

# 3단계: Claude로 사용자 설명 생성
claude -p "위 데이터 변환 과정을 비기술 사용자에게 한국어로 설명:
- 왜 변환이 필요한지
- 어떤 변화가 일어나는지
- 피벗테이블에서 어떤 이점이 있는지"
```

### 패턴 2: 오류 진단 및 해결

```bash
# Excel 명령어 실행 후 오류 발생 시
oa excel data-analyze --file-path "problem.xlsx" --range "A1:Z100" 2>&1 | \
  codex exec "이 Excel 데이터 분석 오류를 해석하고 해결 방법 제안"

# 해결책 검증
gemini -p "제안된 해결책의 pandas 구현 코드와 테스트 방법"
```

### 패턴 3: 복잡한 변환 시나리오

```bash
# 다단계 변환이 필요한 경우
echo "복잡한 보고서 형식:
- 병합된 헤더 3행
- 월별/분기별 중첩 구조
- 소계 행이 데이터와 혼재" | \
codex exec "단계별 변환 전략 수립"

# 각 단계별 구현
gemini -p "1단계: 병합된 셀 해제 및 헤더 평탄화 pandas 코드"
gemini -p "2단계: 중첩 구조를 단일 계층으로 변환하는 코드"
gemini -p "3단계: 소계 행 식별 및 제거 로직"
```

## 데이터 변환 명령어와 AI 통합

### data-analyze 명령어 확장

```python
# 향후 구현 시 AI 도구 통합 예시
def analyze_with_ai(data_preview, issues):
    if ai_tools_available():
        # Codex로 고급 분석
        codex_analysis = subprocess.run([
            "codex", "exec",
            f"이 데이터 구조를 분석: {data_preview}\n문제점: {issues}"
        ], capture_output=True, text=True)

        # Gemini로 변환 코드 제안
        gemini_code = subprocess.run([
            "gemini", "-p",
            f"다음 문제를 해결하는 pandas 코드: {issues}"
        ], capture_output=True, text=True)

        return {
            "ai_analysis": codex_analysis.stdout,
            "suggested_code": gemini_code.stdout,
            "ai_assisted": True
        }
    return {"ai_assisted": False}
```

### data-transform 명령어 확장

```python
def transform_with_ai_guidance(source_data, transform_type):
    if transform_type == "auto" and ai_tools_available():
        # 복잡한 케이스는 AI에게 위임
        ai_strategy = get_ai_transform_strategy(source_data)
        return apply_ai_strategy(ai_strategy)
    else:
        # 기본 변환 로직 사용
        return standard_transform(source_data, transform_type)
```

## 성능 최적화

### AI 도구 호출 최적화

1. **캐싱**: 동일한 데이터 패턴에 대한 AI 분석 결과 캐시
2. **비동기**: 여러 AI 도구 동시 호출
3. **폴백**: AI 도구 실패 시 기본 알고리즘 사용
4. **배치**: 여러 데이터셋을 한 번에 분석

```bash
# 배치 처리 예시
for file in *.xlsx; do
  echo "Analyzing $file..."
  oa excel data-analyze --file-path "$file" --output-file "${file%.xlsx}_analysis.json"
done | codex exec "이 모든 분석 결과를 종합하여 공통 패턴과 권장사항 도출"
```

## 사용자 가이드

### 초보자 워크플로우

```bash
# 1. 간단한 분석 (AI 없이)
oa excel data-analyze --file-path "my_data.xlsx"

# 2. 문제 발견 시 AI 도움 요청
oa excel data-analyze --file-path "my_data.xlsx" --output-file "analysis.json"
codex exec "$(cat analysis.json) 이 분석 결과를 바탕으로 구체적인 해결 방법 제안"

# 3. 변환 실행
oa excel data-transform --source-range "A1:Z100" --transform-type "auto"
```

### 고급 사용자 워크플로우

```bash
# 복합 분석
{
  echo "=== 데이터 구조 분석 ==="
  oa excel data-analyze --file-path "complex.xlsx" --format json

  echo "=== AI 전략 수립 ==="
  codex exec "위 분석을 바탕으로 최적화된 변환 전략 수립"

  echo "=== 코드 생성 ==="
  gemini -p "제안된 전략의 구체적 pandas 구현"
} | tee analysis_report.txt

# AI 분석 결과 기반 실행
oa excel data-transform --source-range "A1:Z200" --transform-type "auto" \
  --ai-guidance analysis_report.txt
```

## 문제 해결

### 일반적인 문제

1. **AI 도구 미설치**: 기본 알고리즘으로 폴백
2. **AI 응답 지연**: 타임아웃 설정 및 비동기 처리
3. **AI 분석 오류**: 사용자에게 수동 옵션 제공

### 로그 및 디버깅

```bash
# AI 도구 연동 상태 확인
oa info --check-ai-tools

# 상세 로그 활성화
OA_AI_DEBUG=1 oa excel data-analyze --file-path "debug.xlsx"
```

## 향후 계획

1. **AI 도구 자동 감지**: 설치된 도구 자동 인식
2. **커스텀 프롬프트**: 사용자 정의 AI 프롬프트 템플릿
3. **성능 모니터링**: AI 도구별 성능 측정 및 최적화
4. **오프라인 모드**: 로컬 모델 지원

이 가이드는 AI 도구와의 효과적인 통합을 통해 Excel 데이터 변환 작업의 효율성과 정확성을 크게 향상시킬 수 있는 방법을 제시합니다.