# 데이터 변환 명령어 구현 계획

## 전체 개요

Excel 데이터를 피벗테이블용으로 변환하는 두 가지 새 명령어(`data-analyze`, `data-transform`)의 단계별 구현 계획입니다.

## Phase 1: `data-analyze` 명령어 구현

### 1.1 핵심 기능 설계

#### 파일 구조
```
pyhub_office_automation/excel/
├── data_analyze.py          # 새 명령어 파일
├── data_transform.py        # 새 명령어 파일 (Phase 2)
└── data_utils.py           # 공통 유틸리티
```

#### 주요 분석 기능

1. **데이터 구조 감지**
   - 헤더 행 수 감지
   - 병합된 셀 탐지
   - 데이터 타입 분석
   - 빈 셀 패턴 분석

2. **문제점 식별**
   - Cross-tab 형식 감지
   - Multi-level 헤더 감지
   - 소계 행 감지
   - Wide vs Long 형식 판단

3. **변환 가능성 평가**
   - 피벗테이블 적합성 점수
   - 예상 변환 결과 크기
   - 권장 변환 방법

### 1.2 명령어 인터페이스

```bash
oa excel data-analyze [OPTIONS]

Options:
  --file-path TEXT         Excel 파일 경로
  --workbook-name TEXT     열린 워크북 이름
  --sheet TEXT            시트 이름 (기본: 활성 시트)
  --range TEXT            분석할 범위 (기본: 사용된 범위)
  --output-file TEXT      분석 결과 저장 파일
  --detailed              상세 분석 활성화
  --ai-assist             AI 도구 연동 분석
  --output-format [json|text]  출력 형식
```

### 1.3 출력 형식

```json
{
  "command": "data-analyze",
  "version": "1.0.0",
  "timestamp": "2025-09-22T12:00:00",
  "file_info": {
    "name": "sales_report.xlsx",
    "sheet": "Data",
    "range": "A1:M50"
  },
  "analysis": {
    "data_shape": {"rows": 50, "cols": 13},
    "header_rows": 1,
    "data_rows": 49,
    "format_type": "cross_tab",
    "pivot_ready": false,
    "confidence_score": 0.85,
    "issues": [
      {
        "type": "cross_tab_format",
        "severity": "high",
        "description": "월별 데이터가 열로 배치되어 피벗테이블에 부적합",
        "affected_range": "B1:M1",
        "solution": "unpivot"
      },
      {
        "type": "merged_cells",
        "severity": "medium",
        "description": "3개의 병합된 셀 발견",
        "affected_range": "A1:A3",
        "solution": "unmerge"
      }
    ],
    "recommendations": [
      "월별 컬럼(B:M)을 세로 형식으로 변환하세요",
      "A열의 병합된 셀을 해제하고 값을 채워넣으세요"
    ],
    "transformation_impact": {
      "current_shape": [50, 13],
      "expected_shape": [588, 3],
      "data_loss_risk": "low",
      "processing_time_estimate": "2-5초"
    }
  },
  "ai_analysis": {
    "available": true,
    "strategy": "unpivot 변환 후 병합 해제 권장",
    "code_suggestion": "pd.melt() 함수 사용",
    "complexity_score": 0.6
  }
}
```

### 1.4 구현 단계

#### Step 1.1: 기본 골격 구성 (2-3일)
- [ ] `data_analyze.py` 파일 생성
- [ ] Click 기반 CLI 인터페이스
- [ ] 기본 워크북 연결 로직
- [ ] 범위 읽기 및 기본 분석

#### Step 1.2: 분석 알고리즘 구현 (5-7일)
- [ ] 데이터 구조 감지 로직
- [ ] 5가지 주요 문제 감지 알고리즘
- [ ] 피벗테이블 적합성 평가 점수
- [ ] 변환 권장사항 생성

#### Step 1.3: AI 도구 통합 (3-4일)
- [ ] 로컬 AI 도구 감지
- [ ] AI 분석 요청 포맷
- [ ] AI 응답 파싱 및 통합
- [ ] 폴백 메커니즘

#### Step 1.4: 테스트 및 검증 (2-3일)
- [ ] 단위 테스트 작성
- [ ] 다양한 데이터 형식 테스트
- [ ] AI 통합 테스트
- [ ] 성능 최적화

## Phase 2: `data-transform` 명령어 구현

### 2.1 변환 엔진 설계

#### 변환 타입별 구현

1. **`unpivot` 변환**
   ```python
   def unpivot_data(df, id_vars, value_vars, var_name='Variable', value_name='Value'):
       return pd.melt(df, id_vars=id_vars, value_vars=value_vars,
                      var_name=var_name, value_name=value_name)
   ```

2. **`unmerge` 변환**
   ```python
   def unmerge_data(df, range_info):
       # 병합된 셀 해제 및 forward fill
       return df.fillna(method='ffill')
   ```

3. **`flatten-headers` 변환**
   ```python
   def flatten_headers(df, header_rows):
       # 다단계 헤더를 단일 헤더로 결합
       new_headers = []
       for col in df.columns:
           # 계층적 헤더 평탄화 로직
       return df
   ```

4. **`remove-subtotals` 변환**
   ```python
   def remove_subtotals(df, subtotal_patterns):
       # 소계 행 식별 및 제거
       return df[~df.apply(is_subtotal_row, axis=1)]
   ```

5. **`auto` 변환**
   ```python
   def auto_transform(df, analysis_result):
       # 분석 결과 기반 자동 변환 적용
       for issue in analysis_result['issues']:
           df = apply_transformation(df, issue['solution'])
       return df
   ```

### 2.2 명령어 인터페이스

```bash
oa excel data-transform [OPTIONS]

Options:
  --file-path TEXT              Excel 파일 경로
  --workbook-name TEXT          열린 워크북 이름
  --source-range TEXT           소스 데이터 범위
  --transform-type [unpivot|unmerge|flatten-headers|remove-subtotals|auto]
  --output-sheet TEXT           결과 시트 이름
  --output-file TEXT            결과 파일 저장 경로
  --id-columns TEXT             ID로 유지할 컬럼 (unpivot용)
  --value-columns TEXT          값 컬럼 지정 (unpivot용)
  --preview                     변환 미리보기 (실제 변환 안함)
  --ai-guidance TEXT            AI 분석 파일 경로
  --backup                      원본 백업 생성
```

### 2.3 구현 단계

#### Step 2.1: 변환 엔진 구현 (7-10일)
- [ ] pandas 기반 변환 함수들
- [ ] 각 변환 타입별 로직 구현
- [ ] 데이터 검증 및 오류 처리
- [ ] 변환 전후 비교 기능

#### Step 2.2: Excel 통합 (4-5일)
- [ ] xlwings 연동 로직
- [ ] 변환 결과 Excel 쓰기
- [ ] 새 시트 생성 및 관리
- [ ] 원본 백업 기능

#### Step 2.3: AI 가이던스 통합 (3-4일)
- [ ] AI 분석 결과 파싱
- [ ] AI 권장사항 적용
- [ ] 복잡한 케이스 AI 위임
- [ ] 사용자 피드백 수집

## Phase 3: 통합 및 최적화

### 3.1 명령어 통합

#### CLI 등록
```python
# pyhub_office_automation/cli/excel.py
@excel_cli.command("data-analyze")
def data_analyze_command(**kwargs):
    from ..excel.data_analyze import data_analyze
    return data_analyze(**kwargs)

@excel_cli.command("data-transform")
def data_transform_command(**kwargs):
    from ..excel.data_transform import data_transform
    return data_transform(**kwargs)
```

### 3.2 워크플로우 최적화

#### 연계 워크플로우
```bash
# 분석 → 변환 → 피벗테이블 생성 파이프라인
oa excel data-analyze --file-path "report.xlsx" --output-file "analysis.json"
oa excel data-transform --ai-guidance "analysis.json" --transform-type "auto"
oa excel pivot-create --source-range "Transformed!A1:C1000" --auto-position
```

### 3.3 성능 최적화

1. **메모리 최적화**: 청크 단위 데이터 처리
2. **속도 최적화**: 벡터화 연산 활용
3. **캐싱**: 분석 결과 및 변환 패턴 캐시
4. **병렬 처리**: 대용량 데이터 병렬 변환

## Phase 4: 문서화 및 배포

### 4.1 사용자 문서

1. **명령어 가이드**: 각 옵션 상세 설명
2. **사용 시나리오**: 실제 업무 사례
3. **문제 해결 가이드**: 일반적 오류 해결
4. **AI 통합 가이드**: AI 도구 활용법

### 4.2 개발자 문서

1. **API 문서**: 내부 함수 문서화
2. **확장 가이드**: 새 변환 타입 추가
3. **테스트 가이드**: 테스트 케이스 작성
4. **성능 분석**: 벤치마크 결과

## 예상 일정

| Phase | 기간 | 주요 마일스톤 |
|-------|------|---------------|
| Phase 1 | 2-3주 | data-analyze 명령어 완성 |
| Phase 2 | 3-4주 | data-transform 명령어 완성 |
| Phase 3 | 1-2주 | 통합 및 최적화 |
| Phase 4 | 1주 | 문서화 및 배포 |
| **총 기간** | **7-10주** | **완전한 기능 출시** |

## 성공 지표

1. **기능 성능**: 1000행 데이터 5초 이내 분석/변환
2. **정확도**: 95% 이상 올바른 문제 감지
3. **사용성**: 비기술 사용자 80% 성공률
4. **AI 통합**: 로컬 AI 도구 90% 가용성
5. **안정성**: 99% 이상 오류 없는 변환

이 계획에 따라 단계적으로 구현하면 사용자들이 Excel 데이터를 쉽게 피벗테이블용으로 변환할 수 있는 강력한 도구를 제공할 수 있습니다.