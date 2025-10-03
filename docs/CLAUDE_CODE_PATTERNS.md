# Claude Code 특화 분석 패턴

## 목차
- [상세 분석 및 체계적 접근](#상세-분석-및-체계적-접근)
- [문제 해결 방법론](#문제-해결-방법론)
- [코드 리뷰 및 최적화](#코드-리뷰-및-최적화)
- [고급 Excel 활용 패턴](#고급-excel-활용-패턴)
- [즉시 분석 패턴 (table-list 활용)](#즉시-분석-패턴-table-list-활용)

---

## 상세 분석 및 체계적 접근

Claude의 깊이 있는 분석 능력을 활용한 Excel 자동화 패턴:

### 코드 품질 및 구조 분석

```python
# Claude Code가 excel 자동화 스크립트를 분석할 때 중점 사항
def analyze_excel_workflow():
    """
    1. 데이터 무결성 검증
    2. 에러 처리 패턴
    3. 성능 최적화 기회
    4. 코드 재사용성
    """

    # 단계별 검증 워크플로우
    steps = [
        "oa excel workbook-list",  # 현황 파악
        "데이터 구조 분석",                    # 스키마 검토
        "비즈니스 로직 검증",                  # 요구사항 부합성
        "성능 및 확장성 검토"                  # 최적화 기회
    ]

    return steps
```

---

## 문제 해결 방법론

### 체계적 디버깅 접근

```bash
# 1. 상황 진단
oa excel workbook-list --format json

# 2. 데이터 구조 분석
oa excel workbook-info  # All details included by default

# 3. 샘플 데이터 검증
oa excel range-read --sheet "Sheet1" --range "A1:E5"

# 4. 에러 재현 및 분석
# (문제가 되는 명령어 단계별 실행)

# 5. 해결책 구현 및 검증
```

---

## 코드 리뷰 및 최적화

### Excel 자동화 코드 리뷰 체크리스트

```python
def review_excel_automation():
    """
    Claude Code의 Excel 자동화 코드 리뷰 포인트
    """
    checklist = {
        "에러 처리": [
            "파일 존재 여부 확인",
            "시트 존재 여부 확인",
            "범위 유효성 검증",
            "데이터 타입 검증"
        ],
        "성능": [
            "대용량 데이터 처리 최적화",
            "메모리 사용량 관리",
            "I/O 작업 최소화",
            "배치 처리 활용"
        ],
        "유지보수성": [
            "모듈화된 함수 설계",
            "설정값 외부화",
            "로깅 및 모니터링",
            "문서화 완성도"
        ]
    }
    return checklist
```

---

## 고급 Excel 활용 패턴

### 복합 데이터 분석 파이프라인

```python
import subprocess
import json
import pandas as pd
from pathlib import Path

class ExcelAnalysisPipeline:
    """체계적인 Excel 데이터 분석 파이프라인"""

    def __init__(self, workbook_name=None):
        self.workbook_name = workbook_name
        self.context = {}

    def analyze_structure(self):
        """데이터 구조 분석"""
        cmd = ['oa', 'excel', 'workbook-info']  # All details included by default
        if self.workbook_name:
            cmd.extend(['--workbook-name', self.workbook_name])

        result = subprocess.run(cmd, capture_output=True, text=True)
        self.context['structure'] = json.loads(result.stdout)
        return self.context['structure']

    def extract_data(self, sheet, range_addr):
        """데이터 추출 및 검증"""
        cmd = ['oa', 'excel', 'range-read',
               '--sheet', sheet, '--range', range_addr, '--format', 'json']
        if self.workbook_name:
            cmd.extend(['--workbook-name', self.workbook_name])

        result = subprocess.run(cmd, capture_output=True, text=True)
        data = json.loads(result.stdout)

        # 데이터 품질 검증
        df = pd.DataFrame(data.get('data', []))
        self.context['data_quality'] = {
            'rows': len(df),
            'columns': len(df.columns),
            'null_count': df.isnull().sum().sum(),
            'duplicates': df.duplicated().sum()
        }

        return df

    def generate_insights(self, df):
        """데이터 인사이트 생성"""
        insights = {
            'summary_stats': df.describe().to_dict(),
            'data_types': df.dtypes.to_dict(),
            'missing_data': df.isnull().sum().to_dict()
        }

        # 비즈니스 인사이트 추가
        if 'sales' in df.columns or '매출' in df.columns:
            sales_col = 'sales' if 'sales' in df.columns else '매출'
            insights['sales_analysis'] = {
                'total_sales': df[sales_col].sum(),
                'avg_sales': df[sales_col].mean(),
                'top_performers': df.nlargest(5, sales_col).to_dict()
            }

        return insights

    def create_dashboard(self, insights):
        """대시보드 차트 생성"""
        charts_created = []

        # 요약 통계 차트
        summary_chart = self._create_summary_chart()
        if summary_chart:
            charts_created.append(summary_chart)

        # 추세 분석 차트
        trend_chart = self._create_trend_chart()
        if trend_chart:
            charts_created.append(trend_chart)

        return charts_created

    def _create_summary_chart(self):
        """요약 차트 생성"""
        cmd = ['oa', 'excel', 'chart-add',
               '--sheet', 'Dashboard',
               '--data-range', 'A1:B10',
               '--chart-type', 'Column',
               '--title', '데이터 요약']

        result = subprocess.run(cmd, capture_output=True, text=True)
        return json.loads(result.stdout) if result.returncode == 0 else None
```

### 문서화 및 지식 관리

#### 자동 문서 생성

```python
def generate_analysis_report(pipeline_results):
    """분석 결과 자동 문서화"""
    report = f"""
# Excel 데이터 분석 보고서

## 데이터 개요
- 워크북: {pipeline_results['workbook']}
- 시트 수: {len(pipeline_results['sheets'])}
- 총 데이터 행: {pipeline_results['total_rows']}

## 데이터 품질 평가
- 결측값: {pipeline_results['missing_values']}%
- 중복값: {pipeline_results['duplicates']}개
- 데이터 완성도: {pipeline_results['completeness']}%

## 주요 인사이트
{pipeline_results['insights']}

## 권장 액션
{pipeline_results['recommendations']}

## 생성된 차트
{pipeline_results['charts']}
"""
    return report
```

---

## 즉시 분석 패턴 (table-list 활용)

### Claude Code + table-list 최적 활용법

**즉시 분석 패턴**:

```bash
# Claude Code가 선호하는 효율적 워크플로우
oa excel table-list --format json
# ☝️ 한 번의 호출로 Claude가 즉시 파악:
# - 테이블 구조 (11개 컬럼: 순위, 게임명, 플랫폼, 발행일, 장르, 퍼블리셔, 판매량x4, 글로벌판매량)
# - 샘플 데이터 (Wii 스포츠 82.74M, 슈퍼 마리오 40.24M 등)
# - 데이터 품질 (998행, 정형화된 숫자 데이터)
# - 비즈니스 컨텍스트 (게임 판매 분석 데이터)

# Claude가 즉시 제안 가능한 분석들:
# 1. "글로벌 판매량 Top 10 막대 차트를 만들어드릴까요?"
# 2. "지역별 판매량 비교 (북미 vs 유럽 vs 일본 vs 기타)는 어떨까요?"
# 3. "장르별 집계나 플랫폼별 분석도 가능합니다."
# 4. "발행 연도별 트렌드 분석도 해볼까요?"
```

### Smart Chart Recommendation Engine

```python
def claude_smart_chart_suggestions(table_data):
    """
    Claude Code가 table-list 데이터를 분석해 최적 차트 추천
    """
    recommendations = []

    # 컬럼 분석 기반 자동 추천
    columns = table_data.get("columns", [])
    sample_data = table_data.get("sample_data", [])

    if "글로벌 판매량" in columns and "게임명" in columns:
        recommendations.append({
            "type": "Column",
            "title": "게임별 글로벌 판매량 Top 10",
            "reason": "순위 데이터와 판매량 수치로 Top 10 시각화 최적",
            "command": "oa excel chart-add --data-range 'GameData[글로벌 판매량]' --chart-type 'Column'"
        })

    if "북미 판매량" in columns and "유럽 판매량" in columns:
        recommendations.append({
            "type": "Scatter",
            "title": "북미 vs 유럽 판매량 상관관계",
            "reason": "두 지역 판매량 간의 상관성 분석",
            "command": "oa excel chart-add --x-range 'GameData[북미 판매량]' --y-range 'GameData[유럽 판매량]'"
        })

    return recommendations

# 실제 활용: Claude가 즉시 적절한 차트 제안
chart_suggestions = claude_smart_chart_suggestions(table_list_response["data"]["tables"][0])
```

### Data Quality Instant Assessment

```python
def claude_data_quality_check(sample_data):
    """
    샘플 데이터만으로 Claude가 즉시 품질 평가
    """
    quality_report = {
        "data_completeness": "✅ NULL 값 없음",
        "data_types": "✅ 숫자 데이터 정상 (41.49, 29.02 등)",
        "business_logic": "✅ 판매량 합계 로직 확인 가능 (지역별 → 글로벌)",
        "recommendations": [
            "발행일을 연도 형식으로 변환하여 시계열 분석",
            "판매량 단위 백만장으로 해석하여 차트 레이블링",
            "상위 게임들의 플랫폼 트렌드 분석 가능"
        ]
    }
    return quality_report
```

---

## Claude Code 장점 활용

### 권장 작업 순서

1. **요구사항 분석**: 비즈니스 목표와 데이터 요구사항 명확화
2. **환경 검증**: `oa excel workbook-list`로 현재 상태 확인
3. **데이터 탐색**: 구조 분석 및 샘플 데이터 검토
4. **분석 설계**: 단계별 분석 프로세스 설계
5. **실행 및 검증**: 각 단계별 결과 검증
6. **결과 정리**: 인사이트 요약 및 액션 아이템 제시

### 핵심 강점

1. **정확한 분석**: 데이터 무결성과 비즈니스 로직 검증
2. **체계적 접근**: 단계별 분석 프로세스 설계
3. **품질 관리**: 코드 리뷰와 최적화 제안
4. **지식 정리**: 자동 문서화와 인사이트 요약

---

## 참고 문서

- [CLAUDE.md](../CLAUDE.md) - AI Agent Quick Reference
- [SHELL_USER_GUIDE.md](./SHELL_USER_GUIDE.md) - Shell Mode 완벽 가이드
- [ADVANCED_FEATURES.md](./ADVANCED_FEATURES.md) - Map Chart, 고급 차트 기능
