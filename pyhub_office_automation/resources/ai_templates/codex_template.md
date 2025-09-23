## Codex CLI 특화 기능

### 인라인 코드 실행 최적화

Codex CLI의 인라인 코드 실행 특성을 활용한 `oa` 명령 사용법:

```python
# Codex 환경에서 oa 명령 실행
import subprocess
import json

def run_oa_command(command_args):
    """oa 명령을 실행하고 JSON 결과를 반환"""
    result = subprocess.run(['oa'] + command_args,
                          capture_output=True, text=True)
    if result.returncode == 0:
        try:
            return json.loads(result.stdout)
        except json.JSONDecodeError:
            return {"raw_output": result.stdout}
    else:
        return {"error": result.stderr}

# 사용 예시
workbooks = run_oa_command(['excel', 'workbook-list', '--format', 'json'])
print(f"현재 {len(workbooks.get('workbooks', []))}개의 워크북이 열려있습니다.")
```

### 코드 블록 기반 워크플로우

```python
# 1단계: 상황 파악
workbook_info = run_oa_command(['excel', 'workbook-list', '--detailed'])

# 2단계: 데이터 읽기
data = run_oa_command(['excel', 'range-read',
                      '--sheet', 'Sheet1',
                      '--range', 'A1:C100',
                      '--format', 'json'])

# 3단계: 데이터 처리 (pandas 활용)
import pandas as pd
import io

if 'csv_data' in data:
    df = pd.read_csv(io.StringIO(data['csv_data']))

    # 간단한 분석
    summary = df.describe()

    # 결과를 다시 Excel에 쓰기
    summary_json = summary.to_json()
    run_oa_command(['excel', 'range-write',
                   '--sheet', 'Summary',
                   '--range', 'A1',
                   '--data', summary_json])
```

### Codex 장점 활용

1. **즉석 코드 생성**: 복잡한 데이터 변환 로직을 바로 작성
2. **API 통합**: `oa` 명령을 Python 함수로 래핑하여 활용
3. **오류 처리**: try-catch로 안전한 명령 실행

### 권장 패턴

```python
class OAHelper:
    """oa 명령어 헬퍼 클래스"""

    @staticmethod
    def get_workbooks():
        """열린 워크북 목록 조회"""
        return run_oa_command(['excel', 'workbook-list', '--format', 'json'])

    @staticmethod
    def read_range(sheet, range_addr, workbook_name=None):
        """범위 데이터 읽기"""
        cmd = ['excel', 'range-read', '--sheet', sheet, '--range', range_addr]
        if workbook_name:
            cmd.extend(['--workbook-name', workbook_name])
        return run_oa_command(cmd)

    @staticmethod
    def create_chart(sheet, data_range, chart_type, title):
        """차트 생성"""
        return run_oa_command(['excel', 'chart-add',
                              '--sheet', sheet,
                              '--data-range', data_range,
                              '--chart-type', chart_type,
                              '--title', title])

# 사용법
helper = OAHelper()
workbooks = helper.get_workbooks()
data = helper.read_range('Sheet1', 'A1:C10')
chart = helper.create_chart('Sheet1', 'A1:C10', 'Column', '판매 데이터')
```