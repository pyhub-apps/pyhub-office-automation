## python 직접 실행

+ oa 가 지원하는 기능 외에 추가로 데이터 변환/분석이 필요하면 아래 경로의 python을 활용해.
    - python 경로 : {python_path}
+ 설치되지 않은 라이브러리는 `{python_path} -m pip install 팩키지명` 명령으로 설치해
+ matplotlib 차트 생성에서는 Malgun Gothic 폰트를 사용하고, 300dpi 로 생성하자.

### Python 사용 예시

```python
# 한글 폰트 설정 (matplotlib)
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm

# 한글 폰트 설정
plt.rcParams['font.family'] = 'Malgun Gothic'
plt.rcParams['axes.unicode_minus'] = False

# 고해상도 설정
plt.rcParams['figure.dpi'] = 300
plt.rcParams['savefig.dpi'] = 300
```

### 대용량 데이터 처리

```python
# 여러 Excel 파일 일괄 처리 (10개 이상 파일 시)
import pandas as pd
from pathlib import Path

def process_multiple_files(file_pattern):
    all_data = []
    for file_path in Path().glob(file_pattern):
        df = pd.read_excel(file_path)
        df['source_file'] = file_path.name
        all_data.append(df)

    return pd.concat(all_data, ignore_index=True)

# 사용 예시
combined_data = process_multiple_files("data/*.xlsx")
```

### 추천 라이브러리

- **pandas**: Excel/CSV 데이터 처리
- **openpyxl**: Excel 파일 읽기/쓰기
- **matplotlib**: 차트 생성
- **seaborn**: 통계 차트
- **numpy**: 수치 계산