# PYHWPX 완전 가이드

HWP 문서 자동화를 위한 pyhwpx 라이브러리 완전 참조 가이드

## 📋 목차

1. [소개](#소개)
2. [설치 및 설정](#설치-및-설정)
3. [기본 개념](#기본-개념)
4. [핵심 메서드](#핵심-메서드)
5. [문서 조작](#문서-조작)
6. [서식 및 스타일](#서식-및-스타일)
7. [고급 기능](#고급-기능)
8. [사용 사례별 가이드](#사용-사례별-가이드)
9. [베스트 프랙티스](#베스트-프랙티스)
10. [문제 해결](#문제-해결)

---

## 소개

pyhwpx는 Python에서 HWP(한글) 문서를 자동화할 수 있는 라이브러리입니다. 문서 생성, 편집, 서식 적용, 데이터 변환 등을 프로그래밍 방식으로 처리할 수 있습니다.

### 주요 특징

- 📝 **문서 생성**: 새 HWP 문서 생성 및 기존 문서 편집
- 🎨 **서식 제어**: 스타일, 글꼴, 정렬 등 완전한 서식 제어
- 📊 **표 관리**: 동적 표 생성 및 데이터 입력
- 🖼️ **이미지 삽입**: 다양한 이미지 형식 지원
- 🔄 **변환 기능**: PDF, Markdown 등 다양한 형식으로 변환
- 📬 **메일 머지**: 템플릿 기반 대량 문서 생성

---

## 설치 및 설정

```bash
pip install pyhwpx
```

### 기본 요구사항

- Windows 운영체제
- 한글 2010 이상 설치
- Python 3.6+

---

## 기본 개념

### Hwp 객체 생성과 종료

```python
from pyhwpx import Hwp

# HWP 객체 생성
hwp = Hwp()

# 작업 수행
hwp.insert_text("Hello, HWP!")

# 반드시 종료 (리소스 정리)
hwp.quit()
```

⚠️ **중요**: `hwp.quit()`을 반드시 호출하여 리소스를 정리해야 합니다.

---

## 핵심 메서드

### 1. 객체 생성 및 제어

#### `Hwp()`

+ **목적**: HWP 객체 생성
+ **언제**: 새 문서 작업을 시작할 때
+ **왜**: HWP 애플리케이션과의 연결을 설정

```python
hwp = Hwp()  # 새 HWP 인스턴스 생성
```

#### `quit()`

+ **목적**: HWP 애플리케이션 종료
+ **언제**: 모든 작업 완료 후
+ **왜**: 메모리 해제 및 리소스 정리

```python
hwp.quit()  # 반드시 호출!
```

### 2. 파일 조작

#### `save(file_path)`

+ **목적**: 현재 문서를 HWP 파일로 저장
+ **언제**: 새 문서를 처음 저장할 때
+ **왜**: 작업 내용을 파일로 보존

```python
hwp.save("output/document.hwp")
```

#### `save_as(file_path, format="HWP")`

+ **목적**: 다른 이름 또는 형식으로 저장
+ **언제**: PDF 변환, 백업 생성 시
+ **왜**: 다양한 형식 지원 및 파일 관리

```python
hwp.save_as("output/document.pdf", format="PDF")
hwp.save_as("backup/document_backup.hwp")
```

#### `open(file_path)`

+ **목적**: 기존 HWP 파일 열기
+ **언제**: 기존 문서를 편집할 때
+ **왜**: 템플릿 사용 및 문서 수정

```python
hwp.open("template.hwp")
```

### 3. 텍스트 조작

#### `insert_text(text)`

+ **목적**: 현재 커서 위치에 텍스트 삽입
+ **언제**: 모든 텍스트 입력 작업
+ **왜**: 가장 기본적인 문서 작성 기능

```python
hwp.insert_text("안녕하세요!\n")
hwp.insert_text("두 번째 줄입니다.")
```

#### `find_replace(old_text, new_text, replace_all=False)`

+ **목적**: 텍스트 찾기 및 바꾸기
+ **언제**: 대량 텍스트 수정, 템플릿 변수 치환
+ **왜**: 효율적인 텍스트 변경 작업

```python
# 첫 번째 발견된 텍스트만 변경
count = hwp.find_replace("구 텍스트", "새 텍스트")

# 모든 발견된 텍스트 변경
count = hwp.find_replace("{name}", "김철수", replace_all=True)
```

#### `find(text)`

+ **목적**: 텍스트 검색
+ **언제**: 특정 위치로 이동해야 할 때
+ **왜**: 정확한 위치에서 작업 수행

```python
if hwp.find("특정 텍스트"):
    hwp.insert_text(" - 찾았음!")
```

#### `get_selected_text()`

+ **목적**: 현재 선택된 텍스트 가져오기
+ **언제**: 문서 내용을 분석할 때
+ **왜**: 기존 내용 확인 및 처리

```python
hwp.run("SelectLineEnd")
selected = hwp.get_selected_text()
```

---

## 문서 조작

### 1. 커서 이동 (run 메서드)

#### 기본 이동 명령어


| 명령어 | 목적 | 언제 사용 | 예시 |
|--------|------|-----------|------|
| `MoveDocBegin` | 문서 시작으로 이동 | 전체 문서 작업 전 | `hwp.run("MoveDocBegin")` |
| `MoveDocEnd` | 문서 끝으로 이동 | 내용 추가 시 | `hwp.run("MoveDocEnd")` |
| `MoveLineBegin` | 줄 시작으로 이동 | 줄 단위 편집 시 | `hwp.run("MoveLineBegin")` |
| `MoveLineEnd` | 줄 끝으로 이동 | 줄 끝에 추가 시 | `hwp.run("MoveLineEnd")` |
| `MoveToNextLine` | 다음 줄로 이동 | 줄 단위 순회 시 | `hwp.run("MoveToNextLine")` |
| `MoveToPrevLine` | 이전 줄로 이동 | 역방향 순회 시 | `hwp.run("MoveToPrevLine")` |
| `MoveToNextPara` | 다음 단락으로 이동 | 단락 단위 작업 | `hwp.run("MoveToNextPara")` |


#### 선택 명령어


| 명령어 | 목적 | 언제 사용 | 예시 |
|--------|------|-----------|------|
| `SelectLineEnd` | 줄 끝까지 선택 | 줄 전체 서식 적용 | `hwp.run("SelectLineEnd")` |
| `SelectParaEnd` | 단락 끝까지 선택 | 단락 서식 적용 | `hwp.run("SelectParaEnd")` |
| `SelectWord` | 단어 선택 | 단어 단위 서식 | `hwp.run("SelectWord")` |
| `SelectToDocEnd` | 문서 끝까지 선택 | 대량 서식 적용 | `hwp.run("SelectToDocEnd")` |


### 2. 위치 정보

#### `get_cursor_pos()`

+ **목적**: 현재 커서 위치 확인
+ **언제**: 위치 기반 작업 제어 시
+ **왜**: 정확한 위치 제어 및 루프 제어

```python
pos1 = hwp.get_cursor_pos()
hwp.run("MoveToNextPara")
pos2 = hwp.get_cursor_pos()

if pos1 == pos2:  # 위치가 변하지 않음 = 문서 끝
    break
```

---

## 서식 및 스타일

### 1. 스타일 적용

#### `set_style(style_name)`

**목적**: 선택된 텍스트에 스타일 적용
**언제**: 제목, 본문 등 스타일 구분 시
**왜**: 일관된 문서 서식 및 가독성 향상

```python
# 제목 스타일 적용
hwp.insert_text("문서 제목")
hwp.run("MoveToPrevLine")
hwp.run("SelectLineEnd")
hwp.set_style("제목 1")
hwp.run("MoveDocEnd")
```

#### 주요 스타일 목록

- `"제목 1"` - 큰 제목 (H1)
- `"제목 2"` - 중간 제목 (H2)
- `"제목 3"` - 소제목 (H3)
- `"본문"` - 일반 본문
- `"개조식"` - 불릿 리스트
- `"번호목록"` - 번호 리스트

#### `get_style()`

+ **목적**: 현재 위치의 스타일 확인
+ **언제**: 스타일 기반 문서 분석 시
+ **왜**: 조건부 처리 및 문서 구조 파악

```python
current_style = hwp.get_style()
if current_style == "제목 1":
    print("현재 위치는 제목입니다")
```

### 2. 문자 서식

#### `run("CharShapeBold")`

+ **목적**: 선택된 텍스트를 굵게 처리
+ **언제**: 강조 텍스트 생성 시
+ **왜**: 시각적 강조 효과

```python
hwp.insert_text("중요한 내용")
hwp.run("MoveToPrevLine")
hwp.run("SelectLineEnd")
hwp.run("CharShapeBold")
```

---

## 고급 기능

### 1. 표 관리

#### `insert_table(rows, cols)`

+ **목적**: 표 생성
+ **언제**: 데이터를 표 형태로 표시할 때
+ **왜**: 구조화된 데이터 표현

```python
# 3행 4열 표 생성
hwp.insert_table(rows=3, cols=4)
```

#### `goto_addr(cell_address)`

+ **목적**: 특정 표 셀로 이동
+ **언제**: 표 데이터 입력 시
+ **왜**: 정확한 셀 위치 제어

```python
# a1 셀로 이동 (첫 번째 행, 첫 번째 열)
hwp.goto_addr("a1")
hwp.insert_text("헤더1")

# b2 셀로 이동 (두 번째 행, 두 번째 열)
hwp.goto_addr("b2")
hwp.insert_text("데이터")
```

#### `run("TableOutside")`

+ **목적**: 표 밖으로 이동
+ **언제**: 표 작업 완료 후
+ **왜**: 표 이후 내용 추가

```python
hwp.run("TableOutside")
hwp.insert_text("\n표 아래 내용")
```

### 2. 이미지 삽입

#### `insert_picture(image_path)`

+ **목적**: 이미지 파일 삽입
+ **언제**: 시각적 요소 추가 시
+ **왜**: 문서의 시각적 효과 향상

```python
hwp.insert_picture("images/chart.png")
# 지원 형식: JPG, PNG, GIF, BMP
```

### 3. 문서 변환

#### PDF 변환

```python
hwp.save_as("document.pdf", format="PDF")
```

#### Markdown 변환 (고급)

- 스타일 기반 변환
- 제목 구조 유지
- 목록 형식 변환

---

## 사용 사례별 가이드

### 1. 새 문서 생성

+ **언제**: 완전히 새로운 문서 작성
+ **왜**: 자동화된 문서 생성

```python
from pyhwpx import Hwp

hwp = Hwp()
hwp.insert_text("새 문서 내용")
hwp.save("new_document.hwp")
hwp.quit()
```

### 2. 템플릿 활용

+ **언제**: 정형화된 문서 생성
+ **왜**: 일관된 형식 유지 및 효율성

```python
hwp = Hwp()

# 템플릿 데이터
data = {"{name}": "김철수", "{date}": "2024-01-01"}

# 템플릿 텍스트 작성
hwp.insert_text("안녕하세요, {name}님! 날짜: {date}")

# 변수 치환
hwp.run("MoveDocBegin")
for placeholder, value in data.items():
    hwp.find_replace(placeholder, value, replace_all=True)

hwp.save("personalized.hwp")
hwp.quit()
```

### 3. 기존 문서 편집

+ **언제**: 기존 파일 수정
+ **왜**: 템플릿 기반 작업, 내용 업데이트

```python
hwp = Hwp()
hwp.open("existing.hwp")

# 문서 끝에 내용 추가
hwp.run("MoveDocEnd")
hwp.insert_text("\n추가 내용")

hwp.save_as("updated.hwp")
hwp.quit()
```

### 4. 데이터 기반 표 생성

+ **언제**: Excel 데이터를 HWP 표로 변환
+ **왜**: 보고서 자동화

```python
import pandas as pd

# 데이터 준비
df = pd.read_excel("data.xlsx")

hwp = Hwp()

# 표 생성 (헤더 + 데이터 행)
hwp.insert_table(rows=len(df)+1, cols=len(df.columns))

# 헤더 입력
for col_idx, col_name in enumerate(df.columns):
    hwp.goto_addr(f"{chr(ord('a')+col_idx)}1")
    hwp.insert_text(col_name)

# 데이터 입력
for row_idx, (_, row) in enumerate(df.iterrows()):
    for col_idx, value in enumerate(row):
        hwp.goto_addr(f"{chr(ord('a')+col_idx)}{row_idx+2}")
        hwp.insert_text(str(value))

hwp.save("data_table.hwp")
hwp.quit()
```

### 5. 대량 문서 생성 (메일 머지)

+ **언제**: 개인화된 문서 대량 생성
+ **왜**: 인증서, 초대장 등 개별 문서 필요

```python
recipients = [
    {"name": "김철수", "course": "Python"},
    {"name": "이영희", "course": "HWP 자동화"}
]

for recipient in recipients:
    hwp = Hwp()

    hwp.insert_text(f"""
    수료증

    {recipient['name']}님은 {recipient['course']} 과정을
    성실히 수료하였습니다.
    """)

    filename = f"certificate_{recipient['name']}.hwp"
    hwp.save(filename)
    hwp.quit()
```

### 6. 문서 형식 변환

+ **언제**: 다른 형식으로 변환 필요
+ **왜**: 웹 공유, 인쇄, 보관

```python
# PDF 변환
hwp = Hwp()
hwp.open("source.hwp")
hwp.save_as("output.pdf", format="PDF")
hwp.quit()
```

---

## 베스트 프랙티스

### 1. 리소스 관리

```python
# ✅ 올바른 방법 - 반드시 quit() 호출
try:
    hwp = Hwp()
    hwp.insert_text("내용")
    hwp.save("file.hwp")
finally:
    hwp.quit()

# 또는 context manager 패턴 사용 (권장)
class HwpContext:
    def __enter__(self):
        self.hwp = Hwp()
        return self.hwp

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.hwp.quit()

with HwpContext() as hwp:
    hwp.insert_text("안전한 사용")
    hwp.save("safe.hwp")
```

### 2. 오류 처리

```python
try:
    hwp = Hwp()
    hwp.open("존재하지않는파일.hwp")
except Exception as e:
    print(f"파일 열기 실패: {e}")
finally:
    if 'hwp' in locals():
        hwp.quit()
```

### 3. 대량 작업 시 효율성

```python
# ✅ 효율적 - 한 번에 많은 작업
hwp = Hwp()
for item in large_data:
    hwp.insert_text(f"{item}\n")
hwp.save("bulk.hwp")
hwp.quit()

# ❌ 비효율적 - 매번 객체 생성
for item in large_data:
    hwp = Hwp()
    hwp.insert_text(f"{item}\n")
    hwp.save(f"item_{item}.hwp")
    hwp.quit()
```

### 4. 스타일 적용 패턴

```python
def apply_heading(hwp, text, level=1):
    """제목 적용 헬퍼 함수"""
    hwp.insert_text(text)
    hwp.run("MoveToPrevLine")
    hwp.run("SelectLineEnd")
    hwp.set_style(f"제목 {level}")
    hwp.run("MoveDocEnd")
    hwp.insert_text("\n\n")

# 사용
hwp = Hwp()
apply_heading(hwp, "문서 제목", 1)
apply_heading(hwp, "섹션 제목", 2)
```

---

## 문제 해결

### 1. 일반적인 오류

#### "HWP 애플리케이션을 찾을 수 없음"

+ **원인**: 한글 프로그램이 설치되지 않음
+ **해결**: 한글 2010 이상 설치 필요

#### "파일 저장 실패"

+ **원인**: 경로가 존재하지 않음, 권한 부족
+ **해결**:

```python
import os
os.makedirs("output", exist_ok=True)  # 폴더 생성
hwp.save("output/file.hwp")
```

#### "텍스트 찾기 실패"

+ **원인**: 대소문자, 공백 차이
+ **해결**: 정확한 텍스트 사용, 디버깅으로 확인

### 2. 성능 문제

#### 대량 데이터 처리 시 느림

**해결**:

- 불필요한 GUI 업데이트 최소화
- 배치 작업으로 처리
- 중간 저장으로 메모리 관리

### 3. 디버깅 팁

```python
# 현재 위치 확인
print(f"커서 위치: {hwp.get_cursor_pos()}")

# 현재 스타일 확인
print(f"현재 스타일: {hwp.get_style()}")

# 선택된 텍스트 확인
hwp.run("SelectLineEnd")
print(f"선택된 텍스트: {hwp.get_selected_text()}")
```

---

## 참고 자료

### 메서드 빈도 순 정리 (실제 사용량 기준)

1. **매우 높음**: `insert_text()`, `save()`, `quit()`
2. **높음**: `run()`, `find_replace()`, `set_style()`
3. **중간**: `open()`, `save_as()`, `insert_table()`, `goto_addr()`
4. **낮음**: `insert_picture()`, `get_style()`, `get_cursor_pos()`

### 학습 순서 추천

1. **기초**: 객체 생성 → 텍스트 입력 → 저장
2. **중급**: 스타일 적용 → 찾기/바꾸기 → 표 생성
3. **고급**: 이미지 삽입 → PDF 변환 → 메일 머지

