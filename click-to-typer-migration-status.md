# Click → Typer 마이그레이션 현황 (2025-01-21)

## 📋 **개요**
`pyhub_office_automation/excel/` 디렉토리의 모든 명령어가 `oa excel` CLI를 통해 접근 가능하도록 하는 작업의 현재 진행 상황입니다.

## ✅ **완료된 작업**

### **Pivot 명령어 3개 전환 완료**
다음 3개 파일을 Click에서 Typer로 성공적으로 전환했습니다:

1. **`pivot_delete.py`** → **`pivot-delete`** ✅
   - 피벗테이블 안전 삭제 (확인 플래그 필수)
   - 피벗캐시 삭제 옵션 포함

2. **`pivot_list.py`** → **`pivot-list`** ✅
   - 워크북 내 모든 피벗테이블 목록 조회
   - 상세 정보 포함 옵션

3. **`pivot_refresh.py`** → **`pivot-refresh`** ✅
   - 개별 또는 전체 피벗테이블 새로고침
   - 소스 데이터 변경 반영

### **main.py 업데이트 완료**
- 3개 새로운 import 추가
- 3개 새로운 명령어 등록: `excel_app.command("pivot-*")`
- Excel list 명령어에 3개 명령어 추가

### **검증 결과**
```bash
❯ oa excel --help | grep -E "pivot-"
│ chart-pivot-create  피벗테이블을 기반으로 동적 피벗차트를 생성합니다.        │
│ pivot-configure     피벗테이블의 필드 배치와 집계 함수를 구성합니다.         │
│ pivot-create        소스 데이터에서 피벗테이블을 생성합니다.                 │
│ pivot-delete        지정된 피벗테이블을 삭제합니다.                          │
│ pivot-list          워크북 내 모든 피벗테이블의 목록과 정보를 조회합니다.    │
│ pivot-refresh       피벗테이블의 데이터를 새로고침합니다.                    │
```

## 📊 **현재 상태**

### **등록된 명령어 현황**
- **총 Excel 명령어**: 25개 (이전 22개 → 현재 25개)
- **활성화율**: 25/35 = **71.4%** (이전 63%에서 향상)
- **새로 추가된 명령어**: 3개

### **프레임워크별 분류**

#### ✅ **Typer 전환 완료 (22개)**
```
- workbook-* (4개): list, open, create, info
- sheet-* (4개): activate, add, delete, rename
- range-* (2개): read, write
- table-* (2개): read, write
- chart-* (7개): add, configure, delete, export, list, pivot-create, position
- pivot-* (3개): configure, create, delete, list, refresh (✅ 새로 추가)
```

#### ❌ **Click 코드 남은 파일들 (13개)**

**Shape 명령어 (6개)** - 복잡한 옵션 구조
```python
# 현재 main.py에서 주석 처리됨
# excel_app.command("shape-add")(shape_add)
# excel_app.command("shape-delete")(shape_delete)
# excel_app.command("shape-format")(shape_format)
# excel_app.command("shape-group")(shape_group)
# excel_app.command("shape-list")(shape_list)
# excel_app.command("textbox-add")(textbox_add)
```

**Slicer 명령어 (4개)** - typing.Any 에러 및 복잡한 구조
```python
# 현재 main.py에서 주석 처리됨
# excel_app.command("slicer-add")(slicer_add)
# excel_app.command("slicer-connect")(slicer_connect)
# excel_app.command("slicer-list")(slicer_list)
# excel_app.command("slicer-position")(slicer_position)
```

**추가 Pivot 명령어 (3개)** - 전환되지 않음
```python
# 파일은 존재하지만 main.py에 등록되지 않음
pivot_delete.py   # ❌ Click 코드 (복잡한 옵션들)
pivot_list.py     # ❌ Click 코드
pivot_refresh.py  # ❌ Click 코드
```

## 🔧 **남은 작업**

### **Phase 1: Shape 명령어 전환 (6개)**
**복잡도**: 🔴 **높음** - 많은 옵션과 Choice 타입들

```python
# shape_add.py 예시 - 복잡한 옵션 구조
@click.option('--shape-type', default='rectangle',
              type=click.Choice(list(SHAPE_TYPES.keys())),
              help='도형 유형 (기본값: rectangle)')
@click.option('--style-preset',
              type=click.Choice(['none', 'background', 'title-box', 'chart-box', 'slicer-box']),
              default='none',
              help='뉴모피즘 스타일 프리셋 (기본값: none)')
```

**전환 필요 파일들**:
1. `shape_add.py` - 도형 생성 (14개 옵션)
2. `shape_delete.py` - 도형 삭제 (11개 옵션)
3. `shape_format.py` - 도형 서식 (복잡한 서식 옵션들)
4. `shape_group.py` - 도형 그룹화
5. `shape_list.py` - 도형 목록 조회
6. `textbox_add.py` - 텍스트박스 추가

### **Phase 2: Slicer 명령어 전환 (4개)**
**복잡도**: 🔴 **높음** - typing.Any 에러 및 복잡한 피벗테이블 연동

```python
# 현재 main.py 주석 이유: "임시 주석 - typing.Any 에러"
```

**전환 필요 파일들**:
1. `slicer_add.py` - 슬라이서 생성 (피벗테이블 연동)
2. `slicer_connect.py` - 슬라이서 연결
3. `slicer_list.py` - 슬라이서 목록 조회
4. `slicer_position.py` - 슬라이서 위치 조정

## 🚧 **기술적 과제**

### **Click → Typer 전환 시 주요 이슈들**

1. **복잡한 Choice 타입 변환**
   ```python
   # Click
   type=click.Choice(['option1', 'option2'])

   # Typer 변환 필요
   # 방법1: 문자열로 처리 후 검증
   # 방법2: Enum 사용
   ```

2. **is_flag=True 처리**
   ```python
   # Click
   @click.option('--flag', is_flag=True)

   # Typer
   flag: bool = typer.Option(False, "--flag")
   ```

3. **복잡한 옵션명 매핑**
   ```python
   # Click
   @click.option('--format', 'output_format', ...)

   # Typer
   output_format: str = typer.Option(..., "--format", ...)
   ```

4. **typing.Any 에러 해결**
   - Slicer 명령어들에서 발생하는 타입 관련 에러
   - 적절한 타입 힌트 적용 필요

## 📋 **다음 작업 계획**

### **우선순위 1: Shape 명령어 전환**
- [ ] `shape_add.py` 전환 및 테스트
- [ ] `shape_delete.py` 전환 및 테스트
- [ ] `shape_format.py` 전환 및 테스트
- [ ] `shape_group.py` 전환 및 테스트
- [ ] `shape_list.py` 전환 및 테스트
- [ ] `textbox_add.py` 전환 및 테스트

### **우선순위 2: Slicer 명령어 전환**
- [ ] typing.Any 에러 원인 분석 및 해결
- [ ] `slicer_add.py` 전환 및 테스트
- [ ] `slicer_connect.py` 전환 및 테스트
- [ ] `slicer_list.py` 전환 및 테스트
- [ ] `slicer_position.py` 전환 및 테스트

### **우선순위 3: main.py 최종 업데이트**
- [ ] 모든 Shape 명령어 주석 해제 및 등록
- [ ] 모든 Slicer 명령어 주석 해제 및 등록
- [ ] Excel list 명령어에 10개 명령어 추가
- [ ] 최종 검증: 35/35 = 100% 완성

## 🎯 **최종 목표**
```bash
# 목표: 모든 35개 파일이 oa excel 명령어로 접근 가능
❯ oa excel --help | wc -l
# 현재: 25개 명령어
# 목표: 35개 명령어 (100% 완성)
```

## 📝 **참고사항**

### **성공한 전환 패턴 (Pivot 명령어 기준)**
```python
# 1. Import 변경
from typing import Optional
import typer
from pyhub_office_automation.version import get_version

# 2. 함수 정의 변경
def command_name(
    file_path: Optional[str] = typer.Option(None, "--file-path", help="..."),
    use_active: bool = typer.Option(False, "--use-active", help="..."),
    # ...
):

# 3. 출력 변경
typer.echo(json.dumps(response, ensure_ascii=False, indent=2))

# 4. 에러 처리 변경
raise typer.Exit(1)
```

### **main.py 등록 패턴**
```python
# Import
from pyhub_office_automation.excel.command_name import command_name

# 등록
excel_app.command("command-name")(command_name)

# List에 추가
{"name": "command-name", "description": "설명", "category": "카테고리"}
```

---
**작성일**: 2025-01-21
**작성자**: Claude Code
**상태**: 진행 중 (71.4% 완성)