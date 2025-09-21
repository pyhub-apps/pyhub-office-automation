# [Enhancement] Click → Typer 마이그레이션 완료 필요 (13개 명령어 남음)

## 📋 **이슈 요약**
`pyhub_office_automation/excel/` 디렉토리에 35개의 Python 파일이 있지만, 현재 25개만 `oa excel` CLI를 통해 접근 가능합니다. 남은 10개 파일은 아직 Click 프레임워크를 사용하고 있어서 Typer로 전환이 필요합니다.

## 🎯 **목표**
- **현재**: 25/35 = 71.4% 완성
- **목표**: 35/35 = 100% 완성
- 모든 Excel 자동화 명령어를 `oa excel` 명령어를 통해 일관된 방식으로 접근

## ✅ **완료된 작업 (2025-01-21)**

### **Pivot 명령어 3개 전환 완료**
- ✅ `pivot_delete.py` → `oa excel pivot-delete`
- ✅ `pivot_list.py` → `oa excel pivot-list`
- ✅ `pivot_refresh.py` → `oa excel pivot-refresh`

### **main.py 업데이트 완료**
- ✅ 3개 새로운 import 및 명령어 등록
- ✅ Excel list 명령어에 3개 추가

**검증 결과**:
```bash
❯ oa excel --help | grep -c "│"
25  # 총 25개 명령어 등록됨
```

## ❌ **남은 작업 (13개 파일)**

### **🔴 Phase 1: Shape 명령어 (6개) - 복잡한 옵션 구조**

현재 `main.py`에서 주석 처리된 상태:
```python
# Shape Commands (Click 기반이므로 Typer 변환 필요)
# excel_app.command("shape-add")(shape_add)
# excel_app.command("shape-delete")(shape_delete)
# excel_app.command("shape-format")(shape_format)
# excel_app.command("shape-group")(shape_group)
# excel_app.command("shape-list")(shape_list)
# excel_app.command("textbox-add")(textbox_add)
```

**파일 상태**:
1. `shape_add.py` - 14개 복잡한 옵션 (도형 유형, 스타일 프리셋 등)
2. `shape_delete.py` - 11개 옵션 (안전 삭제 기능)
3. `shape_format.py` - 복잡한 서식 옵션들
4. `shape_group.py` - 도형 그룹화 기능
5. `shape_list.py` - 도형 목록 조회
6. `textbox_add.py` - 텍스트박스 생성

### **🔴 Phase 2: Slicer 명령어 (4개) - typing.Any 에러**

현재 `main.py`에서 주석 처리된 상태:
```python
# Slicer Commands (임시 주석 - typing.Any 에러)
# excel_app.command("slicer-add")(slicer_add)
# excel_app.command("slicer-connect")(slicer_connect)
# excel_app.command("slicer-list")(slicer_list)
# excel_app.command("slicer-position")(slicer_position)
```

**파일 상태**:
1. `slicer_add.py` - 피벗테이블 연동 슬라이서 생성
2. `slicer_connect.py` - 슬라이서 연결 관리
3. `slicer_list.py` - 슬라이서 목록 조회
4. `slicer_position.py` - 슬라이서 위치 조정

### **🟡 Phase 3: 추가 Pivot 명령어 (3개) - 전환 완료, 등록 필요**

파일은 존재하지만 main.py에 등록되지 않은 상태:
1. `pivot_delete.py` ✅ (이미 전환 완료)
2. `pivot_list.py` ✅ (이미 전환 완료)
3. `pivot_refresh.py` ✅ (이미 전환 완료)

## 🚧 **기술적 과제**

### **1. 복잡한 Click 옵션 구조 변환**
```python
# 변환 전 (Click)
@click.option('--shape-type', default='rectangle',
              type=click.Choice(list(SHAPE_TYPES.keys())),
              help='도형 유형 (기본값: rectangle)')

# 변환 후 (Typer) - 방법 검토 필요
shape_type: str = typer.Option("rectangle", "--shape-type",
                              help="도형 유형 (rectangle, oval, line, arrow 등)")
```

### **2. typing.Any 에러 해결**
Slicer 명령어들에서 발생하는 타입 힌트 관련 에러 해결 필요

### **3. is_flag=True 변환 패턴**
```python
# Click
@click.option('--dry-run', is_flag=True)

# Typer
dry_run: bool = typer.Option(False, "--dry-run")
```

## 📋 **작업 계획**

### **Phase 1: Shape 명령어 전환 (높은 우선순위)**
- [ ] `shape_add.py` Click → Typer 전환
- [ ] `shape_delete.py` Click → Typer 전환
- [ ] `shape_format.py` Click → Typer 전환
- [ ] `shape_group.py` Click → Typer 전환
- [ ] `shape_list.py` Click → Typer 전환
- [ ] `textbox_add.py` Click → Typer 전환
- [ ] main.py에서 6개 명령어 주석 해제 및 등록

### **Phase 2: Slicer 명령어 전환 (중간 우선순위)**
- [ ] typing.Any 에러 원인 분석
- [ ] `slicer_add.py` Click → Typer 전환
- [ ] `slicer_connect.py` Click → Typer 전환
- [ ] `slicer_list.py` Click → Typer 전환
- [ ] `slicer_position.py` Click → Typer 전환
- [ ] main.py에서 4개 명령어 주석 해제 및 등록

### **Phase 3: 최종 검증 및 문서화**
- [ ] `oa excel --help` 명령어로 35개 명령어 확인
- [ ] 각 명령어 `--help` 옵션 정상 동작 확인
- [ ] Excel list 명령어에 10개 추가
- [ ] README 업데이트

## 🔍 **성공 기준**

### **기능 테스트**
```bash
# 모든 명령어 등록 확인
❯ oa excel --help | grep -c "│.*command"
35  # 목표: 35개 명령어

# 개별 명령어 도움말 확인
❯ oa excel shape-add --help
# Typer 기반 도움말 정상 출력

❯ oa excel slicer-list --help
# Typer 기반 도움말 정상 출력
```

### **일관성 확인**
- 모든 명령어가 동일한 Typer 패턴 사용
- 에러 처리: `raise typer.Exit(1)`
- 출력: `typer.echo(json.dumps(...))`
- 옵션 정의: `typer.Option(...)` 사용

## 📚 **참고 자료**

### **성공한 전환 패턴 (Pivot 명령어 기준)**
완료된 `pivot_delete.py`, `pivot_list.py`, `pivot_refresh.py` 파일을 참고하여 동일한 패턴 적용

### **관련 파일**
- 📄 `click-to-typer-migration-status.md` - 상세 진행 상황
- 📁 `pyhub_office_automation/excel/` - 전환 대상 파일들
- 📄 `pyhub_office_automation/cli/main.py` - 명령어 등록 파일

## 🏷️ **라벨**
- `enhancement`
- `cli`
- `typer`
- `migration`
- `good first issue` (개별 파일 전환은 초보자도 가능)

## 👥 **담당자**
- **현재 진행자**: @allieus
- **리뷰어**: 필요 시 배정

---
**우선순위**: High
**난이도**: Medium (복잡한 옵션 구조)
**예상 소요시간**: 2-3 시간 (Shape) + 1-2 시간 (Slicer) + 1시간 (최종 검증)