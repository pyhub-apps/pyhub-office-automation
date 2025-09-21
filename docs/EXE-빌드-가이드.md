# 🚀 Windows EXE 빌드 가이드

Python 설치 없이 실행 가능한 단일 EXE 파일을 생성하는 완전한 가이드입니다.

## 📋 목차

1. [시스템 요구사항](#-시스템-요구사항)
2. [빌드 환경 설정](#-빌드-환경-설정)
3. [빌드 실행](#-빌드-실행)
4. [빌드 옵션](#-빌드-옵션)
5. [빌드 결과물](#-빌드-결과물)
6. [문제 해결](#-문제-해결)
7. [고급 설정](#-고급-설정)

## 🖥️ 시스템 요구사항

### 필수 요구사항
- **운영체제**: Windows 10/11 (64비트)
- **Python**: 3.13 이상
- **PowerShell**: 5.1 이상 (Windows 10/11 기본 포함)
- **디스크 여유공간**: 최소 2GB (빌드 과정 포함)

### 권장 요구사항
- **RAM**: 8GB 이상 (빌드 속도 향상)
- **CPU**: 4코어 이상 (병렬 처리 지원)
- **Microsoft Excel**: Excel 자동화 기능 테스트용
- **한글(HWP)**: HWP 자동화 기능 테스트용

## ⚙️ 빌드 환경 설정

### 1. 프로젝트 클론
```bash
git clone https://github.com/pyhub-apps/pyhub-office-automation.git
cd pyhub-office-automation
```

### 2. 가상환경 설정
```powershell
# Python 가상환경 생성
python -m venv .venv

# 가상환경 활성화
.venv\Scripts\activate

# 의존성 설치
pip install -r requirements.txt

# 개발 모드로 패키지 설치
pip install -e .
```

### 3. 빌드 도구 설치
```powershell
# PyInstaller 및 빌드 도구 설치
pip install pyinstaller
pip install pillow  # 아이콘 생성용
```

## 🔨 빌드 실행

### 기본 빌드 (권장)
```powershell
# 단일 EXE 파일 생성 (기본 설정)
.\build_windows.ps1

# 또는 명시적 옵션 지정
.\build_windows.ps1 -BuildType onefile
```

### 메타데이터 포함 빌드
```powershell
# 버전 정보 및 메타데이터 포함
.\build_windows.ps1 -BuildType onefile -GenerateMetadata
```

### CI 모드 빌드
```powershell
# 자동화된 환경용 (사용자 입력 없음)
.\build_windows.ps1 -BuildType onefile -CiMode
```

## 🎛️ 빌드 옵션

### 기본 빌드 타입
| 옵션 | 설명 | 결과물 | 권장 용도 |
|------|------|---------|-----------|
| `onedir` | 디렉터리 형태 | `dist/oa/` 폴더 | 개발 및 테스트 |
| `onefile` | 단일 파일 | `dist/oa.exe` | **배포용 (권장)** |

### 추가 옵션
| 매개변수 | 설명 | 예제 |
|----------|------|------|
| `-GenerateMetadata` | 버전 정보 및 메타데이터 포함 | `.\build_windows.ps1 -GenerateMetadata` |
| `-CiMode` | CI/CD 환경용 (사용자 입력 없음) | `.\build_windows.ps1 -CiMode` |
| `-UseSpec` | 기존 spec 파일 사용 | `.\build_windows.ps1 -UseSpec` |
| `-Help` | 도움말 표시 | `.\build_windows.ps1 -Help` |

### 빌드 명령어 예제
```powershell
# 🎯 프로덕션 배포용 (권장)
.\build_windows.ps1 -BuildType onefile -GenerateMetadata

# 🧪 개발 테스트용
.\build_windows.ps1 -BuildType onedir

# 🤖 CI/CD 자동화용
.\build_windows.ps1 -BuildType onefile -CiMode -GenerateMetadata

# 📋 기존 설정 재사용
.\build_windows.ps1 -UseSpec
```

## 📦 빌드 결과물

### 생성되는 파일들

#### OneFIle 모드 (단일 EXE)
```
dist/
├── oa.exe                 # ✨ 메인 실행 파일 (약 50-100MB)
└── build_metadata.json    # 📊 빌드 정보 (옵션)
```

#### OneDir 모드 (디렉터리)
```
dist/
└── oa/
    ├── oa.exe              # 메인 실행 파일
    ├── _internal/          # 라이브러리 파일들
    │   ├── python313.dll
    │   ├── xlwings/
    │   ├── pyhwpx/
    │   └── ...
    └── assets/             # 아이콘 및 리소스
```

### 빌드 메타데이터 (GenerateMetadata 옵션)
```json
{
  "build_time": "2024-09-22T08:30:15+09:00",
  "version": "7.2539.67",
  "git_commit": "b826940",
  "build_type": "onefile",
  "file_size": "67,108,864",
  "sha256": "a1b2c3d4e5f6...",
  "python_version": "3.13.0",
  "platform": "Windows-10-10.0.19045-SP0"
}
```

## 🧪 빌드 검증

### 자동 테스트 실행
```powershell
# COM 캐시 경고 해결 확인
python test_com_cache_fix.py

# 아이콘 통합 확인
python test_icon_integration.py

# 빌드된 EXE 파일 기본 동작 확인
dist\oa.exe --version
dist\oa.exe info
```

### 수동 검증 체크리스트
- [ ] ✅ EXE 파일이 정상적으로 생성됨
- [ ] 🎨 아이콘이 Windows 탐색기에서 표시됨
- [ ] 🔇 첫 실행 시 COM 경고 메시지 없음
- [ ] 📊 `oa.exe --version` 명령어 정상 동작
- [ ] 📋 `oa.exe excel list` 명령어 정상 동작
- [ ] 🏢 `oa.exe hwp list` 명령어 정상 동작 (HWP 설치 시)

## 🔧 문제 해결

### 자주 발생하는 문제들

#### 1. 빌드 실패: "PyInstaller not found"
```powershell
# 해결방법: PyInstaller 설치
pip install pyinstaller

# 가상환경 확인
.venv\Scripts\activate
where python
```

#### 2. 빌드 실패: "아이콘 파일을 찾을 수 없음"
```powershell
# 해결방법: 아이콘 자동 생성
python create_icon.py

# 아이콘 파일 확인
ls pyhub_office_automation\assets\icons\logo.ico
```

#### 3. EXE 실행 시 "DLL 로드 실패"
**원인**: Visual C++ 재배포 패키지 누락
**해결방법**:
```powershell
# Microsoft Visual C++ 재배포 패키지 설치 필요
# https://aka.ms/vs/17/release/vc_redist.x64.exe
```

#### 4. Excel 자동화 오류
**원인**: COM 등록 문제
**해결방법**:
```powershell
# Excel이 설치되어 있는지 확인
# 관리자 권한으로 Excel 한 번 실행
# 사용자 계정 컨트롤(UAC) 설정 확인
```

#### 5. 한글 경로 문제 (macOS 개발 시)
**원인**: macOS NFD 자소 분리
**해결방법**: 프로젝트에 내장된 `normalize_path()` 함수가 자동 처리

### 빌드 로그 확인
```powershell
# 빌드 상세 로그 확인
.\build_windows.ps1 -BuildType onefile -Verbose

# PyInstaller 로그 분석
cat build\oa\warn-oa.txt
```

### 고급 디버깅
```powershell
# PyInstaller 직접 실행 (디버그 모드)
pyinstaller --onefile --debug all oa.spec

# 임포트 오류 분석
python -c "import pyhub_office_automation; print('OK')"

# COM 관련 문제 진단
python -c "import win32com.client; print('COM OK')"
```

## ⚙️ 고급 설정

### PyInstaller Spec 파일 커스터마이징

#### 기본 Spec 파일 (`oa.spec`)
```python
# -*- mode: python ; coding: utf-8 -*-

import os
from pathlib import Path

# 프로젝트 루트 경로
project_root = Path.cwd()

a = Analysis(
    ['pyhub_office_automation/cli/main.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('pyhub_office_automation/assets', 'assets'),
    ],
    hiddenimports=[
        'win32com.gen_py',
        'win32com.client.gencache',
        'win32com.shell.shell',
        'pywintypes',
        'pythoncom',
        'xlwings.pro',
        'pyhwpx',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=['runtime_hook_win32com.py'],
    excludes=[
        'matplotlib', 'scipy', 'sklearn', 'tkinter',
        'IPython', 'jupyter', 'notebook',
    ],
    noarchive=False,
    optimize=0,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='oa',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='pyhub_office_automation/assets/icons/logo.ico'
)
```

#### Spec 파일 수정 옵션

**콘솔 창 숨기기** (GUI 모드):
```python
exe = EXE(
    # ...
    console=False,  # 콘솔 창 숨김
    # ...
)
```

**UPX 압축 비활성화** (호환성 우선):
```python
exe = EXE(
    # ...
    upx=False,  # UPX 압축 비활성화
    # ...
)
```

**추가 라이브러리 포함**:
```python
hiddenimports=[
    # 기본 항목들...
    'your_custom_module',
    'third_party_library',
]
```

### 빌드 최적화

#### 파일 크기 최적화
```python
# excludes에 불필요한 모듈 추가
excludes=[
    'matplotlib', 'scipy', 'sklearn', 'tkinter',
    'IPython', 'jupyter', 'notebook',
    'test', 'tests', 'unittest',
    'doctest', 'pdb', 'profile',
]
```

#### 빌드 속도 향상
```powershell
# 병렬 빌드 (멀티코어 활용)
$env:PYINSTALLER_COMPILE_BOOTLOADER = "1"
pyinstaller --onefile oa.spec --log-level=INFO
```

### CI/CD 통합

#### GitHub Actions 예제
```yaml
name: Build Windows EXE
on: [push, pull_request]

jobs:
  build:
    runs-on: windows-latest
    steps:
      - uses: actions/checkout@v4

      - name: Set up Python
        uses: actions/setup-python@v4
        with:
          python-version: '3.13'

      - name: Install dependencies
        run: pip install -r requirements.txt

      - name: Build EXE
        run: |
          .\build_windows.ps1 -BuildType onefile -CiMode -GenerateMetadata

      - name: Upload artifact
        uses: actions/upload-artifact@v3
        with:
          name: oa-windows-exe
          path: dist/oa.exe
```

## 📚 추가 리소스

### 공식 문서
- [PyInstaller 공식 문서](https://pyinstaller.readthedocs.io/)
- [xlwings 문서](https://docs.xlwings.org/)
- [pyhwpx 문서](https://github.com/cometcomputer/pyhwpx)

### 프로젝트 관련 문서
- [`CLAUDE.md`](../CLAUDE.md) - 프로젝트 전체 가이드
- [`specs/xlwings.md`](../specs/xlwings.md) - xlwings 사용법
- [`docs/차트-명령어-가이드.md`](./차트-명령어-가이드.md) - 차트 기능 상세 가이드

### 문제 신고
- [GitHub Issues](https://github.com/pyhub-apps/pyhub-office-automation/issues)
- [GitHub Discussions](https://github.com/pyhub-apps/pyhub-office-automation/discussions)

---

## 🎉 성공적인 빌드를 위한 체크리스트

### 빌드 전 확인사항
- [ ] ✅ Python 3.13+ 설치됨
- [ ] 📁 프로젝트 루트 디렉터리에서 실행
- [ ] 🔄 가상환경 활성화됨
- [ ] 📦 모든 의존성 설치됨
- [ ] 🎨 아이콘 파일 존재함

### 빌드 실행
- [ ] 💻 PowerShell에서 빌드 스크립트 실행
- [ ] ⏱️ 빌드 완료까지 대기 (5-15분)
- [ ] 📊 빌드 로그에서 오류 확인

### 빌드 후 검증
- [ ] 📁 `dist/oa.exe` 파일 생성 확인
- [ ] 🎨 아이콘 표시 확인
- [ ] 🚀 기본 명령어 동작 확인
- [ ] 📋 Excel/HWP 기능 테스트

**고품질 Windows EXE 배포를 위한 모든 준비가 완료되었습니다!** 🚀