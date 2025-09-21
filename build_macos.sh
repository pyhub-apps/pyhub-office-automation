#!/bin/bash
# macOS/Linux용 PyInstaller 빌드 스크립트
# pyhub-office-automation을 macOS/Linux 실행 파일로 빌드

# 기본값 설정
BUILD_TYPE="onedir"
CI_MODE=false
CLEAN=true
TEST=true
USE_SPEC=false
GENERATE_METADATA=false

# 도움말 함수
show_help() {
    cat << EOF
macOS/Linux PyInstaller 빌드 스크립트

사용법: $0 [옵션]

옵션:
    --onefile          단일 실행 파일로 빌드 (기본값: onedir)
    --onedir           폴더 형태로 빌드 (기본값)
    --ci               CI 모드 (자동 진행, 사용자 입력 없음)
    --no-clean         빌드 전 기존 파일을 정리하지 않음
    --no-test          빌드 후 테스트를 실행하지 않음
    --use-spec         기존 oa.spec 파일 사용
    --metadata         빌드 메타데이터 생성
    --help             이 도움말 표시

예제:
    $0                         # 기본 빌드 (onedir)
    $0 --onefile --metadata    # 단일 파일로 빌드하고 메타데이터 생성
    $0 --ci --onefile          # CI 환경에서 단일 파일 빌드
EOF
}

# 명령줄 인수 파싱
while [[ $# -gt 0 ]]; do
    case $1 in
        --onefile)
            BUILD_TYPE="onefile"
            shift
            ;;
        --onedir)
            BUILD_TYPE="onedir"
            shift
            ;;
        --ci)
            CI_MODE=true
            shift
            ;;
        --no-clean)
            CLEAN=false
            shift
            ;;
        --no-test)
            TEST=false
            shift
            ;;
        --use-spec)
            USE_SPEC=true
            shift
            ;;
        --metadata)
            GENERATE_METADATA=true
            shift
            ;;
        --help)
            show_help
            exit 0
            ;;
        *)
            echo "❌ 알 수 없는 옵션: $1"
            show_help
            exit 1
            ;;
    esac
done

echo "=========================================="
echo "pyhub-office-automation macOS/Linux Build"
echo "Build Type: $BUILD_TYPE"
echo "CI Mode: $CI_MODE"
echo "Clean: $CLEAN"
echo "Test: $TEST"
echo "Use Spec: $USE_SPEC"
echo "Generate Metadata: $GENERATE_METADATA"
echo "=========================================="

# 오류 발생 시 스크립트 중단
set -e

# 기존 빌드 파일 정리
if [ "$CLEAN" = true ]; then
    echo "🧹 Cleaning previous build files..."
    rm -rf build dist oa.spec 2>/dev/null || true
    echo "   Cleanup completed"
fi

# Python 및 PyInstaller 확인
echo "🔍 Checking dependencies..."
if ! command -v python3 &> /dev/null; then
    echo "❌ Python3이 설치되어 있지 않습니다."
    exit 1
fi
python_version=$(python3 --version)
echo "   Python: $python_version"

if ! command -v pyinstaller &> /dev/null; then
    echo "❌ PyInstaller가 설치되어 있지 않습니다. 'pip install pyinstaller'를 실행하세요."
    exit 1
fi
pyinstaller_version=$(pyinstaller --version)
echo "   PyInstaller: $pyinstaller_version"

# 프로젝트 정보 확인
echo "📦 Getting project information..."
if [ -f "pyhub_office_automation/__init__.py" ] && command -v python3 &> /dev/null; then
    version=$(python3 -c "import sys; sys.path.insert(0, 'pyhub_office_automation'); from version import get_version; print(get_version())" 2>/dev/null || echo "unknown")
    echo "   Version: $version"
else
    echo "   ⚠️ 버전 정보를 가져올 수 없습니다."
    version="unknown"
fi

# PyInstaller 빌드
echo "🔨 Building with PyInstaller..."

if [ "$USE_SPEC" = true ] && [ -f "oa.spec" ]; then
    echo "   Using existing oa.spec file..."
    pyinstaller oa.spec
    if [ "$BUILD_TYPE" = "onefile" ]; then
        echo "   Note: BuildType 'onefile' specified, but using spec file. Check spec file configuration."
    fi
else
    echo "   Building with command-line arguments..."
    pyinstaller \
      --$BUILD_TYPE \
      --name oa \
      --console \
      --exclude-module matplotlib \
      --exclude-module scipy \
      --exclude-module sklearn \
      --exclude-module tkinter \
      --exclude-module IPython \
      --exclude-module jupyter \
      --exclude-module PIL.ImageQt \
      --noconfirm \
      --clean \
      pyhub_office_automation/cli/main.py
fi

echo "✅ Build completed successfully!"

# 빌드 결과 확인
if [ "$BUILD_TYPE" = "onefile" ]; then
    exe_path="./dist/oa"
else
    exe_path="./dist/oa/oa"
fi

if [ ! -f "$exe_path" ]; then
    echo "❌ 빌드된 실행파일을 찾을 수 없습니다: $exe_path"
    exit 1
fi

# 파일 크기 확인
file_size_bytes=$(stat -f%z "$exe_path" 2>/dev/null || stat -c%s "$exe_path" 2>/dev/null || echo "0")
file_size_mb=$(echo "scale=2; $file_size_bytes / 1024 / 1024" | bc 2>/dev/null || echo "unknown")

echo "📁 Build output:"
echo "   Location: $exe_path"
echo "   Size: ${file_size_mb} MB"

# 빌드 메타데이터 생성
if [ "$GENERATE_METADATA" = true ]; then
    echo "📊 Generating build metadata..."
    if command -v shasum &> /dev/null; then
        sha256_hash=$(shasum -a 256 "$exe_path" | cut -d' ' -f1)
    elif command -v sha256sum &> /dev/null; then
        sha256_hash=$(sha256sum "$exe_path" | cut -d' ' -f1)
    else
        sha256_hash="unavailable"
    fi

    build_time=$(date -u +"%Y-%m-%d %H:%M:%S UTC")
    os_info=$(uname -a)

    cat > build-metadata.json << EOF
{
  "BuildInfo": {
    "Version": "$version",
    "BuildTime": "$build_time",
    "BuildType": "$BUILD_TYPE",
    "UseSpec": $USE_SPEC,
    "CiMode": $CI_MODE
  },
  "FileInfo": {
    "Location": "$exe_path",
    "SizeMB": $file_size_mb,
    "SHA256": "$sha256_hash"
  },
  "Environment": {
    "OS": "$(uname -s)",
    "Architecture": "$(uname -m)",
    "SystemInfo": "$os_info",
    "ShellVersion": "$BASH_VERSION"
  }
}
EOF
    echo "   Metadata saved to: build-metadata.json"
    echo "   SHA256: ${sha256_hash:0:16}..."
fi

# 실행 권한 확인 및 설정
if [ ! -x "$exe_path" ]; then
    echo "🔧 Setting executable permissions..."
    chmod +x "$exe_path"
fi

# 테스트 실행
if [ "$TEST" = true ]; then
    echo "🧪 Testing build..."

    # 버전 테스트
    echo "   Testing --version option..."
    if version_output=$("$exe_path" --version 2>&1); then
        echo "   ✅ Version test passed: $version_output"
    else
        echo "   ⚠️ Version test failed: $version_output"
    fi

    # 기본 명령어 테스트
    echo "   Testing excel list command..."
    if excel_output=$("$exe_path" excel list --format text 2>&1); then
        echo "   ✅ Excel list test completed"
    else
        echo "   ℹ️ Excel list test failed (expected - Excel not available): $excel_output"
    fi

    # info 명령어 테스트
    echo "   Testing info command..."
    if info_output=$("$exe_path" info --format json 2>&1); then
        echo "   ✅ Info test completed"
    else
        echo "   ℹ️ Info test failed (expected - Office not available): $info_output"
    fi

    echo "✅ Basic tests completed"
fi

# 성공 메시지
echo ""
echo "🎉 =========================================="
echo "🎉 Build completed successfully!"
echo "🎉 =========================================="
echo "📁 Executable location: $exe_path"
echo "📊 File size: ${file_size_mb} MB"
echo "📋 Version: $version"
echo ""

if [ "$CI_MODE" = false ]; then
    echo "사용법:"
    echo "  $exe_path --version"
    echo "  $exe_path info"
    echo "  $exe_path excel list"
    echo "  $exe_path hwp list"
    echo ""
    echo "Press Enter to continue..."
    read -r
fi