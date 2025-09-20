#!/bin/bash
# macOS용 PyInstaller 빌드 스크립트
# pyhub-office-automation을 macOS 실행 파일로 빌드

echo "=========================================="
echo "pyhub-office-automation macOS Build"
echo "=========================================="

# 기존 빌드 파일 정리
echo "Cleaning previous build files..."
rm -rf build dist oa.spec

# PyInstaller로 빌드
echo "Building with PyInstaller..."
pyinstaller \
  --onedir \
  --name oa \
  --console \
  --exclude-module matplotlib \
  --exclude-module scipy \
  --exclude-module sklearn \
  --exclude-module tkinter \
  --exclude-module IPython \
  --exclude-module jupyter \
  --noconfirm \
  pyhub_office_automation/cli/main.py

# 빌드 결과 확인
echo ""
echo "=========================================="
echo "Build completed!"
echo "=========================================="
echo ""
echo "Testing build..."
./dist/oa/oa excel list --format text

echo ""
echo "Build location: dist/oa/oa"
echo ""

# 실행 권한 확인
if [ -x "./dist/oa/oa" ]; then
    echo "✅ Executable is ready to use"
else
    echo "❌ Making executable..."
    chmod +x ./dist/oa/oa
fi

echo ""
echo "To run: ./dist/oa/oa --help"