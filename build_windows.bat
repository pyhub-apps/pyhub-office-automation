@echo off
REM Windows용 PyInstaller 빌드 스크립트
REM pyhub-office-automation을 Windows 실행 파일로 빌드

echo ==========================================
echo pyhub-office-automation Windows Build
echo ==========================================

REM 기존 빌드 파일 정리
echo Cleaning previous build files...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist oa.spec del oa.spec

REM PyInstaller로 빌드
echo Building with PyInstaller...
pyinstaller ^
  --onedir ^
  --name oa ^
  --console ^
  --exclude-module matplotlib ^
  --exclude-module scipy ^
  --exclude-module sklearn ^
  --exclude-module tkinter ^
  --exclude-module IPython ^
  --exclude-module jupyter ^
  --noconfirm ^
  pyhub_office_automation\cli\main.py

REM 빌드 결과 확인
echo.
echo ==========================================
echo Build completed!
echo ==========================================
echo.
echo Testing build...
dist\oa\oa.exe excel list --format text

echo.
echo Build location: dist\oa\oa.exe
echo.
pause