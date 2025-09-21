@echo off
REM Local validation script for pyhub-office-automation
REM Run this before committing to catch lint errors early

echo 🔍 Running local validation checks...
echo ==================================

REM Check if in correct directory
if not exist "pyproject.toml" (
    echo ❌ Error: Run this script from the project root directory
    exit /b 1
)

REM Initialize status
set "all_passed=true"

echo 📝 1. Checking code formatting with Black...
black --check pyhub_office_automation\
if %errorlevel% equ 0 (
    echo ✅ Black formatting: PASSED
) else (
    echo ❌ Black formatting: FAILED
    echo    Run: black pyhub_office_automation\
    set "all_passed=false"
)

echo.
echo 📦 2. Checking import sorting with isort...
isort --check-only pyhub_office_automation\
if %errorlevel% equ 0 (
    echo ✅ Import sorting: PASSED
) else (
    echo ❌ Import sorting: FAILED
    echo    Run: isort pyhub_office_automation\
    set "all_passed=false"
)

echo.
echo 🔧 3. Running flake8 linting...
flake8 pyhub_office_automation\ --count --select=E9,F63,F7,F82 --show-source --statistics
if %errorlevel% equ 0 (
    echo ✅ Flake8 ^(critical^): PASSED
) else (
    echo ❌ Flake8 ^(critical^): FAILED
    set "all_passed=false"
)

REM Additional flake8 check (non-critical)
echo.
echo 📊 4. Running additional flake8 checks...
flake8 pyhub_office_automation\ --count --exit-zero --max-complexity=10 --max-line-length=127 --statistics
echo ℹ️  Additional checks completed ^(warnings only^)

echo.
echo ==================================
if "%all_passed%"=="true" (
    echo 🎉 All validation checks PASSED!
    echo ✅ Ready to commit
    exit /b 0
) else (
    echo 💥 Some validation checks FAILED!
    echo ❌ Fix issues before committing
    exit /b 1
)