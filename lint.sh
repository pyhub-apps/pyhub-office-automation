#!/bin/bash
# Local validation script for pyhub-office-automation
# Run this before committing to catch lint errors early

echo "🔍 Running local validation checks..."
echo "=================================="

# Check if in correct directory
if [ ! -f "pyproject.toml" ]; then
    echo "❌ Error: Run this script from the project root directory"
    exit 1
fi

# Initialize status
all_passed=true

echo "📝 1. Checking code formatting with Black..."
if black --check pyhub_office_automation/; then
    echo "✅ Black formatting: PASSED"
else
    echo "❌ Black formatting: FAILED"
    echo "   Run: black pyhub_office_automation/"
    all_passed=false
fi

echo ""
echo "📦 2. Checking import sorting with isort..."
if isort --check-only pyhub_office_automation/; then
    echo "✅ Import sorting: PASSED"
else
    echo "❌ Import sorting: FAILED"
    echo "   Run: isort pyhub_office_automation/"
    all_passed=false
fi

echo ""
echo "🔧 3. Running flake8 linting..."
if flake8 pyhub_office_automation/ --count --select=E9,F63,F7,F82 --show-source --statistics; then
    echo "✅ Flake8 (critical): PASSED"
else
    echo "❌ Flake8 (critical): FAILED"
    all_passed=false
fi

# Additional flake8 check (non-critical)
echo ""
echo "📊 4. Running additional flake8 checks..."
flake8 pyhub_office_automation/ --count --exit-zero --max-complexity=10 --max-line-length=127 --statistics
echo "ℹ️  Additional checks completed (warnings only)"

echo ""
echo "=================================="
if [ "$all_passed" = true ]; then
    echo "🎉 All validation checks PASSED!"
    echo "✅ Ready to commit"
    exit 0
else
    echo "💥 Some validation checks FAILED!"
    echo "❌ Fix issues before committing"
    exit 1
fi