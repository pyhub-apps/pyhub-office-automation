#!/bin/bash
# Local validation script for pyhub-office-automation
# Run this before committing to catch lint errors early

# 기본값 설정
FIX=false
QUICK=false
VERBOSE=false

# 도움말 함수
show_help() {
    cat << EOF
macOS/Linux 코드 검증 스크립트

사용법: $0 [옵션]

옵션:
    --fix              자동으로 포맷팅 및 임포트 정렬 문제 수정
    --quick            중요한 검사만 실행 (빠른 피드백)
    --verbose          상세한 출력 표시
    --help             이 도움말 표시

예제:
    $0                 # 모든 검증 실행
    $0 --fix           # 문제를 자동으로 수정하면서 검증
    $0 --quick         # 빠른 검증 (중요한 것만)
    $0 --verbose       # 상세한 출력과 함께 검증
EOF
}

# 명령줄 인수 파싱
while [[ $# -gt 0 ]]; do
    case $1 in
        --fix)
            FIX=true
            shift
            ;;
        --quick)
            QUICK=true
            shift
            ;;
        --verbose)
            VERBOSE=true
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

echo "🔍 Running local validation checks..."
echo "Fix: $FIX"
echo "Quick: $QUICK"
echo "Verbose: $VERBOSE"
echo "=================================="

# Check if in correct directory
if [ ! -f "pyproject.toml" ]; then
    echo "❌ Error: Run this script from the project root directory"
    exit 1
fi

# Initialize status
all_passed=true
checks=()

# 색상 출력 함수
print_green() {
    echo -e "\033[32m$1\033[0m"
}

print_red() {
    echo -e "\033[31m$1\033[0m"
}

print_yellow() {
    echo -e "\033[33m$1\033[0m"
}

print_cyan() {
    echo -e "\033[36m$1\033[0m"
}

# 1. Black 포맷팅 체크
echo "📝 1. Checking code formatting with Black..."

if [ "$FIX" = true ]; then
    echo "   Auto-fixing formatting issues..."
    black_output=$(black pyhub_office_automation/ 2>&1)
    black_exit=$?
else
    black_output=$(black --check pyhub_office_automation/ 2>&1)
    black_exit=$?
fi

if [ $black_exit -eq 0 ]; then
    print_green "✅ Black formatting: PASSED"
    checks+=("Black:PASSED:")
else
    print_red "❌ Black formatting: FAILED"
    if [ "$FIX" = false ]; then
        print_yellow "    Run: black pyhub_office_automation/"
    fi
    all_passed=false
    checks+=("Black:FAILED:$black_output")
fi

if [ "$VERBOSE" = true ] && [ -n "$black_output" ]; then
    echo "   Output: $black_output"
fi

echo ""

# 2. isort 임포트 정렬 체크
echo "📦 2. Checking import sorting with isort..."

if [ "$FIX" = true ]; then
    echo "   Auto-fixing import sorting..."
    isort_output=$(isort pyhub_office_automation/ 2>&1)
    isort_exit=$?
else
    isort_output=$(isort --check-only pyhub_office_automation/ 2>&1)
    isort_exit=$?
fi

if [ $isort_exit -eq 0 ]; then
    print_green "✅ Import sorting: PASSED"
    checks+=("isort:PASSED:")
else
    print_red "❌ Import sorting: FAILED"
    if [ "$FIX" = false ]; then
        print_yellow "    Run: isort pyhub_office_automation/"
    fi
    all_passed=false
    checks+=("isort:FAILED:$isort_output")
fi

if [ "$VERBOSE" = true ] && [ -n "$isort_output" ]; then
    echo "   Output: $isort_output"
fi

echo ""

# 3. Flake8 린팅 (중요한 오류들)
echo "🔧 3. Running flake8 linting (critical)..."
flake8_output=$(flake8 pyhub_office_automation/ --count --select=E9,F63,F7,F82 --show-source --statistics 2>&1)
flake8_exit=$?

if [ $flake8_exit -eq 0 ]; then
    print_green "✅ Flake8 (critical): PASSED"
    checks+=("Flake8 Critical:PASSED:")
else
    print_red "❌ Flake8 (critical): FAILED"
    all_passed=false
    checks+=("Flake8 Critical:FAILED:$flake8_output")
fi

if [ "$VERBOSE" = true ] && [ -n "$flake8_output" ]; then
    echo "   Output: $flake8_output"
fi

# 4. 추가 Flake8 체크 (Quick 모드가 아닐 때만)
if [ "$QUICK" = false ]; then
    echo ""
    echo "📊 4. Running additional flake8 checks..."
    flake8_additional_output=$(flake8 pyhub_office_automation/ --count --exit-zero --max-complexity=10 --max-line-length=127 --statistics 2>&1)
    print_cyan "ℹ️  Additional checks completed (warnings only)"
    checks+=("Flake8 Additional:INFO:$flake8_additional_output")

    if [ "$VERBOSE" = true ] && [ -n "$flake8_additional_output" ]; then
        echo "   Output: $flake8_additional_output"
    fi
fi

echo ""
echo "=================================="

# 결과 요약
if [ "$all_passed" = true ]; then
    print_green "🎉 All validation checks PASSED!"
    print_green "✅ Ready to commit"

    if [ "$FIX" = true ]; then
        print_cyan "🔧 Auto-fix was applied where possible"
    fi
else
    print_red "💥 Some validation checks FAILED!"
    print_red "❌ Fix issues before committing"

    # 실패한 체크들 나열
    echo ""
    print_yellow "Failed checks:"
    for check in "${checks[@]}"; do
        IFS=':' read -r tool status message <<< "$check"
        if [ "$status" = "FAILED" ]; then
            print_red "  - $tool"
        fi
    done

    if [ "$FIX" = false ]; then
        echo ""
        print_cyan "💡 Tip: Use --fix flag to automatically fix formatting issues:"
        print_cyan "    $0 --fix"
    fi
fi

# 상세 결과 표시 (Verbose 모드)
if [ "$VERBOSE" = true ]; then
    echo ""
    print_cyan "📋 Detailed Results:"
    printf "%-20s %-10s %s\n" "Tool" "Status" "Details"
    echo "=================================================="
    for check in "${checks[@]}"; do
        IFS=':' read -r tool status message <<< "$check"
        if [ "$status" = "PASSED" ]; then
            printf "%-20s \033[32m%-10s\033[0m %s\n" "$tool" "$status" ""
        elif [ "$status" = "FAILED" ]; then
            printf "%-20s \033[31m%-10s\033[0m %s\n" "$tool" "$status" "${message:0:50}..."
        else
            printf "%-20s \033[36m%-10s\033[0m %s\n" "$tool" "$status" ""
        fi
    done
fi

# 성능 정보
echo ""
print_cyan "⚡ Performance tips:"
echo "  - Use --quick for faster feedback during development"
echo "  - Use --fix to automatically resolve formatting issues"
echo "  - Use --verbose for detailed diagnostic information"

exit $([ "$all_passed" = true ] && echo 0 || echo 1)