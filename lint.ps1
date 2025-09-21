#!/usr/bin/env pwsh
<#
.SYNOPSIS
    Local validation script for pyhub-office-automation
.DESCRIPTION
    Run this before committing to catch lint errors early.
    Checks code formatting, import sorting, and linting.
.PARAMETER Fix
    Automatically fix formatting and import issues
.PARAMETER Quick
    Run only critical checks (skip additional flake8)
.PARAMETER Verbose
    Show detailed output from all tools
.EXAMPLE
    .\lint.ps1
    Run all validation checks
.EXAMPLE
    .\lint.ps1 -Fix
    Run checks and automatically fix issues
.EXAMPLE
    .\lint.ps1 -Quick
    Run only critical checks for faster feedback
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [switch]$Fix,

    [Parameter(Mandatory = $false)]
    [switch]$Quick,

    [Parameter(Mandatory = $false)]
    [switch]$Verbose
)

# PowerShell 스트릭트 모드 활성화
Set-StrictMode -Version Latest

# UTF-8 인코딩 설정
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::InputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

Write-Host "🔍 Running local validation checks..."
Write-Host "=================================="

# 에러 발생 시 스크립트 중단하지 않음 (개별 체크별로 처리)
$ErrorActionPreference = "Continue"

try {
    # 프로젝트 루트 디렉터리 확인
    if (-not (Test-Path "pyproject.toml")) {
        Write-Error "❌ Error: Run this script from the project root directory"
        exit 1
    }

    # 상태 초기화
    $allPassed = $true
    $checks = @()

    # 1. Black 포맷팅 체크
    Write-Host "📝 1. Checking code formatting with Black..."

    if ($Fix) {
        Write-Host "   Auto-fixing formatting issues..."
        $blackResult = & black pyhub_office_automation\ 2>&1
        $blackExitCode = $LASTEXITCODE
    } else {
        $blackResult = & black --check pyhub_office_automation\ 2>&1
        $blackExitCode = $LASTEXITCODE
    }

    if ($blackExitCode -eq 0) {
        Write-Host "✅ Black formatting: PASSED" -ForegroundColor Green
        $checks += [PSCustomObject]@{ Tool = "Black"; Status = "PASSED"; Message = "" }
    } else {
        Write-Host "❌ Black formatting: FAILED" -ForegroundColor Red
        if (-not $Fix) {
            Write-Host "    Run: black pyhub_office_automation\" -ForegroundColor Yellow
        }
        $allPassed = $false
        $checks += [PSCustomObject]@{ Tool = "Black"; Status = "FAILED"; Message = $blackResult }
    }

    if ($Verbose -and $blackResult) {
        Write-Host "   Output: $blackResult" -ForegroundColor Gray
    }

    Write-Host ""

    # 2. isort 임포트 정렬 체크
    Write-Host "📦 2. Checking import sorting with isort..."

    if ($Fix) {
        Write-Host "   Auto-fixing import sorting..."
        $isortResult = & isort pyhub_office_automation\ 2>&1
        $isortExitCode = $LASTEXITCODE
    } else {
        $isortResult = & isort --check-only pyhub_office_automation\ 2>&1
        $isortExitCode = $LASTEXITCODE
    }

    if ($isortExitCode -eq 0) {
        Write-Host "✅ Import sorting: PASSED" -ForegroundColor Green
        $checks += [PSCustomObject]@{ Tool = "isort"; Status = "PASSED"; Message = "" }
    } else {
        Write-Host "❌ Import sorting: FAILED" -ForegroundColor Red
        if (-not $Fix) {
            Write-Host "    Run: isort pyhub_office_automation\" -ForegroundColor Yellow
        }
        $allPassed = $false
        $checks += [PSCustomObject]@{ Tool = "isort"; Status = "FAILED"; Message = $isortResult }
    }

    if ($Verbose -and $isortResult) {
        Write-Host "   Output: $isortResult" -ForegroundColor Gray
    }

    Write-Host ""

    # 3. Flake8 린팅 (중요한 오류들)
    Write-Host "🔧 3. Running flake8 linting (critical)..."
    $flake8Result = & flake8 pyhub_office_automation\ --count --select=E9,F63,F7,F82 --show-source --statistics 2>&1
    $flake8ExitCode = $LASTEXITCODE

    if ($flake8ExitCode -eq 0) {
        Write-Host "✅ Flake8 (critical): PASSED" -ForegroundColor Green
        $checks += [PSCustomObject]@{ Tool = "Flake8 Critical"; Status = "PASSED"; Message = "" }
    } else {
        Write-Host "❌ Flake8 (critical): FAILED" -ForegroundColor Red
        $allPassed = $false
        $checks += [PSCustomObject]@{ Tool = "Flake8 Critical"; Status = "FAILED"; Message = $flake8Result }
    }

    if ($Verbose -and $flake8Result) {
        Write-Host "   Output: $flake8Result" -ForegroundColor Gray
    }

    # 4. 추가 Flake8 체크 (Quick 모드가 아닐 때만)
    if (-not $Quick) {
        Write-Host ""
        Write-Host "📊 4. Running additional flake8 checks..."
        $flake8AdditionalResult = & flake8 pyhub_office_automation\ --count --exit-zero --max-complexity=10 --max-line-length=127 --statistics 2>&1
        Write-Host "ℹ️  Additional checks completed (warnings only)" -ForegroundColor Cyan
        $checks += [PSCustomObject]@{ Tool = "Flake8 Additional"; Status = "INFO"; Message = $flake8AdditionalResult }

        if ($Verbose -and $flake8AdditionalResult) {
            Write-Host "   Output: $flake8AdditionalResult" -ForegroundColor Gray
        }
    }

    Write-Host ""
    Write-Host "=================================="

    # 결과 요약
    if ($allPassed) {
        Write-Host "🎉 All validation checks PASSED!" -ForegroundColor Green
        Write-Host "✅ Ready to commit" -ForegroundColor Green

        if ($Fix) {
            Write-Host "🔧 Auto-fix was applied where possible" -ForegroundColor Cyan
        }
    } else {
        Write-Host "💥 Some validation checks FAILED!" -ForegroundColor Red
        Write-Host "❌ Fix issues before committing" -ForegroundColor Red

        # 실패한 체크들 나열
        $failedChecks = $checks | Where-Object { $_.Status -eq "FAILED" }
        if ($failedChecks) {
            Write-Host ""
            Write-Host "Failed checks:" -ForegroundColor Yellow
            foreach ($check in $failedChecks) {
                Write-Host "  - $($check.Tool)" -ForegroundColor Red
            }
        }

        if (-not $Fix) {
            Write-Host ""
            Write-Host "💡 Tip: Use -Fix flag to automatically fix formatting issues:" -ForegroundColor Cyan
            Write-Host "    .\lint.ps1 -Fix" -ForegroundColor Cyan
        }
    }

    # 상세 결과 표시 (Verbose 모드)
    if ($Verbose) {
        Write-Host ""
        Write-Host "📋 Detailed Results:" -ForegroundColor Cyan
        $checks | Format-Table -AutoSize
    }

    # 성능 정보
    Write-Host ""
    Write-Host "⚡ Performance tips:" -ForegroundColor Cyan
    Write-Host "  - Use -Quick for faster feedback during development" -ForegroundColor Gray
    Write-Host "  - Use -Fix to automatically resolve formatting issues" -ForegroundColor Gray
    Write-Host "  - Use -Verbose for detailed diagnostic information" -ForegroundColor Gray

    exit $(if ($allPassed) { 0 } else { 1 })
}
catch {
    Write-Error "❌ Unexpected error during validation: $($_.Exception.Message)"
    exit 1
}