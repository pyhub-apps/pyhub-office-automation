#!/usr/bin/env pwsh
<#
.SYNOPSIS
    Windows용 PyInstaller 빌드 스크립트
.DESCRIPTION
    pyhub-office-automation을 Windows 실행 파일로 빌드합니다.
.PARAMETER BuildType
    빌드 타입 (onefile 또는 onedir, 기본값: onedir)
.PARAMETER CiMode
    CI 모드 활성화 (자동 진행, 사용자 입력 없음)
.PARAMETER Clean
    빌드 전 기존 파일 정리 (기본값: $true)
.PARAMETER Test
    빌드 후 테스트 실행 (기본값: $true)
.PARAMETER UseSpec
    기존 oa.spec 파일 사용 (기본값: $false)
.PARAMETER GenerateMetadata
    빌드 메타데이터 JSON 파일 생성 (기본값: $false)
.EXAMPLE
    .\build_windows.ps1
    기본 설정으로 빌드 (onedir 모드)
.EXAMPLE
    .\build_windows.ps1 -BuildType onefile -CiMode
    CI 환경에서 onefile 모드로 빌드
.EXAMPLE
    .\build_windows.ps1 -BuildType onedir -Clean:$false
    기존 파일을 정리하지 않고 onedir 모드로 빌드
.EXAMPLE
    .\build_windows.ps1 -BuildType onefile -GenerateMetadata
    빌드 메타데이터와 함께 onefile 모드로 빌드
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [ValidateSet("onefile", "onedir")]
    [string]$BuildType = "onedir",

    [Parameter(Mandatory = $false)]
    [switch]$CiMode,

    [Parameter(Mandatory = $false)]
    [bool]$Clean = $true,

    [Parameter(Mandatory = $false)]
    [bool]$Test = $true,

    [Parameter(Mandatory = $false)]
    [switch]$UseSpec,

    [Parameter(Mandatory = $false)]
    [switch]$GenerateMetadata
)

# PowerShell 스트릭트 모드 활성화
Set-StrictMode -Version Latest

# UTF-8 인코딩 설정
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::InputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

Write-Host "=========================================="
Write-Host "pyhub-office-automation Windows Build"
Write-Host "Build Type: $BuildType"
Write-Host "CI Mode: $CiMode"
Write-Host "Clean: $Clean"
Write-Host "Test: $Test"
Write-Host "Use Spec: $UseSpec"
Write-Host "Generate Metadata: $GenerateMetadata"
Write-Host "=========================================="

# 에러 발생 시 스크립트 중단
$ErrorActionPreference = "Stop"

try {
    # 기존 빌드 파일 정리
    if ($Clean) {
        Write-Host "🧹 Cleaning previous build files..."
        $itemsToRemove = @("build", "dist")

        # UseSpec이 false인 경우에만 oa.spec 삭제
        if (-not $UseSpec) {
            $itemsToRemove += "oa.spec"
        }

        foreach ($item in $itemsToRemove) {
            if (Test-Path $item) {
                Remove-Item -Recurse -Force $item -ErrorAction SilentlyContinue
                Write-Host "   Removed: $item"
            }
        }

        if ($UseSpec -and (Test-Path "oa.spec")) {
            Write-Host "   Preserved: oa.spec (UseSpec enabled)"
        }

        Write-Host "   Cleanup completed"
    }

    # Python 및 PyInstaller 확인
    Write-Host "🔍 Checking dependencies..."
    try {
        $pythonVersion = python --version 2>&1
        Write-Host "   Python: $pythonVersion"
    }
    catch {
        throw "Python이 설치되어 있지 않거나 PATH에 없습니다."
    }

    try {
        $pyinstallerVersion = pyinstaller --version 2>&1
        Write-Host "   PyInstaller: $pyinstallerVersion"
    }
    catch {
        throw "PyInstaller가 설치되어 있지 않습니다. 'pip install pyinstaller' 또는 'uv add pyinstaller'를 실행하세요."
    }

    # 프로젝트 정보 확인 및 빌드 버전 고정
    Write-Host "📦 Getting project information..."
    try {
        $version = python -c "import sys; sys.path.insert(0, 'pyhub_office_automation'); from version import get_version; print(get_version())"
        Write-Host "   Version: $version"

        # 빌드 시 버전 고정을 위한 __version__.py 파일 생성 (GitHub Actions에서 이미 생성되었으면 스킵)
        $versionPyPath = "pyhub_office_automation\__version__.py"
        if (-not (Test-Path $versionPyPath)) {
            Write-Host "🔧 Creating fixed version file for build..."
            $versionPyContent = "__version__ = '$version'"
            $versionPyContent | Out-File -FilePath $versionPyPath -Encoding UTF8
            Write-Host "   Fixed version file created: $versionPyPath"
        } else {
            Write-Host "🔧 Fixed version file already exists, skipping creation"
            $existingVersion = Get-Content $versionPyPath | Select-String "__version__" | ForEach-Object { $_.Line -replace '.*"([^"]*)".*', '$1' }
            Write-Host "   Using existing version: $existingVersion"
        }
    }
    catch {
        Write-Warning "버전 정보를 가져올 수 없습니다: $($_.Exception.Message)"
        $version = "unknown"
    }

    # win32com 캐시 사전 준비 (COM 재구축 경고 방지)
    Write-Host "🔧 Preparing win32com cache..."
    try {
        python -c "
import sys
import tempfile
import warnings

# PyInstaller 빌드 전 win32com 캐시 사전 생성
try:
    with warnings.catch_warnings():
        warnings.simplefilter('ignore')
        import win32com.client

        # Excel 애플리케이션 COM 인터페이스 사전 생성
        win32com.client.gencache.is_readonly = False

        # Excel COM 클래스 사전 로드 (경고 없이)
        try:
            xl = win32com.client.Dispatch('Excel.Application')
            xl.Quit()
            print('   Excel COM interface pre-generated successfully')
        except:
            print('   Excel COM interface pre-generation skipped (Excel not available)')

except ImportError:
    print('   win32com not available, skipping cache preparation')
except Exception as e:
    print(f'   Warning: win32com cache preparation failed: {e}')
"
    }
    catch {
        Write-Warning "win32com 캐시 준비 실패: $($_.Exception.Message)"
    }

    # 아이콘 파일 확인
    Write-Host "🎨 Checking icon files..."
    $iconPath = "pyhub_office_automation\assets\icons\logo.ico"
    if (Test-Path $iconPath) {
        $iconSize = [math]::Round((Get-Item $iconPath).Length / 1024, 1)
        Write-Host "   Icon file found: $iconPath ($iconSize KB)"
    }
    else {
        Write-Warning "Icon file not found: $iconPath"
        Write-Host "   Creating icon files..."
        try {
            python create_icon.py
            if ($LASTEXITCODE -eq 0) {
                Write-Host "   ✅ Icon files created successfully"
            }
            else {
                Write-Warning "Icon creation failed, continuing without icon"
            }
        }
        catch {
            Write-Warning "Icon creation script failed: $($_.Exception.Message)"
        }
    }

    # PyInstaller 빌드 인수 준비
    Write-Host "🔨 Building with PyInstaller..."

    if ($UseSpec -and (Test-Path "oa.spec")) {
        Write-Host "   Using existing oa.spec file..."
        $buildArgs = @("oa.spec")

        # spec 파일을 사용할 때는 BuildType에 따라 수정이 필요할 수 있음
        if ($BuildType -eq "onefile") {
            Write-Host "   Note: BuildType 'onefile' specified, but using spec file. Check spec file configuration."
        }

        # 아이콘 파일 경로 확인
        if (Test-Path $iconPath) {
            Write-Host "   Icon will be included from spec file: $iconPath"
        }
    }
    else {
        Write-Host "   Building with command-line arguments..."
        $excludeModules = @(
            "matplotlib",
            "scipy",
            "sklearn",
            "tkinter",
            "IPython",
            "jupyter",
            "PIL.ImageQt"
        )

        $buildArgs = @(
            "--$BuildType",
            "--name", "oa",
            "--console",
            "--noconfirm",
            "--clean"
        )

        # 아이콘 추가 (존재하는 경우)
        if (Test-Path $iconPath) {
            $buildArgs += @("--icon", $iconPath)
            Write-Host "   Adding icon: $iconPath"
        }

        # 제외할 모듈 추가
        foreach ($module in $excludeModules) {
            $buildArgs += @("--exclude-module", $module)
        }


        # 리소스 파일 추가
        $buildArgs += @("--add-data", "pyhub_office_automation\resources;pyhub_office_automation\resources")
        $buildArgs += @("--add-data", "README.md;.")
        Write-Host "   Adding resources: pyhub_office_automation\resources and README.md"

        # 메인 스크립트 경로
        $buildArgs += "pyhub_office_automation\cli\main.py"
    }

    Write-Host "   Build arguments: $($buildArgs -join ' ')"
    Write-Host "   Starting build process..."

    # PyInstaller 실행
    & pyinstaller @buildArgs

    if ($LASTEXITCODE -ne 0) {
        throw "PyInstaller 빌드가 실패했습니다. (Exit code: $LASTEXITCODE)"
    }

    Write-Host "✅ Build completed successfully!"

    # 빌드 결과 확인
    if ($BuildType -eq "onefile") {
        $exePath = "dist\oa.exe"
    }
    else {
        $exePath = "dist\oa\oa.exe"
    }

    if (-not (Test-Path $exePath)) {
        throw "빌드된 실행파일을 찾을 수 없습니다: $exePath"
    }

    $fileSize = [math]::Round((Get-Item $exePath).Length / 1MB, 2)
    Write-Host "📁 Build output:"
    Write-Host "   Location: $exePath"
    Write-Host "   Size: ${fileSize} MB"

    # version.txt 파일 생성
    Write-Host "📝 Creating version.txt file..."
    try {
        # KST 시간 계산 (UTC+9)
        $utcTime = Get-Date -AsUTC
        $kstTime = $utcTime.AddHours(9)
        $buildDate = $kstTime.ToString("yyyy-MM-dd HH:mm:ss") + " KST"
        $gitTag = try { git describe --tags --exact-match HEAD 2>&1 } catch { "local-build" }

        $versionContent = @"
pyhub-office-automation
버전: $version
빌드 시간: $buildDate
Git 태그: $gitTag
빌드 타입: $BuildType
"@

        if ($BuildType -eq "onefile") {
            $versionPath = "dist\version.txt"
        } else {
            $versionPath = "dist\oa\version.txt"
        }

        $versionContent | Out-File -FilePath $versionPath -Encoding UTF8
        Write-Host "   Version file created: $versionPath"
    }
    catch {
        Write-Warning "version.txt 생성 실패: $($_.Exception.Message)"
    }

    # 빌드 메타데이터 생성
    if ($GenerateMetadata) {
        Write-Host "📊 Generating build metadata..."
        try {
            $hash = Get-FileHash $exePath -Algorithm SHA256
            $buildMetadata = [ordered]@{
                BuildInfo = [ordered]@{
                    Version = $version
                    BuildTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss UTC" -AsUTC
                    BuildType = $BuildType
                    UseSpec = $UseSpec.IsPresent
                    CiMode = $CiMode.IsPresent
                }
                FileInfo = [ordered]@{
                    Location = $exePath
                    SizeMB = $fileSize
                    SHA256 = $hash.Hash
                }
                Environment = [ordered]@{
                    PowerShellVersion = $PSVersionTable.PSVersion.ToString()
                    OSVersion = [System.Environment]::OSVersion.ToString()
                    MachineName = [System.Environment]::MachineName
                }
            }

            $metadataJson = $buildMetadata | ConvertTo-Json -Depth 3
            $metadataPath = "build-metadata.json"
            $metadataJson | Out-File -FilePath $metadataPath -Encoding UTF8
            Write-Host "   Metadata saved to: $metadataPath"
            Write-Host "   SHA256: $($hash.Hash.Substring(0, 16))..."
        }
        catch {
            Write-Warning "메타데이터 생성 실패: $($_.Exception.Message)"
        }
    }

    # 테스트 실행
    if ($Test) {
        Write-Host "🧪 Testing build..."

        # 버전 테스트
        Write-Host "   Testing --version option..."
        try {
            $output = & $exePath --version 2>&1
            Write-Host "   Version output: $output"
        }
        catch {
            Write-Warning "버전 테스트 실패: $($_.Exception.Message)"
        }

        # 기본 명령어 테스트
        Write-Host "   Testing excel list command..."
        try {
            $output = & $exePath excel list --format text 2>&1
            Write-Host "   Excel list test completed"
        }
        catch {
            Write-Warning "Excel list 테스트 실패 (예상됨 - Excel이 설치되지 않은 경우): $($_.Exception.Message)"
        }

        # info 명령어 테스트
        Write-Host "   Testing info command..."
        try {
            $output = & $exePath info --format json 2>&1
            Write-Host "   Info test completed"
        }
        catch {
            Write-Warning "Info 테스트 실패 (예상됨 - Office 프로그램이 설치되지 않은 경우): $($_.Exception.Message)"
        }

        # llm-guide 리소스 로딩 테스트
        Write-Host "   Testing llm-guide resource loading..."
        try {
            $output = & $exePath llm-guide --format text 2>&1
            if ($output -like "*초심자 대응 가이드*") {
                Write-Host "   ✅ llm-guide resource loading successful"
            } else {
                Write-Warning "llm-guide 리소스 로딩 실패 - fallback 메시지만 출력됨"
            }
        }
        catch {
            Write-Warning "llm-guide 테스트 실패: $($_.Exception.Message)"
        }

        Write-Host "✅ Basic tests completed"
    }

    # 성공 메시지
    Write-Host ""
    Write-Host "🎉 =========================================="
    Write-Host "🎉 Build completed successfully!"
    Write-Host "🎉 =========================================="
    Write-Host "📁 Executable location: $exePath"
    Write-Host "📊 File size: ${fileSize} MB"
    Write-Host "📋 Version: $version"
    Write-Host ""

    if (-not $CiMode) {
        Write-Host "사용법:"
        Write-Host "  $exePath --version"
        Write-Host "  $exePath info"
        Write-Host "  $exePath excel list"
        Write-Host "  $exePath hwp list"
        Write-Host ""
        Read-Host "Press Enter to continue..."
    }
}
catch {
    Write-Error "❌ 빌드 실패: $($_.Exception.Message)"
    if (-not $CiMode) {
        Read-Host "Press Enter to exit..."
    }
    exit 1
}