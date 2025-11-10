# Build script for DemoApp
param(
    [string]$Configuration = "Debug",
    [string]$Platform = "x64"
)

$ErrorActionPreference = "Stop"

# Change to project directory
Set-Location "C:\Program Files\Solomon Technology Corp\SolVision6\Version_6.1.4\SampleCode\C_Sharp\DemoApp"

# Find Visual Studio and MSBuild
$vsWhere = "C:\Program Files (x86)\Microsoft Visual Studio\Installer\vswhere.exe"
if (!(Test-Path $vsWhere)) {
    Write-Error "vswhere.exe not found. Is Visual Studio installed?"
    exit 1
}

$vsPath = & $vsWhere -latest -property installationPath
if (!$vsPath) {
    Write-Error "Visual Studio installation not found"
    exit 1
}

$msbuild = Join-Path $vsPath "MSBuild\Current\Bin\MSBuild.exe"
if (!(Test-Path $msbuild)) {
    $msbuild = Join-Path $vsPath "MSBuild\15.0\Bin\MSBuild.exe"
}

if (!(Test-Path $msbuild)) {
    Write-Error "MSBuild.exe not found"
    exit 1
}

# Build the solution
Write-Host "Building DemoApp.sln ($Configuration|$Platform)..."
& $msbuild "DemoApp.sln" /p:Configuration=$Configuration /p:Platform=$Platform /v:minimal /nologo

if ($LASTEXITCODE -ne 0) {
    Write-Error "Build failed with exit code $LASTEXITCODE"
    exit $LASTEXITCODE
}

Write-Host "Build completed successfully!" -ForegroundColor Green
