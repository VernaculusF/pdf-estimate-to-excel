param(
    [Parameter(Mandatory = $true)]
    [string]$Root
)

$ErrorActionPreference = "Stop"

$inputDir = Join-Path $Root "input"
$outputDir = Join-Path $Root "output"
$legacyInputDir = Join-Path $Root "вход(pdf)"
$legacyOutputDir = Join-Path $Root "выход(excel)"
$projectDir = Join-Path $Root "project"
$venvPython = Join-Path $projectDir "venv\Scripts\python.exe"
$requirements = Join-Path $projectDir "requirements.txt"
$mainScript = Join-Path $projectDir "main.py"

if ((-not (Test-Path -LiteralPath $inputDir)) -and (Test-Path -LiteralPath $legacyInputDir)) {
    $inputDir = $legacyInputDir
}
if ((-not (Test-Path -LiteralPath $outputDir)) -and (Test-Path -LiteralPath $legacyOutputDir)) {
    $outputDir = $legacyOutputDir
}


New-Item -ItemType Directory -Force -Path $inputDir | Out-Null
New-Item -ItemType Directory -Force -Path $outputDir | Out-Null

if (-not (Test-Path -LiteralPath $projectDir)) {
    throw "Project folder not found: $projectDir"
}

if (-not (Test-Path -LiteralPath $venvPython)) {
    Write-Host "Creating virtual environment..."
    python -m venv (Join-Path $projectDir "venv")
}

Write-Host "Installing dependencies..."
& $venvPython -m pip install --disable-pip-version-check -r $requirements

Write-Host ""
Write-Host "Put PDF files into: $inputDir"
Write-Host "Converted Excel files will be written to: $outputDir"
Write-Host ""

& $venvPython $mainScript --input $inputDir --output $outputDir
