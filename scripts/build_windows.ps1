$ErrorActionPreference = "Stop"
$PSNativeCommandUseErrorActionPreference = $true

$ProjectRoot = Split-Path -Parent $PSScriptRoot
Set-Location $ProjectRoot

python -m PyInstaller `
    --noconfirm `
    --clean `
    --windowed `
    --name MDtoWORD `
    --icon assets/MDtoWORD.ico `
    --add-data "assets/ico.png;assets" `
    --version-file packaging/windows_version_info.txt `
    md_to_word_converter.py

$BundleDirectory = Join-Path $ProjectRoot "dist/MDtoWORD"
$Executable = Join-Path $BundleDirectory "MDtoWORD.exe"
$Archive = Join-Path $ProjectRoot "dist/MDtoWORD-Windows-x64.zip"
$Checksum = "$Archive.sha256"

if (-not (Test-Path $Executable)) {
    throw "PyInstaller did not create $Executable"
}

if (Test-Path $Archive) {
    Remove-Item $Archive -Force
}

Compress-Archive -Path $BundleDirectory -DestinationPath $Archive -CompressionLevel Optimal

$VerificationDirectory = Join-Path $env:RUNNER_TEMP "mdtoword-windows-verify"
if (Test-Path $VerificationDirectory) {
    Remove-Item $VerificationDirectory -Recurse -Force
}
Expand-Archive -Path $Archive -DestinationPath $VerificationDirectory

$ArchivedExecutable = Join-Path $VerificationDirectory "MDtoWORD/MDtoWORD.exe"
if (-not (Test-Path $ArchivedExecutable)) {
    throw "Archive verification failed: MDtoWORD.exe is missing"
}

$Hash = (Get-FileHash -Path $Archive -Algorithm SHA256).Hash.ToLowerInvariant()
Set-Content -Path $Checksum -Encoding ascii -Value "$Hash  MDtoWORD-Windows-x64.zip"

Write-Host "Windows bundle: $Archive"
Write-Host "SHA-256: $Hash"
