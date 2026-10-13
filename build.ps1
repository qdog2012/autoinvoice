param([switch]$SkipDependencyInstall)

$ErrorActionPreference = "Stop"

$version = "v1.$(Get-Date -Format 'MMdd.HHmm')"
$sourcePath = Join-Path $PSScriptRoot "autoinvoice.py"
$sourceText = Get-Content -LiteralPath $sourcePath -Raw -Encoding UTF8
$versionPattern = 'APP_VERSION\s*=\s*"v1\.\d{4}\.\d{4}"'
if (-not [regex]::IsMatch($sourceText, $versionPattern)) {
  throw "APP_VERSION not found in autoinvoice.py"
}
$updatedSourceText = [regex]::Replace(
  $sourceText,
  $versionPattern,
  "APP_VERSION = `"$version`"",
  1
)
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
[System.IO.File]::WriteAllText($sourcePath, $updatedSourceText, $utf8NoBom)
Write-Host "Version: $version"

if (-not $SkipDependencyInstall) {
  python -m pip install -r requirements.txt
  if ($LASTEXITCODE -ne 0) { throw "Dependency installation failed" }
}

$pythonPrefix = python -c "import sys; print(sys.prefix)"
$libraryBin = Join-Path $pythonPrefix "Library\bin"
$tclLibrary = python -c "import tkinter; print(tkinter.Tcl().eval('info library'))"
$tkLibrary = Join-Path (Split-Path $tclLibrary -Parent) "tk8.6"

$pyInstallerArgs = @(
  "--noconsole",
  "--onefile",
  "--clean",
  "--name", "AutoInvoice",
  "--hidden-import", "win32timezone",
  "--collect-all", "rapidocr",
  "--collect-binaries", "onnxruntime"
)

if (Test-Path $libraryBin) {
  $pyInstallerArgs += @("--paths", $libraryBin)
  foreach ($dll in @("tcl86t.dll", "tk86t.dll", "liblzma.dll", "libbz2.dll", "LIBBZ2.dll", "ffi.dll", "libexpat.dll", "libmpdec-4.dll")) {
    $dllPath = Join-Path $libraryBin $dll
    if (Test-Path $dllPath) {
      $pyInstallerArgs += @("--add-binary", "$dllPath;.")
    }
  }
}

if (Test-Path $tclLibrary) {
  $pyInstallerArgs += @("--add-data", "$tclLibrary;tcl8.6")
}

if (Test-Path $tkLibrary) {
  $pyInstallerArgs += @("--add-data", "$tkLibrary;tk8.6")
}

$pyInstallerArgs += "autoinvoice.py"

python -m PyInstaller @pyInstallerArgs
if ($LASTEXITCODE -ne 0) { throw "PyInstaller build failed" }

Copy-Item -Path "dist\AutoInvoice.exe" -Destination "AutoInvoice.exe" -Force

Write-Host ""
Write-Host "Build complete: dist\AutoInvoice.exe"
Write-Host "Ready to run: AutoInvoice.exe"
