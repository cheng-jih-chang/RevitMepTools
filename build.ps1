# build.ps1
# Purpose:
# Build RevitMepLogic and copy the output files to .\dist\
# so RevitAddinHost can load the updated logic assembly.
#
# Usage:
# powershell -ExecutionPolicy Bypass -File .\build.ps1

$ErrorActionPreference = "Stop"

Write-Host "=== BUILD RevitMepLogic ==="
dotnet build .\RevitMepLogic\RevitMepLogic.csproj -v minimal

$srcDll = ".\RevitMepLogic\bin\Debug\net48\RevitMepLogic.dll"
$srcPdb = ".\RevitMepLogic\bin\Debug\net48\RevitMepLogic.pdb"
$distDir = ".\dist"
$dstDll = Join-Path $distDir "RevitMepLogic.dll"
$dstPdb = Join-Path $distDir "RevitMepLogic.pdb"

if (!(Test-Path $srcDll)) {
    throw "Source DLL not found: $srcDll"
}

if (!(Test-Path $distDir)) {
    Write-Host "=== CREATE dist FOLDER ==="
    New-Item -ItemType Directory -Path $distDir | Out-Null
}

Write-Host "=== BEFORE COPY ==="
Write-Host "Source:"
(Get-Item $srcDll).LastWriteTime

Write-Host "Destination:"
if (Test-Path $dstDll) {
    (Get-Item $dstDll).LastWriteTime
} else {
    Write-Host "Destination DLL does not exist yet."
}

Write-Host "=== COPY to dist ==="
Copy-Item $srcDll $dstDll -Force

if (Test-Path $srcPdb) {
    Copy-Item $srcPdb $dstPdb -Force
}

Write-Host "=== AFTER COPY ==="
(Get-Item $dstDll).LastWriteTime
Get-FileHash $dstDll

Write-Host "RevitMepLogic updated"