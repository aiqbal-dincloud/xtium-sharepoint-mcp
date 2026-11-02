# Pack Claude Desktop MCPB extension into dist/
$ErrorActionPreference = "Stop"
$Root = Split-Path -Parent $PSScriptRoot
Set-Location $Root

$manifest = Join-Path $Root "manifest.json"
if (-not (Get-Command mcpb -ErrorAction SilentlyContinue)) {
    Write-Error "mcpb CLI not found. Install: npm install -g @anthropic-ai/mcpb"
}

mcpb validate $manifest
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

$meta = Get-Content $manifest -Raw | ConvertFrom-Json
$outDir = Join-Path $Root "dist"
New-Item -ItemType Directory -Force -Path $outDir | Out-Null
$out = Join-Path $outDir "$($meta.name)-$($meta.version).mcpb"

if (Test-Path $out) { Remove-Item $out -Force }
mcpb pack $Root $out
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

mcpb info $out
Write-Host "Packed: $out"
