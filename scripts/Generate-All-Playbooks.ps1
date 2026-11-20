<#
.SYNOPSIS
    Regenerates and verifies all three official self-contained func_SMILES2IMG Office Playbooks
    (Excel .xlsx, Word .docx, PowerPoint .pptx) and syncs examples/ + docs/public/examples/.
#>
param(
    [ValidateSet('All', 'Excel', 'Word', 'PowerPoint')]
    [string]$Target = 'All'
)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
$examplesDir = Join-Path $repoRoot 'examples'
$publicExamplesDir = Join-Path $repoRoot 'docs\public\examples'
New-Item -ItemType Directory -Path $examplesDir, $publicExamplesDir -Force | Out-Null

if ($Target -eq 'All' -or $Target -eq 'Excel') {
    & (Join-Path $PSScriptRoot 'Generate-Excel-Playbook.ps1')
    if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
}

if ($Target -eq 'All' -or $Target -eq 'Word') {
    & (Join-Path $PSScriptRoot 'Generate-Word-Playbook.ps1')
    if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
}

if ($Target -eq 'All' -or $Target -eq 'PowerPoint') {
    & (Join-Path $PSScriptRoot 'Generate-PowerPoint-Playbook.ps1')
    if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
}

if ($Target -eq 'All') {
    & (Join-Path $PSScriptRoot 'Generate-Playbook-Previews.ps1')
    if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
}

Write-Host "`n=== All func_SMILES2IMG Office Playbooks Ready ===" -ForegroundColor Cyan
Get-ChildItem -LiteralPath $examplesDir -File | Sort-Object Name | ForEach-Object {
    Write-Host ("  [OK] {0,-42} ({1,6:N1} KB)" -f $_.Name, ($_.Length / 1KB)) -ForegroundColor Green
}
