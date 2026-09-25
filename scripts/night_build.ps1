# Pure-Python run (no AI): fetch missing data, build all parts, downloaded-CAD materials, Excel copy.
# Usage:  powershell -ExecutionPolicy Bypass -File night_build.ps1 -Project "C:\path\to\project"
param([Parameter(Mandatory = $true)][string]$Project)
$env:FASTENER_PROJECT = (Resolve-Path $Project).Path
$s = $PSScriptRoot
New-Item -ItemType Directory -Force (Join-Path $env:FASTENER_PROJECT "data") | Out-Null
$log = Join-Path $env:FASTENER_PROJECT "data\night_$(Get-Date -Format yyyyMMdd_HHmm).log"
python "$s\fetch_all.py" *>> $log
python "$s\build_parts.py" --all *>> $log
python "$s\apply_materials.py" *>> $log
python "$s\fill_excel.py" *>> $log
"klaar $(Get-Date)" | Out-File -Append $log
