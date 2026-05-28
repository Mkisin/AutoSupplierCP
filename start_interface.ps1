$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $root

python -m uvicorn backend.app:app --host 0.0.0.0 --port 8020
