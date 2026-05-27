$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $root

python -m uvicorn backend.app:app --host 127.0.0.1 --port 8020
