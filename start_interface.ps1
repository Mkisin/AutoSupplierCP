$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $root

$port = 8020
$python = Join-Path $root ".venv\Scripts\python.exe"
if (-not (Test-Path $python)) {
  $python = "python"
}

function Get-ListenerPid([int]$TargetPort) {
  $line = netstat -ano -p tcp | Select-String -Pattern "^\s*TCP\s+\S+:$TargetPort\s+\S+\s+LISTENING\s+(\d+)\s*$" | Select-Object -First 1
  if (-not $line) {
    return $null
  }

  $match = [regex]::Match($line.Line, "LISTENING\s+(\d+)\s*$")
  if (-not $match.Success) {
    return $null
  }

  return [int]$match.Groups[1].Value
}

function Stop-StaleProjectServer([int]$TargetPort) {
  $listenerPid = Get-ListenerPid $TargetPort
  if (-not $listenerPid) {
    return
  }

  try {
    $process = Get-Process -Id $listenerPid -ErrorAction Stop
  } catch {
    throw "Port $TargetPort is busy, but the process with PID $listenerPid could not be inspected."
  }

  $processPath = ""
  try {
    $processPath = $process.Path
  } catch {
    $processPath = ""
  }

  $isProjectPython = $process.ProcessName -ieq "python" -and $processPath -like "$root*"
  $isVenvPython = $processPath -ieq $python

  if (-not ($isProjectPython -or $isVenvPython)) {
    throw "Port $TargetPort is already used by $($process.ProcessName) (PID $listenerPid). The script will not stop a non-project process automatically."
  }

  Write-Host "Freeing port ${TargetPort}; stopping stale process PID $listenerPid..." -ForegroundColor Yellow
  Stop-Process -Id $listenerPid -Force

  Start-Sleep -Milliseconds 700
  $stillListening = Get-ListenerPid $TargetPort
  if ($stillListening) {
    throw "Failed to free port $TargetPort. It is still held by PID $stillListening."
  }
}

Stop-StaleProjectServer -TargetPort $port

Write-Host "Starting app on http://127.0.0.1:${port}" -ForegroundColor Green
& $python -m uvicorn backend.app:app --host 127.0.0.1 --port $port
