param(
    [string]$RepoRoot = "C:\Users\************\receipt_to_excel",
    [string]$DistroName,
    [int]$DockerShutdownDelaySec = 5,
    [switch]$ShutdownAllWsl
)

$ErrorActionPreference = "Stop"

function Write-Step {
    param([string]$Message)
    Write-Host ""
    Write-Host "==> $Message" -ForegroundColor Cyan
}

function Get-DefaultWslDistro {
    $lines = wsl -l -v 2>$null
    foreach ($line in $lines) {
        $cleanLine = ($line -replace "`0", "").Trim()
        if ($cleanLine -match "^\*\s+(?<name>\S+)") {
            return $Matches["name"]
        }
    }

    foreach ($line in $lines) {
        $cleanLine = ($line -replace "`0", "").Trim()
        if (
            $cleanLine -and
            $cleanLine -notmatch "^NAME\s+STATE\s+VERSION$" -and
            $cleanLine -notmatch "^docker-desktop" -and
            $cleanLine -notmatch "^docker-desktop-data"
        ) {
            return ($cleanLine -split "\s+")[0]
        }
    }

    return $null
}

function Convert-WindowsPathToWslPath {
    param([string]$Path)

    $resolved = [System.IO.Path]::GetFullPath($Path)
    $drive = $resolved.Substring(0, 1).ToLower()
    $tail = $resolved.Substring(2).Replace("\", "/")
    return "/mnt/$drive$tail"
}

function Stop-Compose {
    param(
        [string]$ComposeDir,
        [string]$Distro
    )

    if (-not $Distro) {
        throw "WSL distro was not found. Please run this script with -DistroName."
    }

    $wslComposeDir = Convert-WindowsPathToWslPath -Path $ComposeDir
    Write-Step "Stopping docker compose in WSL ($Distro)"
    & wsl.exe -d $Distro --cd $wslComposeDir -- bash -lc "cd '$wslComposeDir' && docker compose down"

    if ($LASTEXITCODE -ne 0) {
        throw "docker compose down failed."
    }
}

function Stop-Wsl {
    param(
        [string]$Distro,
        [bool]$StopAll
    )

    if ($StopAll) {
        Write-Step "Shutting down all WSL instances"
        & wsl.exe --shutdown
        Start-Sleep -Seconds 2
        return
    }

    if ($Distro) {
        Write-Step "Terminating WSL distro ($Distro)"
        & wsl.exe --terminate $Distro
        Start-Sleep -Seconds 2
    }
}

function Wait-ForDockerDesktopToStop {
    param([int]$TimeoutSec = 30)

    $deadline = (Get-Date).AddSeconds($TimeoutSec)
    while ((Get-Date) -lt $deadline) {
        $dockerProcesses = Get-Process -ErrorAction SilentlyContinue | Where-Object {
            $_.ProcessName -like "Docker Desktop" -or
            $_.ProcessName -like "com.docker.*"
        }

        if (-not $dockerProcesses) {
            Write-Host "Docker Desktop is stopped." -ForegroundColor Green
            return
        }

        Start-Sleep -Seconds 2
    }

    throw "Timed out while waiting for Docker Desktop to stop."
}

function Stop-DockerDesktop {
    param([int]$DelaySec)

    if ($DelaySec -gt 0) {
        Write-Step "Waiting ${DelaySec}s before stopping Docker Desktop"
        Start-Sleep -Seconds $DelaySec
    }

    $dockerCli = Join-Path $env:ProgramFiles "Docker\Docker\DockerCli.exe"
    $dockerProcesses = Get-Process -ErrorAction SilentlyContinue | Where-Object {
        $_.ProcessName -like "Docker Desktop" -or
        $_.ProcessName -like "com.docker.*"
    }

    if (-not $dockerProcesses) {
        Write-Step "Docker Desktop is already stopped"
        return
    }

    if (Test-Path $dockerCli) {
        Write-Step "Stopping Docker Desktop via DockerCli.exe"
        & $dockerCli -Shutdown | Out-Null
        Wait-ForDockerDesktopToStop
        return
    }

    Write-Step "Stopping Docker Desktop processes"
    $dockerProcesses | Stop-Process -Force
    Wait-ForDockerDesktopToStop
}

function Confirm-WslStopped {
    param(
        [string]$Distro,
        [bool]$StopAll
    )

    Write-Step "Checking WSL state"
    $lines = wsl -l -v 2>$null
    foreach ($line in $lines) {
        $cleanLine = ($line -replace "`0", "").Trim()
        if (-not $cleanLine -or $cleanLine -match "^NAME\s+STATE\s+VERSION$") {
            continue
        }

        if ($StopAll) {
            if ($cleanLine -match "^(?<name>\S+)\s+(?<state>Running|Stopped)\s+\d+$" -and $Matches["state"] -eq "Running") {
                throw "WSL distro is still running: $($Matches["name"])"
            }
        }
        elseif ($Distro -and $cleanLine -match "^\*?\s*(?<name>\S+)\s+(?<state>Running|Stopped)\s+\d+$") {
            if ($Matches["name"] -eq $Distro -and $Matches["state"] -eq "Running") {
                throw "WSL distro is still running: $Distro"
            }
        }
    }

    Write-Host "WSL is stopped." -ForegroundColor Green
}

$composeDir = Join-Path $RepoRoot "docker"
if (-not (Test-Path $composeDir)) {
    throw "docker compose directory was not found: $composeDir"
}

if (-not $DistroName) {
    $DistroName = Get-DefaultWslDistro
}

if (-not $DistroName) {
    throw "No default WSL distro was found. Please specify -DistroName, for example: Ubuntu-22.04"
}

Write-Step "Stopping receipt_to_excel"
Write-Host "RepoRoot   : $RepoRoot"
Write-Host "ComposeDir : $composeDir"
Write-Host "WSL Distro : $DistroName"

Stop-Compose -ComposeDir $composeDir -Distro $DistroName
Stop-DockerDesktop -DelaySec $DockerShutdownDelaySec
Stop-Wsl -Distro $DistroName -StopAll ([bool]$ShutdownAllWsl)
Confirm-WslStopped -Distro $DistroName -StopAll ([bool]$ShutdownAllWsl)

Write-Host ""
Write-Host "Shutdown flow completed." -ForegroundColor Green
