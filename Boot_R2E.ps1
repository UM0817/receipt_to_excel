param(
    [string]$RepoRoot = "C:\Users\************\receipt_to_excel",
    [string]$DistroName,
    [string]$FrontendUrl = "http://localhost:5173",
    [int]$PostDockerLaunchDelaySec = 15,
    [int]$DockerStartupTimeoutSec = 180,
    [int]$FrontendStartupTimeoutSec = 180
)

$ErrorActionPreference = "Stop"

function Write-Step {
    param([string]$Message)
    Write-Host ""
    Write-Host "==> $Message" -ForegroundColor Cyan
}

function Start-DockerDesktop {
    $dockerDesktopExe = Join-Path $env:ProgramFiles "Docker\Docker\Docker Desktop.exe"

    if (-not (Test-Path $dockerDesktopExe)) {
        throw "Docker Desktop.exe was not found: $dockerDesktopExe"
    }

    $process = Get-Process -Name "Docker Desktop" -ErrorAction SilentlyContinue
    if ($null -eq $process) {
        Write-Step "Starting Docker Desktop"
        Start-Process $dockerDesktopExe | Out-Null
    }
    else {
        Write-Step "Docker Desktop is already running"
    }
}

function Wait-AfterDockerLaunch {
    param([int]$DelaySec)

    if ($DelaySec -le 0) {
        return
    }

    Write-Step "Waiting ${DelaySec}s after Docker Desktop launch"
    Start-Sleep -Seconds $DelaySec
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

function Start-WslTerminal {
    param([string]$Distro)

    if (-not $Distro) {
        Write-Host "No default WSL distro was found, so WSL terminal launch is skipped." -ForegroundColor Yellow
        return
    }

    $wt = Get-Command "wt.exe" -ErrorAction SilentlyContinue
    if ($wt) {
        Write-Step "Starting WSL terminal in Windows Terminal ($Distro)"
        Start-Process $wt.Source -ArgumentList @("-w", "0", "new-tab", "wsl.exe", "-d", $Distro) | Out-Null
        return
    }

    Write-Step "Windows Terminal was not found, starting wsl.exe directly"
    Start-Process "wsl.exe" -ArgumentList @("-d", $Distro) | Out-Null
}

function Start-ComposeUp {
    param(
        [string]$ComposeDir,
        [string]$Distro
    )

    if (-not $Distro) {
        throw "WSL distro was not found. Please run this script with -DistroName."
    }

    $wslComposeDir = Convert-WindowsPathToWslPath -Path $ComposeDir
    $linuxCommand = "cd '$wslComposeDir' && docker compose up"
    $wt = Get-Command "wt.exe" -ErrorAction SilentlyContinue

    if ($wt) {
        Write-Step "Running docker compose up in a new Windows Terminal tab"
        Start-Process $wt.Source -ArgumentList @(
            "-w", "0",
            "new-tab",
            "wsl.exe", "-d", $Distro, "--cd", $wslComposeDir, "--", "bash", "-lc", $linuxCommand
        ) | Out-Null
        return
    }

    Write-Step "Running docker compose up in a new WSL window"
    Start-Process "wsl.exe" -ArgumentList @(
        "-d", $Distro,
        "--cd", $wslComposeDir,
        "--", "bash", "-lc", $linuxCommand
    ) | Out-Null
}

function Wait-DockerDesktopReady {
    param(
        [int]$TimeoutSec,
        [string]$Distro
    )

    Write-Step "Waiting for Docker Desktop to be fully ready"

    $deadline = (Get-Date).AddSeconds($TimeoutSec)
    $engineReady = $false
    $wslReady = $false

    while ((Get-Date) -lt $deadline) {
        if (-not $engineReady) {
            try {
                docker info | Out-Null
                $engineReady = $true
                Write-Host "Docker Engine is ready." -ForegroundColor Green
            }
            catch {
            }
        }

        if ($engineReady -and -not $wslReady -and $Distro) {
            $dockerPath = & wsl.exe -d $Distro -- bash -lc "command -v docker" 2>$null
            if ($LASTEXITCODE -eq 0 -and -not [string]::IsNullOrWhiteSpace(($dockerPath | Out-String))) {
                $wslReady = $true
                Write-Host "Docker command is available in WSL." -ForegroundColor Green
            }
        }

        if ($engineReady -and (($Distro -and $wslReady) -or (-not $Distro))) {
            return
        }

        Start-Sleep -Seconds 3
    }

    if ($engineReady -and -not $wslReady -and $Distro) {
        throw @"
Docker Desktop started, but Docker is still not available in WSL distro '$Distro'.

Please enable Docker Desktop WSL integration:
1. Open Docker Desktop
2. Settings
3. Resources
4. WSL Integration
5. Turn on '$Distro'
"@
    }

    throw "Timed out while waiting for Docker Desktop to be fully ready (${TimeoutSec}s)."
}

function Wait-Frontend {
    param(
        [string]$Url,
        [int]$TimeoutSec
    )

    Write-Step "Waiting for frontend: $Url"

    $deadline = (Get-Date).AddSeconds($TimeoutSec)
    while ((Get-Date) -lt $deadline) {
        try {
            $response = Invoke-WebRequest -Uri $Url -UseBasicParsing -TimeoutSec 5
            if ($response.StatusCode -ge 200 -and $response.StatusCode -lt 500) {
                Write-Host "Frontend is reachable." -ForegroundColor Green
                return
            }
        }
        catch {
            Start-Sleep -Seconds 2
        }
    }

    throw "Timed out while waiting for frontend (${TimeoutSec}s)."
}

function Open-Browser {
    param([string]$Url)

    Write-Step "Opening browser: $Url"
    Start-Process $Url | Out-Null
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

Write-Step "Starting receipt_to_excel"
Write-Host "RepoRoot    : $RepoRoot"
Write-Host "ComposeDir  : $composeDir"
Write-Host "WSL Distro  : $DistroName"
Write-Host "FrontendUrl : $FrontendUrl"

Start-DockerDesktop
Wait-AfterDockerLaunch -DelaySec $PostDockerLaunchDelaySec
Wait-DockerDesktopReady -TimeoutSec $DockerStartupTimeoutSec -Distro $DistroName
Start-WslTerminal -Distro $DistroName
Start-ComposeUp -ComposeDir $composeDir -Distro $DistroName
Wait-Frontend -Url $FrontendUrl -TimeoutSec $FrontendStartupTimeoutSec
Open-Browser -Url $FrontendUrl

Write-Host ""
Write-Host "Startup flow completed." -ForegroundColor Green
