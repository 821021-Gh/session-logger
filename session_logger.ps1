#Requires -Version 5.1
# Session Logger v7 — Tracks all user logins/logouts
# Logic: compare current sessions vs previous state
#   - New account appeared  → write LOGIN
#   - Known account gone    → write LOGOUT (with session duration)
# Runs every minute via SessionLogger_Track scheduled task

param(
    [string]$CsvPath = "$env:USERPROFILE\session_log.csv",
    [string]$StateFile = "$env:TEMP\session_state.json"
)

$ErrorActionPreference = "SilentlyContinue"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# ── CSV Helpers ────────────────────────────────────────────────────────────────

function Init-Csv {
    if (-not (Test-Path $CsvPath)) {
        New-Item -Path $CsvPath -ItemType File -Force | Out-Null
        Set-Content -Path $CsvPath -Value "timestamp,event_type,account,duration_minutes" -Encoding UTF8
    }
}

function Append-Csv {
    param($line)
    Add-Content -Path $CsvPath -Value $line -Encoding UTF8
}

# ── Parse query user ──────────────────────────────────────────────────────────
# Big5/UTF-8 garbled AM/PM marker:
#   上午 (AM) = 3 UTF-8 chars, 下午 (PM) = 4 UTF-8 chars

function Get-QueryUserSessions {
    $tmpOut = "$env:TEMP\_qu_v7.txt"
    Remove-Item $tmpOut -Force -ErrorAction SilentlyContinue
    $null = Start-Process "query.exe" -ArgumentList "user" -NoNewWindow -Wait -PassThru `
        -RedirectStandardOutput $tmpOut -RedirectStandardError "$env:TEMP\_qu_err.txt"

    $sessions = @()
    if (Test-Path $tmpOut) {
        $raw = Get-Content $tmpOut -Raw -ErrorAction SilentlyContinue
        if ($raw) {
            $lines = $raw -split "`r?`n" | Where-Object { $_.Trim() -ne "" }
            for ($i = 1; $i -lt $lines.Count; $i++) {
                $line = $lines[$i]
                if ($line -notmatch "\s(Active|Disc)\s") { continue }

                $tokens = $line.Trim() -split '\s+'
                if ($tokens.Count -lt 7) { continue }

                $rawName = $tokens[0] -replace "^>", ""
                $rawName = ($rawName -split ":")[0].Trim()

                if ([string]::IsNullOrWhiteSpace($rawName)) { continue }
                if ($rawName -notmatch "^[\w\-\.]+$") { continue }
                if ($rawName.Length -gt 50) { continue }

                # Parse time — detect PM by garbled token length
                $hour = 0; $minute = 0
                $isPM = $false
                if ($tokens.Count -ge 8) {
                    if ($tokens[6].Length -ge 4) { $isPM = $true }
                    $hhmm = $tokens[7] -split ":"
                } elseif ($tokens.Count -ge 7) {
                    $hhmm = $tokens[6] -split ":"
                    if ($hhmm.Count -lt 2) {
                        $hhmm = $tokens[7] -split ":"
                        if ($tokens[6].Length -ge 4) { $isPM = $true }
                    }
                }
                try {
                    $hour = [int]$hhmm[0]
                    $minute = [int]$hhmm[1]
                } catch {}

                if ($isPM -and $hour -lt 12) { $hour += 12 }
                if ($hour -ge 24) { $hour -= 24 }

                try {
                    $logonDate = [DateTime]::Parse($tokens[5])
                    $logonTime = Get-Date -Year $logonDate.Year -Month $logonDate.Month -Day $logonDate.Day -Hour $hour -Minute $minute -Second 0
                } catch {
                    $logonTime = [DateTime]::Now
                }

                $sessions += @{ account = $rawName; logonTime = $logonTime }
            }
        }
    }
    Remove-Item $tmpOut -Force -ErrorAction SilentlyContinue
    return $sessions
}

# ── State ─────────────────────────────────────────────────────────────────────

function Load-State {
    if (Test-Path $StateFile) {
        try {
            $content = Get-Content $StateFile -Raw
            return $content | ConvertFrom-Json
        } catch {}
    }
    return @{ known = @(); loginTimes = @{}; lastLogonSnapshots = @{} }
}

function Save-State {
    param($state)
    $json = $state | ConvertTo-Json -Depth 5
    [System.IO.File]::WriteAllText($StateFile, $json, [System.Text.Encoding]::UTF8)
}

# ── BOOT ──────────────────────────────────────────────────────────────────────

function Log-BootOnce {
    $today = [DateTime]::Now.ToString("yyyy-MM-dd")
    if (Test-Path $CsvPath) {
        $lines = Get-Content $CsvPath -ErrorAction SilentlyContinue
        foreach ($r in $lines) {
            if ($r -match ",BOOT,System," -and $r.StartsWith($today)) { return }
        }
    }
    try {
        $os = Get-CimInstance -ClassName Win32_OperatingSystem
        $bootTime = $os.LastBootUpTime
        $mins = [Math]::Round(([DateTime]::Now - $bootTime).TotalMinutes)
        Append-Csv "$($bootTime.ToString('yyyy-MM-dd HH:mm:ss')),BOOT,System,$mins"
    } catch {}
}

# ── Main ──────────────────────────────────────────────────────────────────────

Init-Csv
$now = [DateTime]::Now
$state = Load-State

# Get current active sessions
$sessions = Get-QueryUserSessions

# Build lookup hashmap
$currentActive = @{}
foreach ($s in $sessions) { $currentActive[$s.account] = $s }

# ── LOGIN: in current sessions but NOT in previous known list ─────────────────
foreach ($s in $sessions) {
    $acc = $s.account
    if ($acc -notin $state.known) {
        Append-Csv "$($s.logonTime.ToString('yyyy-MM-dd HH:mm:ss')),LOGIN,$acc,0"
    }
}

# ── LOGOUT: was in previous known list but NOT in current sessions ────────────
foreach ($prevAcc in $state.known) {
    if (-not $currentActive.ContainsKey($prevAcc)) {
        $minutes = 0
        try {
            $loginTs = [DateTime]::Parse($state.loginTimes[$prevAcc])
            $minutes = [Math]::Round(($now - $loginTs).TotalMinutes)
            if ($minutes -lt 0 -or $minutes -gt 1440) { $minutes = 0 }
        } catch {}
        Append-Csv "$($now.ToString('yyyy-MM-dd HH:mm:ss')),LOGOUT,$prevAcc,$minutes"
    }
}

# ── Update state ──────────────────────────────────────────────────────────────
$state.known = @($sessions | ForEach-Object { $_.account })
$state.loginTimes = @{}
foreach ($s in $sessions) {
    $state.loginTimes[$s.account] = $s.logonTime.ToString("yyyy-MM-dd HH:mm:ss")
}
Save-State $state

# ── BOOT ──────────────────────────────────────────────────────────────────────
Log-BootOnce

Write-Output "OK: $($sessions.Count) active | LOGIN=$($sessions.Count) LOGOUT=$( [Math]::Max(0, $state.known.Count - $sessions.Count) )"