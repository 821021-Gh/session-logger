#Requires -Version 5.1
# Session Logger v6 — Tracks all user logins/logouts
# Uses query user (primary) — now with correct AM/PM detection
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

function Get-CsvRows {
    if (-not (Test-Path $CsvPath)) { return @() }
    $lines = Get-Content $CsvPath -ErrorAction SilentlyContinue
    if ($lines.Count -le 1) { return @() }
    return @($lines[1..($lines.Count-1)])
}

function AlreadyToday {
    param($account, $type)
    $today = [DateTime]::Now.ToString("yyyy-MM-dd")
    foreach ($r in Get-CsvRows) {
        $p = $r.Split(",")
        if ($p.Count -ge 4 -and $p[0].StartsWith($today) -and $p[1] -eq $type -and $p[2] -eq $account) {
            return $true
        }
    }
    return $false
}

# ── Parse query user ─────────────────────────────────────────────────────────
# The garbled output: 上午/下午 (AM/PM) is Big5 bytes displayed as UTF-8 garbage.
# Big5: 上午 = 3 bytes → 3 UTF-8 chars; 下午 = 4 bytes → 4 UTF-8 chars
# We count the token length to distinguish them.

function Get-QueryUserSessions {
    $tmpOut = "$env:TEMP\_qu_v6.txt"
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

                # tokens[0] = USERNAME or >username
                # tokens[5] = YYYY/M/D
                # tokens[6] = garbled AM/PM (3 chars = 上午=AM, 4 chars = 下午=PM)
                # tokens[7] = HH:MM

                $rawName = $tokens[0] -replace "^>", ""
                $rawName = ($rawName -split ":")[0].Trim()

                if ([string]::IsNullOrWhiteSpace($rawName)) { continue }
                if ($rawName -notmatch "^[\w\-\.]+$") { continue }
                if ($rawName.Length -gt 50) { continue }
                if ($rawName -match "^(USERNAME|SESSIONNAME|console|ID|STATE|Active|Disc|Idle|TIME|none)$") { continue }

                # Parse time — AM=add 0, PM=add 12
                # 上午 (AM) = 3 UTF-8 chars; 下午 (PM) = 4 UTF-8 chars
                $hour = 0; $minute = 0
                $isPM = $false
                if ($tokens.Count -ge 8) {
                    $garbled = $tokens[6]
                    if ($garbled.Length -ge 4) { $isPM = $true }  # 下午 (PM)
                    $hhmm = $tokens[7] -split ":"
                    try {
                        $hour = [int]$hhmm[0]
                        $minute = [int]$hhmm[1]
                    } catch {}
                } elseif ($tokens.Count -ge 7) {
                    $garbled = $tokens[6]
                    $hhmm = $garbled -split ":"
                    if ($hhmm.Count -lt 2) {
                        # tokens[6] is garbled, tokens[7] has time
                        $garbled2 = $tokens[6]
                        if ($garbled2.Length -ge 4) { $isPM = $true }
                        $hhmm = $tokens[7] -split ":"
                        try {
                            $hour = [int]$hhmm[0]
                            $minute = [int]$hhmm[1]
                        } catch {}
                    } else {
                        try {
                            $hour = [int]$hhmm[0]
                            $minute = [int]$hhmm[1]
                        } catch {}
                    }
                }

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

# ── Main ──────────────────────────────────────────────────────────────────────

Init-Csv
$now = [DateTime]::Now
$state = Load-State
$today = $now.ToString("yyyy-MM-dd")

$sessions = Get-QueryUserSessions

# ── 1. Log BOOT once per day ─────────────────────────────────────────────────
$bootLogged = $false
foreach ($r in Get-CsvRows) {
    if ($r -match ",BOOT,System," -and $r.StartsWith($today)) { $bootLogged = $true; break }
}
if (-not $bootLogged) {
    try {
        $os = Get-CimInstance -ClassName Win32_OperatingSystem
        $bootTime = $os.LastBootUpTime
        $mins = [Math]::Round(($now - $bootTime).TotalMinutes)
        Append-Csv "$($bootTime.ToString('yyyy-MM-dd HH:mm:ss')),BOOT,System,$mins"
    } catch {}
}

# ── 2. LOGIN — new sessions not in state.known ───────────────────────────────
$loggedIn = @{}
foreach ($s in $sessions) {
    $acc = $s.account
    $loggedIn[$acc] = $true
    if ($acc -notin $state.known) {
        if (-not (AlreadyToday $acc "LOGIN")) {
            Append-Csv "$($s.logonTime.ToString('yyyy-MM-dd HH:mm:ss')),LOGIN,$acc,0"
        }
    }
}

# ── 3. LOGOUT — known users who are now gone ─────────────────────────────────
foreach ($acc in $state.known) {
    if (-not $loggedIn.ContainsKey($acc)) {
        if (-not (AlreadyToday $acc "LOGOUT")) {
            $minutes = 0
            try {
                $loginTs = [DateTime]::Parse($state.loginTimes[$acc])
                $minutes = [Math]::Round(($now - $loginTs).TotalMinutes)
                if ($minutes -lt 0 -or $minutes -gt 1440) { $minutes = 0 }
            } catch {}
            Append-Csv "$($now.ToString('yyyy-MM-dd HH:mm:ss')),LOGOUT,$acc,$minutes"
        }
    }
}

# ── 4. Update state ──────────────────────────────────────────────────────────
$state.known = @($sessions | ForEach-Object { $_.account })
$state.loginTimes = @{}
foreach ($s in $sessions) {
    $state.loginTimes[$s.account] = $s.logonTime.ToString("yyyy-MM-dd HH:mm:ss")
}
Save-State $state

Write-Host "Sessions: $($sessions.Count) active | State saved."