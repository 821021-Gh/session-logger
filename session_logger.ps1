#Requires -Version 5.1
# Session Logger v8 — Tracks all user logins/logouts
# Logic:
#   - BOOT: once per day (tracked in state)
#   - LOGIN: account appeared in query user but not in previous known list
#   - LOGOUT: account was in previous known list but now gone
#   - Both Active AND Disc sessions count as online
# State file: clean JSON, written atomically (no BOM)

param(
    [string]$CsvPath    = "$env:USERPROFILE\session_log.csv",
    [string]$StatePath  = "$env:TEMP\session_state.json"
)

$ErrorActionPreference = "SilentlyContinue"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# ── CSV ─────────────────────────────────────────────────────────────────────

function Init-Csv {
    if (-not (Test-Path $CsvPath)) {
        New-Item -Path $CsvPath -ItemType File -Force | Out-Null
        Set-Content -Path $CsvPath -Value "timestamp,event_type,account,duration_minutes" -Encoding UTF8
    }
}

function Append-Csv {
    param($type, $account, $minutes, $ts)
    Add-Content -Path $CsvPath -Value "$ts,$type,$account,$minutes" -Encoding UTF8
}

# ── State (clean JSON, no BOM) ──────────────────────────────────────────────

function Load-State {
    if (Test-Path $StatePath) {
        $raw = Get-Content $StatePath -Raw
        if ($raw) { try { return $raw | ConvertFrom-Json } catch {} }
    }
    return @{ known = @(); loginTimes = @{}; lastBootDate = $null }
}

function Save-State {
    param($state)
    $tmp = $StatePath + ".tmp"
    $json = $state | ConvertTo-Json -Depth 4
    [System.IO.File]::WriteAllText($tmp, $json, [System.Text.Encoding]::UTF8)
    Move-Item -Path $tmp -Destination $StatePath -Force
}

# ── query user parser ──────────────────────────────────────────────────────
# Sessions include both Active AND Disc (disconnected = still logged in)
# AM/PM garbled as UTF-8: AM = 3 chars, PM = 4 chars

function Get-CurrentSessions {
    $outFile = "$env:TEMP\_qu_v8.txt"
    Remove-Item $outFile -Force -ErrorAction SilentlyContinue
    $null = Start-Process "query.exe" -ArgumentList "user" -NoNewWindow -Wait `
        -PassThru -RedirectStandardOutput $outFile -RedirectStandardError "$env:TEMP\_qu_err.txt"

    $sessions = @()
    if (Test-Path $outFile) {
        $raw = Get-Content $outFile -Raw
        if ($raw) {
            $lines = $raw -split "`r?`n"
            for ($i = 1; $i -lt $lines.Count; $i++) {
                $line = $lines[$i].Trim()
                if ($line -notmatch "\s(Active|Disc)\s") { continue }

                $tok = $line -split '\s+'
                if ($tok.Count -lt 7) { continue }

                $name = $tok[0] -replace '^>', ''
                $name = ($name -split ':')[0].Trim()
                if ([string]::IsNullOrWhiteSpace($name)) { continue }
                if ($name -notmatch '^[\w\-\.]+$') { continue }

                $isPM = $tok.Count -ge 8 -and $tok[6].Length -ge 4
                $hh = 0; $mm = 0
                try {
                    if ($tok.Count -ge 8) {
                        $hh = [int]($tok[7] -split ':')[0]
                        $mm = [int]($tok[7] -split ':')[1]
                    } elseif ($tok.Count -ge 7) {
                        $p = $tok[6] -split ':'
                        if ($p.Count -ge 2) { $hh = [int]$p[0]; $mm = [int]$p[1] }
                        else { $hh = [int]($tok[7] -split ':')[0]; $mm = [int]($tok[7] -split ':')[1] }
                    }
                } catch {}

                if ($isPM -and $hh -lt 12) { $hh += 12 }
                if ($hh -ge 24) { $hh -= 24 }

                try {
                    $d = [DateTime]::Parse($tok[5])
                    $logon = Get-Date -Year $d.Year -Month $d.Month -Day $d.Day -Hour $hh -Minute $mm -Second 0
                } catch { $logon = [DateTime]::Now }

                $sessions += @{ account = $name; logonTime = $logon }
            }
        }
    }
    Remove-Item $outFile -Force -ErrorAction SilentlyContinue
    return $sessions
}

# ── BOOT ──────────────────────────────────────────────────────────────────

function Handle-Boot {
    param($state, $now)
    $today = $now.ToString('yyyy-MM-dd')
    if ($state.lastBootDate -eq $today) { return $state }
    try {
        $os = Get-CimInstance Win32_OperatingSystem
        $bootTime = $os.LastBootUpTime
        $mins = [Math]::Round(($now - $bootTime).TotalMinutes)
        if ($mins -lt 0 -or $mins -gt 1440) { $mins = 0 }
        Append-Csv "BOOT" "System" $mins $bootTime
    } catch {}
    $state.lastBootDate = $today
    return $state
}

# ── Main ─────────────────────────────────────────────────────────────────

Init-Csv
$now = [DateTime]::Now
$state = Load-State

$state = Handle-Boot $state $now
$sessions = Get-CurrentSessions

$prevKnown = @($state.known)
$currNames = @($sessions | ForEach-Object { $_.account })
$currLookup = @{}; foreach ($s in $sessions) { $currLookup[$s.account] = $s }

foreach ($s in $sessions) {
    if ($s.account -notin $prevKnown) {
        Append-Csv "LOGIN" $s.account 0 $s.logonTime
    }
}

foreach ($prev in $prevKnown) {
    if (-not $currLookup.ContainsKey($prev)) {
        $mins = 0
        if ($state.loginTimes.$prev) {
            try {
                $loginTs = [DateTime]::Parse($state.loginTimes.$prev)
                $mins = [Math]::Round(($now - $loginTs).TotalMinutes)
                if ($mins -lt 0 -or $mins -gt 1440) { $mins = 0 }
            } catch {}
        }
        Append-Csv "LOGOUT" $prev $mins $now
    }
}

$state.known = $currNames
$state.loginTimes = @{}
foreach ($s in $sessions) {
    $state.loginTimes[$s.account] = $s.logonTime.ToString('yyyy-MM-dd HH:mm:ss')
}
Save-State $state

Write-Output "v8 OK: $($sessions.Count) active | known=$($state.known.Count)"