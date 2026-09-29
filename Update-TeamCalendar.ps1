<#
Update-TeamCalendar.ps1
Scrapes a Winnipeg Central Hockey League team schedule page and writes an .ics file
that calendar apps can subscribe to (webcal).

Settings live in config.json next to this script. Change ScheduleUrl there to point
at a different team or season.

Usage:
  .\Update-TeamCalendar.ps1                  normal run, only rewrites the .ics if games changed
  .\Update-TeamCalendar.ps1 -ShowGames       also print the parsed games
  .\Update-TeamCalendar.ps1 -Force           rewrite even if nothing changed
  .\Update-TeamCalendar.ps1 -HtmlFile p.html parse a saved copy of the page instead of downloading
#>
[CmdletBinding()]
param(
    [string]$ConfigPath = (Join-Path $PSScriptRoot 'config.json'),
    [string]$Url,
    [string]$HtmlFile,
    [switch]$ShowGames,
    [switch]$Force
)

$ErrorActionPreference = 'Stop'

# ---------- config ----------

$cfg = Get-Content -Raw -Path $ConfigPath | ConvertFrom-Json
if ($Url) { $cfg.ScheduleUrl = $Url }
if (-not $cfg.ScheduleUrl) { throw 'ScheduleUrl is empty in config.json' }

$teamName = "$($cfg.TeamName)".Trim()
if (-not $teamName) {
    # Pull the team name from the Team= part of the URL
    if ($cfg.ScheduleUrl -match '[?&]Team=([^&]+)') {
        $teamName = [Uri]::UnescapeDataString($Matches[1].Replace('+', ' ')).Trim()
    }
}

$calName = "$($cfg.CalendarName)".Trim()
if (-not $calName) { $calName = if ($teamName) { $teamName } else { 'Hockey Schedule' } }

$gameMinutes = 60
if ($cfg.GameLengthMinutes) { $gameMinutes = [int]$cfg.GameLengthMinutes }

# One or more reminders, in minutes before the game
$reminders = @()
if ($cfg.ReminderMinutes) { $reminders = @($cfg.ReminderMinutes | ForEach-Object { [int]$_ } | Where-Object { $_ -gt 0 }) }

$outPath = "$($cfg.OutputPath)"
if (-not $outPath) { $outPath = 'team.ics' }
if (-not [IO.Path]::IsPathRooted($outPath)) { $outPath = Join-Path (Split-Path $ConfigPath -Parent) $outPath }

$arenaAddresses = @{}
if ($cfg.ArenaAddresses) {
    foreach ($p in $cfg.ArenaAddresses.PSObject.Properties) { $arenaAddresses[$p.Name.ToLower()] = "$($p.Value)" }
}

# ---------- download ----------

if ($HtmlFile) {
    $html = Get-Content -Raw -Path $HtmlFile
} else {
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $resp = Invoke-WebRequest -Uri $cfg.ScheduleUrl -UseBasicParsing -TimeoutSec 60 `
        -UserAgent 'Mozilla/5.0 (team calendar sync)'
    $html = $resp.Content
}

# ---------- helpers ----------

function Get-CellText([string]$s) {
    $s = $s -replace '(?i)<br\s*/?>', ' '
    $s = $s -replace '<[^>]+>', ' '
    $s = [Net.WebUtility]::HtmlDecode($s)
    ($s -replace '\s+', ' ').Trim()
}

$monthNames = @{ jan=1; feb=2; mar=3; apr=4; may=5; jun=6; jul=7; aug=8; sep=9; oct=10; nov=11; dec=12 }
$dateRx = '(?i)\b(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\.?\s+(\d{1,2})(?:st|nd|rd|th)?(?:,?\s+(\d{4}))?\b'
$numDateRx = '\b(\d{1,2})/(\d{1,2})(?:/(\d{2,4}))?\b'
$timeRx = '(?i)\b(\d{1,2}):(\d{2})\s*([ap])\.?m\.?'

# Season years, used when the page leaves the year off dates
$seasonStart = $null
if ($html -match '(20\d{2})\s*-\s*(20\d{2})\s*Season') { $seasonStart = [int]$Matches[1] }

function Resolve-Year([int]$month) {
    if ($seasonStart) {
        if ($month -ge 7) { return $seasonStart } else { return $seasonStart + 1 }
    }
    # No season on the page, pick the year that puts the date closest to today
    $now = Get-Date
    $best = $now.Year
    $bestGap = [double]::MaxValue
    foreach ($y in ($now.Year - 1), $now.Year, ($now.Year + 1)) {
        $gap = [math]::Abs(((Get-Date -Year $y -Month $month -Day 1) - $now).TotalDays)
        if ($gap -lt $bestGap) { $bestGap = $gap; $best = $y }
    }
    return $best
}

function Find-Date([string]$text) {
    if ($text -match $dateRx) {
        $m = $monthNames[$Matches[1].Substring(0, 3).ToLower()]
        $d = [int]$Matches[2]
        $y = if ($Matches[3]) { [int]$Matches[3] } else { Resolve-Year $m }
        return Get-Date -Year $y -Month $m -Day $d -Hour 0 -Minute 0 -Second 0 -Millisecond 0
    }
    if ($text -match $numDateRx) {
        $m = [int]$Matches[1]; $d = [int]$Matches[2]
        if ($m -ge 1 -and $m -le 12 -and $d -ge 1 -and $d -le 31) {
            $y = if ($Matches[3]) { [int]$Matches[3] } else { Resolve-Year $m }
            if ($y -lt 100) { $y += 2000 }
            return Get-Date -Year $y -Month $m -Day $d -Hour 0 -Minute 0 -Second 0 -Millisecond 0
        }
    }
    return $null
}

function Find-Time([string]$text) {
    if ($text -match $timeRx) {
        $h = [int]$Matches[1] % 12
        if ($Matches[3].ToLower() -eq 'p') { $h += 12 }
        return New-TimeSpan -Hours $h -Minutes ([int]$Matches[2])
    }
    return $null
}

# ---------- parse ----------

$html = $html -replace '(?is)<script\b.*?</script>', '' -replace '(?is)<style\b.*?</style>', ''

# Innermost tables only, so layout tables around the schedule do not confuse the row matching
$tables = [regex]::Matches($html, '(?is)<table\b[^>]*>((?:(?!<table\b).)*?)</table>')

$games = New-Object System.Collections.Generic.List[object]

# Cell text for a row, with colspan cells repeated so columns line up with the header
function Get-RowCells([string]$rowHtml) {
    $out = New-Object System.Collections.Generic.List[string]
    foreach ($c in [regex]::Matches($rowHtml, '(?is)<t[dh]\b([^>]*)>(.*?)</t[dh]>')) {
        $span = 1
        if ($c.Groups[1].Value -match '(?i)colspan\s*=\s*"?(\d+)') { $span = [int]$Matches[1] }
        $text = Get-CellText $c.Groups[2].Value
        for ($k = 0; $k -lt $span; $k++) { $out.Add($text) }
    }
    return ,$out.ToArray()
}

foreach ($t in $tables) {
    $rows = [regex]::Matches($t.Groups[1].Value, '(?is)<tr\b[^>]*>(.*?)</tr>')
    $col = $null
    $currentDate = $null

    foreach ($r in $rows) {
        $rowHtml = $r.Groups[1].Value
        $cells = Get-RowCells $rowHtml
        if ($cells.Count -eq 0) { continue }

        # Look for the header row first.
        # League page: Home / Visitor columns. Team page: Opposition column with "vs" or "at".
        if (-not $col) {
            $lower = @($cells | ForEach-Object { $_.ToLower() })
            $leagueLayout = ($lower -contains 'home') -and (($lower -contains 'visitor') -or ($lower -contains 'away'))
            $teamLayout = ($lower -contains 'opposition') -or ($lower -contains 'opponent')
            if ($leagueLayout -or $teamLayout) {
                $col = @{}
                for ($i = 0; $i -lt $lower.Count; $i++) {
                    switch -Regex ($lower[$i]) {
                        '^home$'                        { $col.Home = $i }
                        '^(visitor|away)$'              { $col.Visitor = $i }
                        '^(opposition|opponent)$'       { $col.Opp = $i }
                        '^time$'                        { $col.Time = $i }
                        '^date$'                        { $col.Date = $i }
                        '^(arena|rink|location|venue)$' { $col.Arena = $i }
                        '^result$' {
                            if (-not $col.ContainsKey('Result')) { $col.Result = $i; $col.ResultEnd = $i }
                            else { $col.ResultEnd = $i }
                        }
                        '^score$' {
                            if ($col.ContainsKey('Visitor')) { $col.VScore = $i } else { $col.HScore = $i }
                        }
                    }
                }
            }
            continue
        }

        $filled = @($cells | Where-Object { $_ -ne '' } | Select-Object -Unique)
        $rowText = $cells -join ' | '

        # Short rows are headings like "Tuesday, September 29" or "Regular Season"
        if ($filled.Count -le 2) {
            $d = Find-Date $rowText
            if ($d) { $currentDate = $d }
            continue
        }

        $date = $null
        if ($col.ContainsKey('Date') -and $cells.Count -gt $col.Date) { $date = Find-Date $cells[$col.Date] }
        if (-not $date) { $date = Find-Date $rowText }
        if (-not $date) { $date = $currentDate }

        $time = $null
        if ($col.ContainsKey('Time') -and $cells.Count -gt $col.Time) { $time = Find-Time $cells[$col.Time] }
        if (-not $time) { $time = Find-Time $rowText }

        if (-not $date -or -not $time) { continue }

        $hs = $null; $vs = $null

        if ($col.ContainsKey('Opp')) {
            if ($cells.Count -le $col.Opp -or -not $teamName) { continue }
            $oppText = $cells[$col.Opp]
            if ($oppText -match '^(?i)(vs\.?|at|@)\s+(.+)$') {
                $isAwayGame = $Matches[1].ToLower() -ne 'vs' -and $Matches[1].ToLower() -ne 'vs.'
                $opp = $Matches[2].Trim()
            } else {
                $isAwayGame = $false
                $opp = $oppText.Trim()
            }
            if (-not $opp) { continue }
            if ($isAwayGame) { $homeTeam = $opp; $vis = $teamName } else { $homeTeam = $teamName; $vis = $opp }

            # Result looks like "-" before the game, and something like "W 4-2" after
            if ($col.ContainsKey('Result')) {
                $resText = ($cells[$col.Result..([math]::Min($col.ResultEnd, $cells.Count - 1))] | Select-Object -Unique) -join ' '
                if ($resText -match '(\d+)\s*-\s*(\d+)') {
                    $a = [int]$Matches[1]; $b = [int]$Matches[2]
                    # Assume our score first, unless a W/L letter says otherwise
                    $us = $a; $them = $b
                    if ($resText -match '(?i)\bW\b' -and $us -lt $them) { $us = $b; $them = $a }
                    if ($resText -match '(?i)\bL\b' -and $us -gt $them) { $us = $b; $them = $a }
                    if ($isAwayGame) { $hs = $them; $vs = $us } else { $hs = $us; $vs = $them }
                }
            }
        } else {
            if ($cells.Count -le [math]::Max($col.Home, $col.Visitor)) { continue }
            $homeTeam = $cells[$col.Home]
            $vis = $cells[$col.Visitor]
            if ($col.ContainsKey('HScore') -and $cells[$col.HScore] -match '^\d+$') { $hs = [int]$cells[$col.HScore] }
            if ($col.ContainsKey('VScore') -and $cells[$col.VScore] -match '^\d+$') { $vs = [int]$cells[$col.VScore] }
        }
        if (-not $homeTeam -or -not $vis) { continue }

        $arena = ''
        if ($col.ContainsKey('Arena') -and $cells.Count -gt $col.Arena) { $arena = $cells[$col.Arena] }

        $gameId = ''
        if ($rowHtml -match '(?i)GameID=(\d+)') { $gameId = $Matches[1] }

        $games.Add([pscustomobject]@{
            Start   = $date.Add($time)
            Home    = $homeTeam
            Visitor = $vis
            Arena   = $arena
            HScore  = $hs
            VScore  = $vs
            GameId  = $gameId
        })
    }
}

if ($games.Count -eq 0) {
    # Leave the old calendar alone if the page failed or the layout changed
    Write-Warning 'No games found on the page. The .ics was not changed.'
    exit 1
}

$games = @($games | Sort-Object Start)

if ($ShowGames) {
    $games | Select-Object @{n="Start";e={$_.Start.ToString("ddd yyyy-MM-dd h:mm tt")}}, Home, HScore, Visitor, VScore, Arena | Format-Table -AutoSize | Out-Host
}

# ---------- build ics ----------

function Escape-Ics([string]$s) {
    $s.Replace('\', '\\').Replace(';', '\;').Replace(',', '\,').Replace("`r`n", '\n').Replace("`n", '\n')
}

function Add-Line($list, [string]$line) {
    # iCal lines must be folded at 75 bytes, 70 chars leaves room for accents
    while ($line.Length -gt 70) {
        $list.Add($line.Substring(0, 70))
        $line = ' ' + $line.Substring(70)
    }
    $list.Add($line)
}

function Get-Uid([string]$seed) {
    $sha = [Security.Cryptography.SHA1]::Create()
    $bytes = $sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($seed.ToLower()))
    (($bytes[0..9] | ForEach-Object { $_.ToString('x2') }) -join '') + '@team-ical'
}

# Current time in Winnipeg, used to decide which games are finished
$tz = $null
foreach ($id in 'America/Winnipeg', 'Central Standard Time') {
    try { $tz = [TimeZoneInfo]::FindSystemTimeZoneById($id); break } catch { }
}
$nowLocal = if ($tz) { [TimeZoneInfo]::ConvertTimeFromUtc([DateTime]::UtcNow, $tz) } else { Get-Date }

$lines = New-Object System.Collections.Generic.List[string]
$lines.Add('BEGIN:VCALENDAR')
$lines.Add('VERSION:2.0')
$lines.Add('PRODID:-//team-ical//Update-TeamCalendar//EN')
$lines.Add('CALSCALE:GREGORIAN')
$lines.Add('METHOD:PUBLISH')
Add-Line $lines ('X-WR-CALNAME:' + (Escape-Ics $calName))
$lines.Add('X-WR-TIMEZONE:America/Winnipeg')
$lines.Add('REFRESH-INTERVAL;VALUE=DURATION:PT6H')
$lines.Add('X-PUBLISHED-TTL:PT6H')

# Central time rules, same as the rest of Manitoba since 2007
@(
    'BEGIN:VTIMEZONE', 'TZID:America/Winnipeg',
    'BEGIN:DAYLIGHT', 'TZOFFSETFROM:-0600', 'TZOFFSETTO:-0500', 'TZNAME:CDT',
    'DTSTART:19700308T020000', 'RRULE:FREQ=YEARLY;BYMONTH=3;BYDAY=2SU', 'END:DAYLIGHT',
    'BEGIN:STANDARD', 'TZOFFSETFROM:-0500', 'TZOFFSETTO:-0600', 'TZNAME:CST',
    'DTSTART:19701101T020000', 'RRULE:FREQ=YEARLY;BYMONTH=11;BYDAY=1SU', 'END:STANDARD',
    'END:VTIMEZONE'
) | ForEach-Object { $lines.Add($_) }

$stamp = [DateTime]::UtcNow.ToString("yyyyMMdd'T'HHmmss'Z'")
$seen = @{}

foreach ($g in $games) {
    $isHome = $teamName -and ($g.Home -ieq $teamName)
    $isAway = $teamName -and ($g.Visitor -ieq $teamName)

    if ($isHome)     { $summary = "$teamName vs $($g.Visitor)" }
    elseif ($isAway) { $summary = "$teamName @ $($g.Home)" }
    else             { $summary = "$($g.Visitor) @ $($g.Home)" }

    $end = $g.Start.AddMinutes($gameMinutes)

    # Add the result once a game is over and a score has been entered
    $desc = "Home: $($g.Home)`nVisitor: $($g.Visitor)"
    if ($end -lt $nowLocal -and $null -ne $g.HScore -and $null -ne $g.VScore -and ($g.HScore + $g.VScore) -gt 0) {
        $desc += "`nFinal: $($g.Home) $($g.HScore), $($g.Visitor) $($g.VScore)"
        if ($isHome -or $isAway) {
            $us = if ($isHome) { $g.HScore } else { $g.VScore }
            $them = if ($isHome) { $g.VScore } else { $g.HScore }
            $res = if ($us -gt $them) { 'W' } elseif ($us -lt $them) { 'L' } else { 'T' }
            $summary += " ($res $us-$them)"
        }
    }
    $desc += "`nSchedule: $($cfg.ScheduleUrl)"

    # UID uses the league GameID, or date and teams, so a time change updates the event instead of duplicating it
    $seed = if ($g.GameId) { 'game|' + $g.GameId } else { $g.Start.ToString('yyyyMMdd') + '|' + $g.Home + '|' + $g.Visitor }
    if ($seen.ContainsKey($seed)) { $seed += '|' + $g.Start.ToString('HHmm') }
    $seen[$seed] = $true

    $location = $g.Arena
    if ($location -and $arenaAddresses.ContainsKey($location.ToLower())) {
        $location = "$location, $($arenaAddresses[$location.ToLower()])"
    }

    $lines.Add('BEGIN:VEVENT')
    $lines.Add('UID:' + (Get-Uid $seed))
    $lines.Add('DTSTAMP:' + $stamp)
    $lines.Add('DTSTART;TZID=America/Winnipeg:' + $g.Start.ToString("yyyyMMdd'T'HHmmss"))
    $lines.Add('DTEND;TZID=America/Winnipeg:' + $end.ToString("yyyyMMdd'T'HHmmss"))
    Add-Line $lines ('SUMMARY:' + (Escape-Ics $summary))
    if ($location) { Add-Line $lines ('LOCATION:' + (Escape-Ics $location)) }
    Add-Line $lines ('DESCRIPTION:' + (Escape-Ics $desc))
    foreach ($reminder in $reminders) {
        $lines.Add('BEGIN:VALARM')
        $lines.Add('ACTION:DISPLAY')
        Add-Line $lines ('DESCRIPTION:' + (Escape-Ics $summary))
        $lines.Add("TRIGGER:-PT$($reminder)M")
        $lines.Add('END:VALARM')
    }
    $lines.Add('END:VEVENT')
}
$lines.Add('END:VCALENDAR')

$newText = ($lines -join "`r`n") + "`r`n"

# ---------- write only if something changed ----------

function Remove-Stamps([string]$s) { ($s -split "`r?`n" | Where-Object { $_ -notmatch '^DTSTAMP:' }) -join "`n" }

if ((Test-Path $outPath) -and -not $Force) {
    $oldText = [IO.File]::ReadAllText($outPath)
    if ((Remove-Stamps $oldText) -eq (Remove-Stamps $newText)) {
        Write-Host "No changes. $($games.Count) games, calendar left as is."
        exit 0
    }
}

$dir = Split-Path $outPath -Parent
if ($dir -and -not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir | Out-Null }
[IO.File]::WriteAllText($outPath, $newText, (New-Object Text.UTF8Encoding $false))
Write-Host "Wrote $($games.Count) games to $outPath"
