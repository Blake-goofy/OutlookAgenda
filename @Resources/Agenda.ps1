param(
    [int]$Days = 5,
    [int]$Max = 12,
    [string]$TimeFormat = 'h:mm tt',
    [string]$DateFormat = 'ddd MMM d',
    [string]$OutInc = $(Join-Path $PSScriptRoot 'Agenda.inc'),
    [string]$OutJson = $(Join-Path $PSScriptRoot 'agenda.json')
)

# ----------------------------------------------
# Outlook helpers
# ----------------------------------------------
function Get-OutlookNamespace {
    try {
        $ol = New-Object -ComObject Outlook.Application
        return $ol.GetNamespace('MAPI')
    } catch {
        Write-Warning "Could not connect to Outlook: $_"
        return $null
    }
}

function Get-DefaultCalendarFolder {
    param($namespace)
    try {
        if (-not $namespace) { return $null }
        return $namespace.GetDefaultFolder(9)  # 9 = olFolderCalendar
    } catch {
        Write-Warning "Could not access calendar folder: $_"
        return $null
    }
}

function Get-UpcomingEvents {
    param($calendar, $start, $end, $max)
    $events = @()
    try {
        if (-not $calendar) { return @() }
        $items = $calendar.Items
        $items.Sort("[Start]")
        $items.IncludeRecurrences = $true
        $filter = "[Start] >= '$($start.ToString('g'))' AND [Start] < '$($end.ToString('g'))'"
        $result = $items.Restrict($filter)
        foreach ($item in $result) {
            $events += $item
            if ($events.Count -ge $max) { break }
        }
        return $events | Sort-Object Start | Select-Object -First $max
    } catch {
        Write-Warning "Could not retrieve events: $_"
        return @()
    }
}

# ----------------------------------------------
# Main
# ----------------------------------------------
$ns = Get-OutlookNamespace
$cal = Get-DefaultCalendarFolder -namespace $ns
$now = Get-Date
$end = $now.AddDays($Days)
$events = Get-UpcomingEvents -calendar $cal -start $now.Date -end $end -max $Max

# Build grouped model
$dayGroups = @{}
if ($events -and $events.Count -gt 0) {
    foreach ($evt in $events) {
        try {
            $d = [DateTime]$evt.Start
            $dateHeader = $d.ToString($DateFormat)

            if (-not $dayGroups.ContainsKey($dateHeader)) {
                $dayGroups[$dateHeader] = @()
            }

            $timePart = $d.ToString($TimeFormat)
            $subject = [string]$evt.Subject
            if ([string]::IsNullOrWhiteSpace($subject)) { $subject = '(No subject)' }

            # Truncate subject to keep line neat
            $maxSubjectLength = 20
            if ($subject.Length -gt $maxSubjectLength) {
                $subject = $subject.Substring(0, $maxSubjectLength - 3) + '...'
            }

            # Compute end time and status
            $eventEnd = $d
            try { if ($evt.End) { $eventEnd = [DateTime]$evt.End } else { $eventEnd = $d.AddMinutes(30) } } catch {}

            $isUpcoming = $d -gt $now -and $d -le $now.AddMinutes(15)
            $isActive   = $d -le $now -and $eventEnd -gt $now
            $isPast     = $eventEnd -lt $now

            $status = if ($isPast) { 'past' } elseif ($isUpcoming -or $isActive) { 'soon' } else { 'normal' }

            $dayGroups[$dateHeader] += [PSCustomObject]@{
                Start   = $d
                End     = $eventEnd
                Subject = $subject
                Status  = $status
                Line    = "   $timePart $subject"
            }
        } catch {
            Write-Warning "Error processing event: $_"
        }
    }
}

# ----------------------------------------------
# Generate include with meters
# ----------------------------------------------
$metersSection = @()
$lastMeterName = 'MeterDivider'
$metersSection += @"
[MeterAnchor]
Meter=Shape
X=(#Padding#+10)
Y=([MeterDivider:YH])
Shape=Rectangle 0,0,0,0
DynamicVariables=1

"@
$dayIndex = 0

if ($dayGroups.Count -eq 0) {
    $metersSection += @"
[MeterNoEvents]
Meter=String
Text=No upcoming events
X=(#Padding#+10)
Y=([${lastMeterName}:YH]+8)
W=(#Width#-2*#Padding#)
FontFace=#FontName#
FontSize=#FontSize#
FontColor=#FontColor#
AntiAlias=1
StringAlign=Left
DynamicVariables=1

"@
    $lastMeterName = 'MeterNoEvents'
} else {
    foreach ($day in ($dayGroups.Keys | Sort-Object { [DateTime]::ParseExact($_, $DateFormat, $null) })) {
        $dayIndex++
        $dayText = $day
    $headerMeterName = "MeterDay${dayIndex}Header"
    $yHeader = 'Y=8R'
    $metersSection += @"
[$headerMeterName]
Meter=String
Text=$dayText
X=(#Padding#+10)
$yHeader
FontFace=#FontName#
FontSize=#FontSize#
FontColor=#DateColor#
FontWeight=700
AntiAlias=1
StringAlign=Left
DynamicVariables=1

"@
        $lastMeterName = $headerMeterName

        foreach ($evt in $dayGroups[$day]) {
            $eventId = -join ((1..6) | ForEach-Object { '{0:X}' -f (Get-Random -Max 16) })
            $eventMeterName = "MeterDay${dayIndex}Evt$eventId"

            switch ($evt.Status) {
                'soon'   { $color = '255,255,0,255' }
                'past'   { $color = '128,128,128,255' }
                default  { $color = '#FontColor#' }
            }

            $metersSection += @"
[$eventMeterName]
Meter=String
Text=$($evt.Line)
X=(#Padding#+10)
Y=2R
W=(#Width#-2*#Padding#)
FontFace=#FontName#
FontSize=#FontSize#
FontColor=$color
AntiAlias=1
ClipString=2
StringAlign=Left
DynamicVariables=1

"@
            $lastMeterName = $eventMeterName
        }
    }
}

# Terminal marker so background can size correctly
$metersSection += @"
[MeterContentEnd]
Meter=Shape
X=0
Y=6R
Shape=Rectangle 0,0,0,0
DynamicVariables=1

"@

# Ensure directory exists and write include
try {
    $null = New-Item -ItemType Directory -Path (Split-Path $OutInc) -ErrorAction SilentlyContinue
    ($metersSection -join '') | Out-File -FilePath $OutInc -Encoding UTF8 -Force
    
    # Write LastUpdated variable to ini file
    $iniPath = Split-Path (Split-Path $OutInc) | Join-Path -ChildPath 'Agenda.ini'
    $updatedAt = (Get-Date).ToString($TimeFormat)
    $iniContent = Get-Content -Path $iniPath -Raw
    $iniContent = $iniContent -replace 'LastUpdated=.*', "LastUpdated=Last updated: $updatedAt"
    $iniContent | Out-File -FilePath $iniPath -Encoding UTF8 -Force
} catch {
    Write-Warning "Failed to write include file '$OutInc': $_"
}

# ----------------------------------------------
# Write JSON for other consumers (optional)
# ----------------------------------------------
try {
    $daysJson = @()
    foreach ($day in ($dayGroups.Keys | Sort-Object { [DateTime]::ParseExact($_, $DateFormat, $null) })) {
        $eventsJson = @()
        foreach ($evt in $dayGroups[$day]) {
            $eventsJson += [PSCustomObject]@{
                start   = $evt.Start.ToString('o')
                end     = $evt.End.ToString('o')
                subject = $evt.Subject
                status  = $evt.Status
                text    = $evt.Line
            }
        }
        $daysJson += [PSCustomObject]@{
            dateHeader = $day
            date       = ([DateTime]::ParseExact($day, $DateFormat, $null)).ToString('yyyy-MM-dd')
            events     = $eventsJson
        }
    }
    $model = [PSCustomObject]@{
        generatedAt = (Get-Date).ToString('o')
        daysToShow  = $Days
        maxItems    = $Max
        timeFormat  = $TimeFormat
        dateFormat  = $DateFormat
        days        = $daysJson
    }
    $json = $model | ConvertTo-Json -Depth 6
    $json | Out-File -FilePath $OutJson -Encoding UTF8 -Force
} catch {
    Write-Warning "Failed to write JSON '$OutJson': $_"
}

Write-Output "Agenda include generated."
