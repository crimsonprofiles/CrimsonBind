#Requires -Version 5.1
<#
.SYNOPSIS
  CrimsonBind: CSV-driven keybind manager. Export/import config.ini, filter by section,
  randomize keys, update config.ini and Debounce addon (no WoW required).

.DESCRIPTION
  Use a single CSV (Section, ActionName, Key, MacroText) as source of truth. Export from
  config.ini to CSV, edit in Excel, then Update Config and/or Update Debounce to write
  binds into the Debounce addon for use in WoW. Filter dropdown shows General vs
  class/spec. Randomize assigns keys from a pool.

  CSV format: Section, ActionName, MacroText, Key, TextureID
  (MacroText is optional for config.ini value-after-semicolon. Key column may be named
  BindPadKey in older CSVs for compatibility.)

.EXAMPLE
  .\CrimsonBind.ps1
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$Script:ConfigPath = ""
$Script:CsvPath = ""
$Script:SavedVarsPath = ""
$Script:WtfDir = $null
$Script:Accounts = @()
$Script:CurrentFilterSection = $null
$Script:GridSearchText = ""
$Script:AllRows = [System.Collections.ArrayList]::new()  # [{ Section, ActionName, MacroText, Key, TextureID }]
$ToolDir = $PSScriptRoot
$Script:BackupRoot = Join-Path $ToolDir "BACKUP"
$Script:ExcludedKeysPath = Join-Path $ToolDir "excluded_keys.csv"
$Script:EnableLogging = $false   # Set $true to write log file under ToolDir\Logs
$Script:ConfigIniKeySuffix = "67701769"   # Suffix for config.ini key format (e.g. +vk70_67701769)

# Optional: write to log file for diagnostics. No-op if $Script:EnableLogging is $false.
function Write-CrimsonBindLog {
    param([string]$Message, [string]$Level = "Info")
    if (-not $Script:EnableLogging -or -not $Message) { return }
    try {
        $logDir = Join-Path $ToolDir "Logs"
        if (-not (Test-Path -LiteralPath $logDir -PathType Container)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }
        $logFile = Join-Path $logDir "CrimsonBind_$(Get-Date -Format 'yyyy-MM-dd').log"
        $line = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [$Level] $Message"
        Add-Content -LiteralPath $logFile -Value $line -Encoding UTF8 -ErrorAction SilentlyContinue
    } catch { }
}

# ---------- Config.ini read/write ----------
# Map CSV key names to config.ini code part (vk = virtual key, sc = scan code). Used by ConvertTo-ConfigIniKey.
$Script:ConfigIniKeyCodeMap = @{
    "F1"="vk70"; "F2"="vk71"; "F3"="vk72"; "F4"="vk73"; "F5"="vk74"; "F6"="vk75"; "F7"="vk76"; "F8"="vk77"; "F9"="vk78"; "F10"="vk79"; "F11"="vk7a"; "F12"="vk7b"
    "0"="vk30"; "1"="vk31"; "2"="vk32"; "3"="vk33"; "4"="vk34"; "5"="vk35"; "6"="vk36"; "7"="vk37"; "8"="vk38"; "9"="vk39"
    "MINUS"="vkbd"; "EQUALS"="vkbb"
    "A"="sc1e"; "B"="sc30"; "C"="sc2e"; "D"="sc20"; "E"="sc12"; "F"="sc21"; "G"="sc22"; "H"="sc23"; "I"="sc17"; "J"="sc24"; "K"="sc25"; "L"="sc26"; "M"="sc32"; "N"="sc31"; "O"="sc18"; "P"="sc19"; "Q"="sc10"; "R"="sc13"; "S"="sc1f"; "T"="sc14"; "U"="sc16"; "V"="sc2f"; "W"="sc11"; "X"="sc2d"; "Y"="sc15"; "Z"="sc2c"
    "LBRACKET"="vkdb"; "RBRACKET"="vkdd"; "BACKSLASH"="vkdc"; "SEMICOLON"="vkba"; "APOSTROPHE"="vkde"; "COMMA"="vkbc"; "PERIOD"="vkbe"; "SLASH"="vkbf"
    "NUMPAD0"="vk60"; "NUMPAD1"="sc4f"; "NUMPAD2"="sc50"; "NUMPAD3"="sc51"; "NUMPAD4"="vk64"; "NUMPAD5"="vk65"; "NUMPAD6"="vk66"; "NUMPAD7"="sc47"; "NUMPAD8"="vk68"; "NUMPAD9"="sc49"
    "TAB"="sc0f"
}
# Reverse map: config.ini code (e.g. sc21, vk70) -> human key name (F, F1). Used by ConvertFrom-ConfigIniKey for export.
$Script:ConfigIniCodeToKeyName = @{}
foreach ($entry in $Script:ConfigIniKeyCodeMap.GetEnumerator()) { $Script:ConfigIniCodeToKeyName[$entry.Value] = $entry.Key }

# Convert config.ini key (e.g. +sc21_67701769) to human CSV key (e.g. SHIFT-F). Returns original string if not recognized or suffix differs from expected (so we don't change keys we can't round-trip).
function ConvertFrom-ConfigIniKey {
    param([string]$ConfigKey)
    $k = ($ConfigKey -replace "^\s+|\s+$", "")
    if (-not $k) { return $k }
    if ($k -notmatch '^([\^!+]*)(vk[0-9a-fA-F]+|sc[0-9a-fA-F]+)_([0-9a-fA-F]+)$') { return $k }
    $suffix = $Matches[3]
    if ($suffix -ne $Script:ConfigIniKeySuffix) { return $k }
    $modPrefix = $Matches[1]
    $code = $Matches[2].ToLowerInvariant()
    $keyName = $Script:ConfigIniCodeToKeyName[$code]
    if (-not $keyName -and $code.Length -gt 2) {
        $prefix = $code.Substring(0, 2)
        $hex = $code.Substring(2)
        if ($hex.Length -eq 1 -and ($prefix -eq "vk" -or $prefix -eq "sc")) { $code = $prefix + "0" + $hex; $keyName = $Script:ConfigIniCodeToKeyName[$code] }
    }
    if (-not $keyName) { return $k }
    $modStr = ""
    if ($modPrefix -match '\^') { $modStr += "CTRL-" }
    if ($modPrefix -match '!') { $modStr += "ALT-" }
    if ($modPrefix -match '\+') { $modStr += "SHIFT-" }
    return $modStr + $keyName
}

# Convert CSV key (e.g. F1, SHIFT-F1, ALT-K) to config.ini format (e.g. +vk70_67701769). Pass through if already in config format.
function ConvertTo-ConfigIniKey {
    param([string]$Key, [string]$Suffix = $Script:ConfigIniKeySuffix)
    $k = ($Key -replace "^\s+|\s+$", "")
    if (-not $k) { return $k }
    if ($k -match '_[0-9a-fA-F]+$' -and ($k -match 'vk[0-9a-fA-F]+' -or $k -match 'sc[0-9a-fA-F]+')) { return $k }
    $parts = $k -split '-'
    $hasCtrl = $false; $hasAlt = $false; $hasShift = $false
    $base = $null
    foreach ($p in $parts) {
        $t = $p.Trim().ToUpperInvariant()
        if ($t -eq 'CTRL') { $hasCtrl = $true } elseif ($t -eq 'ALT') { $hasAlt = $true } elseif ($t -eq 'SHIFT') { $hasShift = $true } else { $base = $p.Trim(); break }
    }
    if (-not $base -and $parts.Count -gt 0) { $base = $parts[-1].Trim() }
    if (-not $base) { return $k }
    $code = $Script:ConfigIniKeyCodeMap[$base]
    if (-not $code) { $code = $Script:ConfigIniKeyCodeMap[$base.ToUpperInvariant()] }
    if (-not $code) { return $k }
    $modStr = ""; if ($hasCtrl) { $modStr += "^" }; if ($hasAlt) { $modStr += "!" }; if ($hasShift) { $modStr += "+" }
    return $modStr + $code + "_" + $Suffix
}

# Section names in config.ini that are not keybind sections (export skips these).
$Script:ConfigIniNonBindSections = @{ "Main" = $true; "SourceSettings" = $true; "PhysInput" = $true }

# Regex: keybind value is optional modifiers then vk/sc + hex + _ + suffix. Used to exclude Main/SourceSettings/PhysInput rows even if section name matching fails.
$Script:ConfigIniKeybindValuePattern = '^[\^!+]*(vk|sc)[0-9a-fA-F]+_[0-9a-fA-F]+$'

# Action names that are universal (General-only); exclude from config.ini import so CSV General rows are preserved.
$Script:ConfigIniExcludeActionNames = @{
    "Human Racial" = $true; "Stoneform" = $true; "Shadowmeld" = $true; "Escape Artist" = $true
    "Gift of the Naaru" = $true; "Darkflight" = $true; "Blood Fury" = $true; "Will of the Forsaken" = $true
    "War Stomp" = $true; "Berserking" = $true; "Arcane Torrent" = $true; "Rocket Jump" = $true
    "Rocket Barrage" = $true; "Quaking Palm" = $true; "Spatial Rift" = $true; "Light's Judgment" = $true
    "Fireblood" = $true; "Arcane Pulse" = $true; "Bull Rush" = $true; "Ancestral Call" = $true
    "Haymaker" = $true; "Regeneratin" = $true; "Bag of Tricks" = $true; "Hyper Organic Light Originator" = $true
    "Azerite Surge" = $true; "Sharpen Blade" = $true
}

# Skip macro text that is template/description so we don't overwrite custom macros.
$Script:ConfigIniDescriptionMacroPattern = '^For instructions on how to set up this hotkey'

function Get-ConfigIniSectionsAndBinds {
    param([string]$ConfigPath)
    $out = [System.Collections.ArrayList]::new()
    if (-not (Test-Path -LiteralPath $ConfigPath -PathType Leaf)) { return $out }
    $content = [System.IO.File]::ReadAllText($ConfigPath, [System.Text.Encoding]::Unicode)
    $lines = $content -split "`r?`n"
    $currentSection = $null
    foreach ($line in $lines) {
        $strip = $line.Trim()
        if ($strip -match '^\[(.+)\]$') {
            $raw = $Matches[1].Trim()
            $currentSection = if ($raw.Length -gt 0 -and [int][char]$raw[0] -eq 0xFEFF) { $raw.Substring(1).Trim() } else { $raw }
            continue
        }
        if ($null -eq $currentSection) { continue }
        if ($Script:ConfigIniNonBindSections.ContainsKey($currentSection)) { continue }
        if ($strip -notmatch '^([^=]+)=(.*)$') { continue }
        $actionName = $Matches[1].Trim()
        $value = $Matches[2].Trim()
        if (-not $actionName) { continue }
        if ($actionName -match '^START') { continue }
        if ($actionName -match '^Target Member\d+$') { continue }
        if ($Script:ConfigIniExcludeActionNames.ContainsKey($actionName)) { continue }
        if ($actionName -match '^Universal\d+ Unit\d+$') { continue }
        $keyPart = $value
        $afterSemi = ""
        if ($value -match '^([^;]*);(.*)$') {
            $keyPart = $Matches[1].Trim()
            $afterSemi = $Matches[2].Trim()
        }
        if ($afterSemi -and $afterSemi -match $Script:ConfigIniDescriptionMacroPattern) { continue }
        if ($keyPart -and $keyPart -notmatch $Script:ConfigIniKeybindValuePattern -and -not $afterSemi) { continue }
        [void]$out.Add([PSCustomObject]@{
            Section    = $currentSection
            ActionName = $actionName
            Key        = $keyPart
            MacroText  = $afterSemi
            TextureID  = "132089"
        })
    }
    return $out
}

function Set-ConfigIniFromRows {
    param([string]$ConfigPath, [System.Collections.ArrayList]$Rows, [string]$OnlySection = $null)
    if (-not (Test-Path -LiteralPath $ConfigPath -PathType Leaf)) { return -1 }
    $generalByActionName = @{}
    foreach ($r in $Rows) {
        if ($r.Section -eq "General") { $generalByActionName[$r.ActionName] = [PSCustomObject]@{ Key = $r.Key; MacroText = $r.MacroText } }
    }
    $keyMap = @{}
    foreach ($r in $Rows) {
        $k = "$($r.Section)|$($r.ActionName)"
        if ($r.Section -eq "General") {
            $keyMap[$k] = [PSCustomObject]@{ Key = $r.Key; MacroText = $r.MacroText }
        } elseif ($generalByActionName.ContainsKey($r.ActionName)) {
            $keyMap[$k] = $generalByActionName[$r.ActionName]
        } else {
            $keyMap[$k] = [PSCustomObject]@{ Key = $r.Key; MacroText = $r.MacroText }
        }
    }
    # When updating only one section, ensure General's keys apply for every General action in that section
    # (so target binds etc. are taken from General even if the CSV has no class row for that action).
    if ($OnlySection -and $OnlySection -ne "General") {
        foreach ($actionName in $generalByActionName.Keys) {
            $keyMap["$OnlySection|$actionName"] = $generalByActionName[$actionName]
        }
    }
    $content = [System.IO.File]::ReadAllText($ConfigPath, [System.Text.Encoding]::Unicode)
    $lines = $content -split "`r?`n"
    $currentSection = $null
    $newLines = [System.Collections.ArrayList]::new()
    $replaced = 0
    $added = 0
    # Track which Section|Action we've output (replaced or passed through) so we can add missing CSV entries.
    $writtenInSection = @{}

    function Add-MissingKeysForSection {
        param($section, $keyMap, [System.Collections.ArrayList]$targetNewLines)
        $inTarget = -not $OnlySection -or ($section -eq $OnlySection)
        if (-not $inTarget) { return 0 }
        $count = 0
        foreach ($mapKey in $keyMap.Keys) {
            if ($mapKey -notmatch '^([^|]+)\|(.+)$') { continue }
            $sec = $Matches[1]
            $act = $Matches[2]
            if ($sec -ne $section) { continue }
            if ($writtenInSection[$section] -and $writtenInSection[$section][$act]) { continue }
            $info = $keyMap[$mapKey]
            $configKey = ConvertTo-ConfigIniKey -Key $info.Key -Suffix $Script:ConfigIniKeySuffix
            $newVal = $configKey
            if ($info.MacroText) { $newVal += "; $($info.MacroText)" }
            [void]$targetNewLines.Add("$act=$newVal")
            $count++
        }
        return $count
    }

    foreach ($line in $lines) {
        $strip = $line.TrimEnd("`r", "`n")
        if ($strip -match '^\[(.+)\]$') {
            $nextSection = $Matches[1].Trim()
            # Before starting the new section, flush any CSV entries that were in keyMap for the previous section but missing from config.
            if ($null -ne $currentSection) {
                $n = Add-MissingKeysForSection -section $currentSection -keyMap $keyMap -targetNewLines $newLines
                if ($n -gt 0) { $added += $n }
            }
            $currentSection = $nextSection
            if (-not $writtenInSection[$currentSection]) { $writtenInSection[$currentSection] = @{} }
            [void]$newLines.Add($line)
            continue
        }
        if ($null -ne $currentSection -and $strip -match '^([^=]+)=(.*)$') {
            $actionName = $Matches[1].Trim()
            if ($actionName) { $writtenInSection[$currentSection][$actionName] = $true }
            $key = "$currentSection|$actionName"
            $inTargetSection = -not $OnlySection -or ($currentSection -eq $OnlySection)
            if ($inTargetSection -and $actionName -and $keyMap.ContainsKey($key)) {
                $info = $keyMap[$key]
                $configKey = ConvertTo-ConfigIniKey -Key $info.Key -Suffix $Script:ConfigIniKeySuffix
                $newVal = $configKey
                if ($info.MacroText) { $newVal += "; $($info.MacroText)" }
                [void]$newLines.Add($Matches[1] + "=" + $newVal)
                $replaced++
            } else { [void]$newLines.Add($line) }
        } else { [void]$newLines.Add($line) }
    }
    # Flush missing for the last section
    if ($null -ne $currentSection) {
        $n = Add-MissingKeysForSection -section $currentSection -keyMap $keyMap -targetNewLines $newLines
        if ($n -gt 0) { $added += $n }
    }

    # Add sections that exist in CSV/keyMap but not in config.ini (so e.g. Hammer of Wrath gets written)
    $sectionsInKeyMap = @{}
    foreach ($mapKey in $keyMap.Keys) {
        if ($mapKey -match '^([^|]+)\|') { $sectionsInKeyMap[$Matches[1]] = $true }
    }
    foreach ($sec in $sectionsInKeyMap.Keys) {
        if ($writtenInSection.ContainsKey($sec)) { continue }
        if ($OnlySection -and $sec -ne $OnlySection) { continue }
        [void]$newLines.Add("")
        [void]$newLines.Add("[$sec]")
        $n = Add-MissingKeysForSection -section $sec -keyMap $keyMap -targetNewLines $newLines
        if ($n -gt 0) { $added += $n }
    }

    $outText = ($newLines -join "`r`n") + "`r`n"
    [System.IO.File]::WriteAllText($ConfigPath, $outText, [System.Text.Encoding]::Unicode)
    return ($replaced + $added)
}

# After writing config.ini, validate that binds match. Returns hashtable: Match, Mismatch, Skipped, NoKeyInCsv, TotalInScope, DuplicateKeyInSection.
function Get-ConfigIniUpdateValidation {
    param([string]$ConfigPath, [System.Collections.ArrayList]$Rows, [string]$OnlySection = $null)
    $result = @{ Match = 0; Mismatch = 0; Skipped = 0; NoKeyInCsv = 0; TotalInScope = 0; MismatchDetails = [System.Collections.ArrayList]::new(); DuplicateKeyInSection = [System.Collections.ArrayList]::new() }
    if (-not (Test-Path -LiteralPath $ConfigPath -PathType Leaf)) { return $result }
    $generalByActionName = @{}
    foreach ($r in $Rows) {
        if ($r.Section -eq "General") { $generalByActionName[$r.ActionName] = [PSCustomObject]@{ Key = $r.Key; MacroText = $r.MacroText } }
    }
    $keyMap = @{}
    foreach ($r in $Rows) {
        $k = "$($r.Section)|$($r.ActionName)"
        if ($r.Section -eq "General") {
            $keyMap[$k] = [PSCustomObject]@{ Key = $r.Key; MacroText = $r.MacroText }
        } elseif ($generalByActionName.ContainsKey($r.ActionName)) {
            $keyMap[$k] = $generalByActionName[$r.ActionName]
        } else {
            $keyMap[$k] = [PSCustomObject]@{ Key = $r.Key; MacroText = $r.MacroText }
        }
    }
    if ($OnlySection -and $OnlySection -ne "General") {
        foreach ($actionName in $generalByActionName.Keys) {
            $keyMap["$OnlySection|$actionName"] = $generalByActionName[$actionName]
        }
    }
    $configRows = Get-ConfigIniSectionsAndBinds -ConfigPath $ConfigPath
    $actualByKey = @{}
    foreach ($c in $configRows) { $actualByKey["$($c.Section)|$($c.ActionName)"] = $c }
    foreach ($r in $Rows) {
        $inScope = -not $OnlySection -or ($r.Section -eq $OnlySection)
        if (-not $inScope) { continue }
        $result.TotalInScope++
        $key = "$($r.Section)|$($r.ActionName)"
        if (-not $keyMap.ContainsKey($key)) { continue }
        $info = $keyMap[$key]
        if (-not $info.Key -or ($info.Key -eq "")) { $result.NoKeyInCsv++; continue }
        $expectedConfigKey = ConvertTo-ConfigIniKey -Key $info.Key -Suffix $Script:ConfigIniKeySuffix
        $expectedMacro = if ($info.MacroText) { $info.MacroText.Trim() } else { "" }
        if (-not $actualByKey.ContainsKey($key)) { $result.Skipped++; continue }
        $actual = $actualByKey[$key]
        $actualKey = if ($actual.Key) { $actual.Key.Trim() } else { "" }
        $actualMacro = if ($actual.MacroText) { $actual.MacroText.Trim() } else { "" }
        if ($actualKey -eq $expectedConfigKey -and $actualMacro -eq $expectedMacro) { $result.Match++ }
        else {
            $result.Mismatch++
            [void]$result.MismatchDetails.Add("$($r.Section) | $($r.ActionName): key or macro differs")
        }
    }
    # Detect duplicate key in same section (only one action can have the key in-game; report so user can change one in CSV)
    $bySectionAndKey = @{}
    foreach ($mapKey in $keyMap.Keys) {
        if ($mapKey -notmatch '^([^|]+)\|(.+)$') { continue }
        $sec = $Matches[1]; $act = $Matches[2]
        if ($OnlySection -and $sec -ne $OnlySection) { continue }
        $info = $keyMap[$mapKey]
        if (-not $info.Key -or ($info.Key -eq "")) { continue }
        $configKey = ConvertTo-ConfigIniKey -Key $info.Key -Suffix $Script:ConfigIniKeySuffix
        $k = "$sec|$configKey"
        if (-not $bySectionAndKey[$k]) { $bySectionAndKey[$k] = [System.Collections.ArrayList]::new() }
        [void]$bySectionAndKey[$k].Add($act)
    }
    foreach ($k in $bySectionAndKey.Keys) {
        $list = $bySectionAndKey[$k]
        if ($list.Count -gt 1) {
            $parts = $k -split '\|', 2
            $sec = $parts[0]
            $csvKey = $keyMap["$sec|$($list[0])"].Key
            [void]$result.DuplicateKeyInSection.Add("$sec`: $csvKey -> $($list -join ', ')")
        }
    }
    return $result
}

# ---------- CSV ----------
function Export-RowsToCsv {
    param([string]$Path, [System.Collections.ArrayList]$Rows)
    $sb = [System.Text.StringBuilder]::new()
    [void]$sb.AppendLine("Section,ActionName,MacroText,Key,TextureID")
    foreach ($r in $Rows) {
        $mac = ($r.MacroText -replace '"', '""')
        $line = "`"$($r.Section)`",`"$($r.ActionName)`",`"$mac`",`"$($r.Key)`",`"$($r.TextureID)`""
        [void]$sb.AppendLine($line)
    }
    [System.IO.File]::WriteAllText($Path, $sb.ToString(), [System.Text.Encoding]::UTF8)
}

function Import-CsvToRows {
    param([string]$Path)
    $rows = [System.Collections.ArrayList]::new()
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { return $rows }
    $csv = Import-Csv -LiteralPath $Path -Encoding UTF8
    foreach ($r in $csv) {
        # Normalize property access (Excel may save UTF-8 BOM; first column can be ﻿Section)
        $bom = [char]0xFEFF
        $sec = $null; $act = $null; $key = $null; $mac = $null; $tex = $null
        foreach ($p in $r.PSObject.Properties) {
            $name = $p.Name.TrimStart($bom)
            switch -Regex ($name) { '^Section$' { $sec = $p.Value } '^ActionName$' { $act = $p.Value } '^BindPadKey$' { $key = $p.Value } '^Key$' { if (-not $key) { $key = $p.Value } } '^MacroText$' { $mac = $p.Value } '^TextureID$' { $tex = $p.Value } }
        }
        $sec = if ($sec) { $sec.ToString().Trim() } else { "" }
        $act = if ($act) { $act.ToString().Trim() } else { "" }
        $key = if ($key) { $key.ToString().Trim() } else { "" }
        $mac = if ($mac) { $mac.ToString().Trim() } else { "" }
        $tex = if ($tex) { $tex.ToString().Trim() } else { "132089" }
        if ($sec -and $act) {
            if ($key) { $key = ConvertFrom-ConfigIniKey -ConfigKey $key }
            [void]$rows.Add([PSCustomObject]@{ Section = $sec; ActionName = $act; MacroText = $mac; Key = $key; TextureID = $tex })
        }
    }
    return $rows
}

# ---------- Randomize key pool (same idea as assign_bindpad_keys.py) ----------
$Script:KeyPool = @(
    "6","7","8","9","0","MINUS","EQUALS",
    "R","T","Y","U","I","O","P","LBRACKET","RBRACKET","BACKSLASH",
    "F","G","H","J","K","L","SEMICOLON","APOSTROPHE",
    "Z","X","C","V","B","N","M","COMMA","PERIOD","SLASH",
    "A","S","D","E","W","Q","TAB"
)
$Script:Mods = @("CTRL-","ALT-","SHIFT-","CTRL-ALT-","CTRL-SHIFT-","ALT-SHIFT-","CTRL-ALT-SHIFT-")

function Get-FullKeyPool {
    $pool = [System.Collections.ArrayList]::new()
    foreach ($m in $Script:Mods) {
        foreach ($k in $Script:KeyPool) {
            $combo = $m + $k
            if ($combo -ne "CTRL-C" -and $combo -ne "CTRL-V") { [void]$pool.Add($combo) }
        }
    }
    return $pool
}

function Get-ExcludedKeys {
    $path = $Script:ExcludedKeysPath
    if (-not $path -or -not (Test-Path -LiteralPath $path -PathType Leaf)) { return @() }
    try {
        $csv = @(Import-Csv -LiteralPath $path -Encoding UTF8 -ErrorAction Stop)
        if ($csv.Count -eq 0) { return @() }
        $first = $csv[0]
        $col = if ($first.PSObject.Properties['Key']) { 'Key' } else { ($first.PSObject.Properties.Name | Select-Object -First 1) }
        if (-not $col) { return @() }
        return @($csv | ForEach-Object { $_.$col.Trim() } | Where-Object { $_ })
    } catch { return @() }
}

function Get-FullKeyPoolForRandomize {
    $full = Get-FullKeyPool
    $excluded = Get-ExcludedKeys
    if ($excluded.Count -eq 0) { return $full }
    $exclSet = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($e in $excluded) { [void]$exclSet.Add($e) }
    $pool = [System.Collections.ArrayList]::new()
    foreach ($k in $full) { if (-not $exclSet.Contains($k)) { [void]$pool.Add($k) } }
    return $pool
}

function Invoke-RandomizeKeys {
    param([System.Collections.ArrayList]$Rows, [string]$OnlySection = $null)
    $fullPool = Get-FullKeyPoolForRandomize
    $rng = [System.Random]::new()

    # Keys currently used in General (so we don't assign them to specs when randomizing)
    $generalKeys = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($r in $Rows) {
        if ($r.Section -eq "General") {
            $k = if ($r.Key) { $r.Key.ToString().Trim() } else { "" }
            if ($k) { [void]$generalKeys.Add($k) }
        }
    }
    $poolWithoutGeneral = [System.Collections.ArrayList]::new()
    foreach ($k in $fullPool) {
        if (-not $generalKeys.Contains($k)) { [void]$poolWithoutGeneral.Add($k) }
    }

    if ($OnlySection) {
        # Randomize only the selected section; unique keys within section and do not use keys bound in General
        $targetRows = [System.Collections.ArrayList]::new()
        foreach ($r in $Rows) { if ($r.Section -eq $OnlySection) { [void]$targetRows.Add($r) } }
        $pool = if ($OnlySection -eq "General") { $fullPool } else { $poolWithoutGeneral }
        $shuffled = [System.Collections.ArrayList]::new()
        foreach ($k in ($pool | Sort-Object { $rng.Next() })) { [void]$shuffled.Add($k) }
        $assignCount = [Math]::Min($shuffled.Count, $targetRows.Count)
        for ($i = 0; $i -lt $assignCount; $i++) {
            $targetRows[$i].Key = $shuffled[$i]
        }
        return
    }

    # Randomize all sections: General first from full pool; each spec from pool minus General's keys (no duplicates within any section)
    $generalList = [System.Collections.ArrayList]::new()
    $bySection = @{}
    foreach ($r in $Rows) {
        if ($r.Section -eq "General") { [void]$generalList.Add($r) }
        else {
            if (-not $bySection[$r.Section]) { $bySection[$r.Section] = [System.Collections.ArrayList]::new() }
            [void]$bySection[$r.Section].Add($r)
        }
    }
    $assignSection = {
        param([System.Collections.ArrayList]$sectionRows, [System.Collections.ArrayList]$pool)
        $shuffled = [System.Collections.ArrayList]::new()
        foreach ($k in ($pool | Sort-Object { $rng.Next() })) { [void]$shuffled.Add($k) }
        $n = [Math]::Min($shuffled.Count, $sectionRows.Count)
        for ($i = 0; $i -lt $n; $i++) { $sectionRows[$i].Key = $shuffled[$i] }
    }
    $assignSection.Invoke($generalList, $fullPool) | Out-Null
    # Rebuild pool excluding General's newly assigned keys so specs don't reuse them
    $generalKeys.Clear()
    foreach ($r in $generalList) {
        $k = if ($r.Key) { $r.Key.ToString().Trim() } else { "" }
        if ($k) { [void]$generalKeys.Add($k) }
    }
    $poolWithoutGeneral = [System.Collections.ArrayList]::new()
    foreach ($k in $fullPool) {
        if (-not $generalKeys.Contains($k)) { [void]$poolWithoutGeneral.Add($k) }
    }
    foreach ($sec in $bySection.Keys) {
        $assignSection.Invoke($bySection[$sec], $poolWithoutGeneral) | Out-Null
    }
}

# Randomize only the currently selected grid rows: assign each to an unassigned key (no duplicate within section, exclude General keys for specs).
function Invoke-RandomizeSelectedGridRows {
    param($Grid, [System.Collections.ArrayList]$Rows, [string]$FilterSection)
    if (-not $Rows -or $Rows.Count -eq 0) { return 0 }
    $selected = @($Grid.SelectedRows)
    if ($selected.Count -eq 0) { return 0 }
    $generalActionNames = @{}
    $generalKeys = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($r in $Rows) {
        if ($r.Section -eq "General") {
            $generalActionNames[$r.ActionName] = $true
            $k = if ($r.Key) { $r.Key.ToString().Trim() } else { "" }
            if ($k) { [void]$generalKeys.Add($k) }
        }
    }
    $selectedPairs = [System.Collections.ArrayList]::new()
    foreach ($row in $selected) {
        $sec = $row.Cells["Section"].Value; $act = $row.Cells["Action"].Value
        $secStr = if ($null -ne $sec) { $sec.ToString().Trim() } else { "" }
        $actStr = if ($null -ne $act) { $act.ToString().Trim() } else { "" }
        if (-not $secStr -or -not $actStr) { continue }
        if ($secStr -ne "General" -and $generalActionNames.ContainsKey($actStr)) { continue }
        [void]$selectedPairs.Add([PSCustomObject]@{ Section = $secStr; ActionName = $actStr })
    }
    if ($selectedPairs.Count -eq 0) { return 0 }
    $fullPool = Get-FullKeyPoolForRandomize
    $poolWithoutGeneral = [System.Collections.ArrayList]::new()
    foreach ($k in $fullPool) {
        if (-not $generalKeys.Contains($k)) { [void]$poolWithoutGeneral.Add($k) }
    }
    $rng = [System.Random]::new()
    $selectedSet = @{}
    foreach ($p in $selectedPairs) { $selectedSet["$($p.Section)|$($p.ActionName)"] = $true }
    $usedBySection = @{}
    foreach ($r in $Rows) {
        $key = "$($r.Section)|$($r.ActionName)"
        if ($selectedSet[$key]) { continue }
        $k = if ($r.Key) { $r.Key.ToString().Trim() } else { "" }
        if (-not $k) { continue }
        $sec = $r.Section
        if (-not $usedBySection[$sec]) { $usedBySection[$sec] = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase) }
        [void]$usedBySection[$sec].Add($k)
    }
    $bySection = @{}
    foreach ($p in $selectedPairs) {
        $sec = $p.Section
        if (-not $bySection[$sec]) { $bySection[$sec] = [System.Collections.ArrayList]::new() }
        [void]$bySection[$sec].Add($p)
    }
    $assigned = 0
    foreach ($sec in $bySection.Keys) {
        $used = $usedBySection[$sec]
        $pool = [System.Collections.ArrayList]::new()
        $sourcePool = if ($sec -eq "General") { $fullPool } else { $poolWithoutGeneral }
        foreach ($k in $sourcePool) {
            if (-not $used -or -not $used.Contains($k)) { [void]$pool.Add($k) }
        }
        $list = $bySection[$sec]
        $shuffled = [System.Collections.ArrayList]::new()
        foreach ($k in ($pool | Sort-Object { $rng.Next() })) { [void]$shuffled.Add($k) }
        $n = [Math]::Min($shuffled.Count, $list.Count)
        for ($i = 0; $i -lt $n; $i++) {
            $p = $list[$i]
            $newKey = $shuffled[$i]
            foreach ($r in $Rows) {
                if ($r.Section -eq $p.Section -and $r.ActionName -eq $p.ActionName) {
                    $r.Key = $newKey
                    $assigned++
                    break
                }
            }
        }
    }
    return $assigned
}

# ---------- Lua helpers (Get-LuaEscaped used by Debounce export) ----------
function Get-LuaEscaped {
    param([string]$s)
    if ($null -eq $s) { return '""' }
    $s = $s -replace '\\','\\\\' -replace '"','\"' -replace "`r`n",'\n' -replace "`n",'\n' -replace "`r",'\n'
    return "`"$s`""
}

# Legacy: not used by tool (CrimsonBind uses Debounce only). Kept for compatibility.
function Write-CrimsonAutoBindDBFromRows {
    param([string]$Path, [System.Collections.ArrayList]$Rows)
    $specs = @{}
    foreach ($r in $Rows) {
        if ($r.Section -eq "General") { continue }
        if ($r.ActionName -match '^START\s') { continue }
        $specKey = $r.Section
        if (-not $specs[$specKey]) {
            $specs[$specKey] = @{ general = [System.Collections.ArrayList]::new(); spec = [System.Collections.ArrayList]::new() }
        }
        [void]$specs[$specKey].spec.Add($r)
    }
    $generalRows = [System.Collections.ArrayList]::new()
    foreach ($r in $Rows) {
        if ($r.Section -eq "General" -and $r.ActionName -notmatch '^START\s') { [void]$generalRows.Add($r) }
    }
    $sb = [System.Text.StringBuilder]::new()
    [void]$sb.AppendLine("-- Generated by CrimsonBind (legacy CrimsonAutoBindDB format; tool now uses Debounce).")
    [void]$sb.AppendLine("CrimsonAutoBindDB = CrimsonAutoBindDB or {}")
    [void]$sb.AppendLine("CrimsonAutoBindDB.savedBinds = CrimsonAutoBindDB.savedBinds or {}")
    foreach ($specKey in ($specs.Keys | Sort-Object)) {
        $data = $specs[$specKey]
        $shortName = $specKey
        if ($specKey -match ' - (.+)$') { $shortName = $Matches[1] }
        [void]$sb.AppendLine("CrimsonAutoBindDB.savedBinds[" + (Get-LuaEscaped $shortName) + "] = {")
        [void]$sb.AppendLine("  general = {")
        foreach ($g in $generalRows) {
            [void]$sb.AppendLine("    { name = " + (Get-LuaEscaped $g.ActionName) + ", key = " + (Get-LuaEscaped $g.Key) + ", macrotext = " + (Get-LuaEscaped $g.MacroText) + " },")
        }
        [void]$sb.AppendLine("  },")
        [void]$sb.AppendLine("  spec = {")
        foreach ($s in $data.spec) {
            [void]$sb.AppendLine("    { name = " + (Get-LuaEscaped $s.ActionName) + ", key = " + (Get-LuaEscaped $s.Key) + ", macrotext = " + (Get-LuaEscaped $s.MacroText) + " },")
        }
        [void]$sb.AppendLine("  },")
        [void]$sb.AppendLine("}")
    }
    $dir = [System.IO.Path]::GetDirectoryName($Path)
    if ($dir -and -not (Test-Path -LiteralPath $dir -PathType Container)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    [System.IO.File]::WriteAllText($Path, $sb.ToString(), [System.Text.Encoding]::UTF8)
}

# ---------- WTF / accounts ----------
# From config.ini [Main] WTFPath= (path can be to Config.wtf or WTF folder). Returns WTF directory.
function Get-WtfDirFromConfig {
    param([string]$ConfigPath)
    if (-not (Test-Path -LiteralPath $ConfigPath -PathType Leaf)) { return $null }
    try {
        $content = Get-Content -LiteralPath $ConfigPath -Encoding Unicode -ErrorAction Stop
    } catch { return $null }
    $inMain = $false
    foreach ($line in $content) {
        $t = $line.Trim()
        if ($t -eq "[Main]") { $inMain = $true; continue }
        if ($inMain -and $t -match '^\[') { break }
        if ($inMain -and $t -match '^WTFPath=(.+)') {
            $wtf = $Matches[1].Trim()
            $full = [System.IO.Path]::GetFullPath($wtf)
            if ([System.IO.Path]::GetExtension($full) -eq '.wtf') {
                return [System.IO.Path]::GetDirectoryName($full)
            }
            return $full
        }
    }
    return $null
}

# Path to WoW Config.wtf from config.ini (WTFPath value). Used to read account name.
function Get-ConfigWtfPathFromIni {
    param([string]$ConfigPath)
    if (-not (Test-Path -LiteralPath $ConfigPath -PathType Leaf)) { return $null }
    try {
        $content = Get-Content -LiteralPath $ConfigPath -Encoding Unicode -ErrorAction Stop
    } catch { return $null }
    $inMain = $false
    foreach ($line in $content) {
        $t = $line.Trim()
        if ($t -eq "[Main]") { $inMain = $true; continue }
        if ($inMain -and $t -match '^\[') { break }
        if ($inMain -and $t -match '^WTFPath=(.+)') {
            return [System.IO.Path]::GetFullPath($Matches[1].Trim())
        }
    }
    return $null
}

# Read account name from WoW Config.wtf (SET accountName "NAME" when "Remember Account Name" is on).
# Do not modify Config.wtf; only read. If account name is not set there, leave the file unchanged.
function Get-AccountNameFromConfigWtf {
    param([string]$ConfigWtfPath)
    if (-not $ConfigWtfPath -or -not (Test-Path -LiteralPath $ConfigWtfPath -PathType Leaf)) { return $null }
    try {
        $content = Get-Content -LiteralPath $ConfigWtfPath -Encoding Unicode -ErrorAction Stop
    } catch { return $null }
    foreach ($line in $content) {
        if ($line -match 'SET\s+accountName\s+"([^"]*)"') {
            $name = $Matches[1].Trim()
            if ($name) { return $name }
        }
    }
    return $null
}

# ---------- Backup / Restore ----------
$BackupSavedVarNames = @("Debounce.lua", "Debounce.lua.bak")

function Get-SavedVariablesDir {
    param([string]$WtfDir, [string]$AccountName)
    if (-not $WtfDir -or -not $AccountName) { return $null }
    return Join-Path $WtfDir "Account\$AccountName\SavedVariables"
}

function New-CrimsonBackup {
    param([string]$ConfigPath, [string]$WtfDir, [string]$AccountName, [string]$CsvPath)
    if (-not $ConfigPath -or -not (Test-Path -LiteralPath $ConfigPath -PathType Leaf)) { return $null }
    if (-not (Test-Path -LiteralPath $Script:BackupRoot -PathType Container)) {
        New-Item -ItemType Directory -Path $Script:BackupRoot -Force | Out-Null
    }
    $timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
    $destDir = Join-Path $Script:BackupRoot $timestamp
    New-Item -ItemType Directory -Path $destDir -Force | Out-Null
    Copy-Item -LiteralPath $ConfigPath -Destination (Join-Path $destDir "config.ini") -Force
    $count = 1
    if ($WtfDir -and $AccountName) {
        $svDir = Get-SavedVariablesDir -WtfDir $WtfDir -AccountName $AccountName
        if ($svDir -and (Test-Path -LiteralPath $svDir -PathType Container)) {
            foreach ($name in $BackupSavedVarNames) {
                $src = Join-Path $svDir $name
                if (Test-Path -LiteralPath $src -PathType Leaf) {
                    Copy-Item -LiteralPath $src -Destination (Join-Path $destDir $name) -Force
                    $count++
                }
            }
        }
    }
    if ($CsvPath -and (Test-Path -LiteralPath $CsvPath -PathType Leaf)) {
        $csvName = [System.IO.Path]::GetFileName($CsvPath)
        Copy-Item -LiteralPath $CsvPath -Destination (Join-Path $destDir $csvName) -Force
        $count++
    }
    return @{ Path = $destDir; FileCount = $count }
}

function Restore-CrimsonBackup {
    param([string]$BackupFolderPath, [string]$ConfigPath, [string]$WtfDir, [string]$AccountName)
    if (-not $BackupFolderPath -or -not (Test-Path -LiteralPath $BackupFolderPath -PathType Container)) { return 0 }
    $count = 0
    $configDest = Join-Path $BackupFolderPath "config.ini"
    if ((Test-Path -LiteralPath $configDest -PathType Leaf) -and $ConfigPath) {
        $destDir = [System.IO.Path]::GetDirectoryName($ConfigPath)
        if (-not (Test-Path -LiteralPath $destDir -PathType Container)) { New-Item -ItemType Directory -Path $destDir -Force | Out-Null }
        Copy-Item -LiteralPath $configDest -Destination $ConfigPath -Force
        $count++
    }
    if ($WtfDir -and $AccountName) {
        $svDir = Get-SavedVariablesDir -WtfDir $WtfDir -AccountName $AccountName
        if (-not (Test-Path -LiteralPath $svDir -PathType Container)) { New-Item -ItemType Directory -Path $svDir -Force | Out-Null }
        foreach ($name in $BackupSavedVarNames) {
            $src = Join-Path $BackupFolderPath $name
            if (Test-Path -LiteralPath $src -PathType Leaf) {
                Copy-Item -LiteralPath $src -Destination (Join-Path $svDir $name) -Force
                $count++
            }
        }
    }
    return $count
}

function Get-BackupFolderList {
    if (-not (Test-Path -LiteralPath $Script:BackupRoot -PathType Container)) { return @() }
    return Get-ChildItem -LiteralPath $Script:BackupRoot -Directory | Sort-Object Name -Descending
}

function Update-BackupRestoreState {
    param($btnBackup, $comboBackup, $btnRestore)
    $hasConfig = $Script:ConfigPath -and (Test-Path -LiteralPath $Script:ConfigPath -PathType Leaf)
    $btnBackup.Enabled = $hasConfig -and $Script:WtfDir
    $btnRestore.Enabled = ($null -ne $comboBackup.SelectedItem) -and $Script:ConfigPath
}

# ---------- Debounce.lua ----------
# CSV Section -> Debounce class key and tab index. General = one tab; class specs = 1,2,3 per class.
function Get-DebounceSectionMapping {
    param([string]$Section)
    if (-not $Section) { return $null }
    $s = $Section.Trim()
    if ($s -ieq "General") { return @{ Class = "GENERAL"; Tab = 1 } }
    if ($s -match '^(.+)\s+-\s+(.+)$') {
        $classPart = $Matches[1].Trim() -replace '\s+',''
        $specPart = $Matches[2].Trim()
        $classKey = $classPart.ToUpperInvariant()
        $specOrder = @{
            "PALADIN" = @("Holy", "Protection", "Retribution")
            "ROGUE" = @("Assassination", "Outlaw", "Subtlety")
            "DEATHKNIGHT" = @("Blood", "Frost", "Unholy")
            "DEMONHUNTER" = @("Havoc", "Vengeance")
            "DRUID" = @("Balance", "Feral", "Guardian", "Restoration")
            "EVOKER" = @("Augmentation", "Devastation", "Preservation")
            "HUNTER" = @("Beast Mastery", "Marksmanship", "Survival")
            "MAGE" = @("Arcane", "Fire", "Frost")
            "MONK" = @("Brewmaster", "Mistweaver", "Windwalker")
            "PRIEST" = @("Discipline", "Holy", "Shadow")
            "SHAMAN" = @("Elemental", "Enhancement", "Restoration")
            "WARLOCK" = @("Affliction", "Demonology", "Destruction")
            "WARRIOR" = @("Arms", "Fury", "Protection")
        }
        $order = $specOrder[$classKey]
        if (-not $order) { $order = @($specPart) }
        $tab = 1
        for ($i = 0; $i -lt $order.Count; $i++) {
            if ($order[$i] -eq $specPart) { $tab = $i + 1; break }
        }
        return @{ Class = $classKey; Tab = $tab }
    }
    return $null
}

# Extract a top-level key block from Debounce.lua text (e.g. ["customStates"] = { ... }). Returns $null if not found.
function Get-DebounceBlock {
    param([string]$Text, [string]$KeyName)
    if (-not $Text) { return $null }
    $pat = '\[\s*"' + [regex]::Escape($KeyName) + '"\s*\]\s*='
    $m = [regex]::Match($Text, $pat)
    if (-not $m.Success) { return $null }
    $start = $m.Index + $m.Length
    $rest = $Text.Substring($start).TrimStart()
    if ($rest.Length -gt 0 -and $rest[0] -eq '{') {
        $depth = 0; $i = 0; $len = $rest.Length
        for (; $i -lt $len; $i++) {
            $c = $rest[$i]
            if ($c -eq '{') { $depth++ }
            elseif ($c -eq '}') { $depth--; if ($depth -eq 0) { $i++; break } }
        }
        return $rest.Substring(0, $i).Trim()
    }
    if ($rest -match '^\d+') { return $Matches[0] }
    $end = $rest.IndexOf("`n")
    if ($end -lt 0) { $end = $rest.Length }
    return $rest.Substring(0, $end).Trim()
}

# Get full line for one key=value (including key) so we can replace or preserve. Returns start index and length of value part.
function Find-DebounceKeyValue {
    param([string]$Text, [string]$KeyName)
    $pat = '\[\s*"' + [regex]::Escape($KeyName) + '"\s*\]\s*='
    $m = [regex]::Match($Text, $pat)
    if (-not $m.Success) { return $null }
    $valueStart = $m.Index + $m.Length
    $rest = $Text.Substring($valueStart).TrimStart()
    $valueLen = 0
    if ($rest -match '^\d+') { $valueLen = $rest.Length; $valueLen = ([regex]::Match($rest, '^\d+')).Length }
    elseif ($rest[0] -eq '{') {
        $depth = 0; $i = 0; $len = $rest.Length
        for (; $i -lt $len; $i++) {
            $c = $rest[$i]
            if ($c -eq '{') { $depth++ }
            elseif ($c -eq '}') { $depth--; if ($depth -eq 0) { $valueLen = $i + 1; break } }
        }
    }
    return @{ KeyStart = $m.Index; ValueStart = $valueStart; ValueLength = $valueLen }
}

# Build Lua for one bind entry (macrotext with name, value, icon, key). Key is normalized to config.ini format (+sc21_67701769) so Debounce receives same format as config.ini.
function Get-DebounceBindLua {
    param($Row)
    $name = Get-LuaEscaped -s $Row.ActionName
    $value = Get-LuaEscaped -s $Row.MacroText
    $icon = 132089
    if ($Row.TextureID -match '^\d+$') { $icon = [int]$Row.TextureID }
    $configKey = ConvertTo-ConfigIniKey -Key $Row.Key -Suffix $Script:ConfigIniKeySuffix
    $key = Get-LuaEscaped -s $configKey
    $lines = @(
        "{",
        "[""type""] = ""macrotext"",",
        "[""name""] = $name,",
        "[""value""] = $value,",
        "[""icon""] = $icon,",
        "[""key""] = $key,",
        "},"
    )
    return $lines -join "`r`n"
}

# Build one tab (array of bind entries) as Lua.
function Get-DebounceTabLua {
    param([System.Collections.ArrayList]$Rows)
    $sb = [System.Text.StringBuilder]::new()
    [void]$sb.AppendLine("{")
    foreach ($r in $Rows) {
        $bind = Get-DebounceBindLua -Row $r
        [void]$sb.AppendLine($bind)
    }
    [void]$sb.AppendLine("},")
    return $sb.ToString().TrimEnd()
}

# GENERAL is a flat list of binds (not class-style tab wrappers). Matches Debounce addon format.
function Get-DebounceGeneralSectionLua {
    param([System.Collections.ArrayList]$Rows)
    $sb = [System.Text.StringBuilder]::new()
    [void]$sb.AppendLine('["GENERAL"] = {')
    foreach ($r in $Rows) {
        $bind = Get-DebounceBindLua -Row $r
        [void]$sb.AppendLine($bind)
    }
    [void]$sb.AppendLine('},')
    return $sb.ToString().TrimEnd()
}

# Number of spec tabs per class (so tab index matches Holy/Protection/Retribution etc.).
$Script:DebounceClassTabCount = @{
    "PALADIN" = 3; "ROGUE" = 3; "DEATHKNIGHT" = 3; "DEMONHUNTER" = 2; "DRUID" = 4; "EVOKER" = 3
    "HUNTER" = 3; "MAGE" = 3; "MONK" = 3; "PRIEST" = 3; "SHAMAN" = 3; "WARLOCK" = 3; "WARRIOR" = 3
}

# Build full class table: tabs 1..N (empty tabs filled so Retribution = tab 3 etc.), optionally [0]. Do not use for GENERAL.
function Get-DebounceSectionTableLua {
    param([string]$ClassKey, [hashtable]$TabsByIndex)
    $maxTab = $Script:DebounceClassTabCount[$ClassKey]
    if (-not $maxTab) { $maxTab = 3 }
    $emptyTabLua = Get-DebounceTabLua -Rows ([System.Collections.ArrayList]::new())
    $sb = [System.Text.StringBuilder]::new()
    [void]$sb.Append("[""$ClassKey""] = {")
    $sb.AppendLine()
    for ($idx = 1; $idx -le $maxTab; $idx++) {
        $tabLua = if ($TabsByIndex[$idx]) { $TabsByIndex[$idx] } else { $emptyTabLua }
        [void]$sb.AppendLine($tabLua)
    }
    if ($TabsByIndex[0]) {
        [void]$sb.AppendLine("[0] = {")
        [void]$sb.AppendLine($TabsByIndex[0])
        [void]$sb.AppendLine("},")
    }
    [void]$sb.AppendLine("},")
    return $sb.ToString().TrimEnd()
}

# Parse existing Debounce.lua and return hashtable: PreservedBlocks (customStates, options, ui, dbver, dever), BindSections (key -> full key=value text).
function Read-DebounceFile {
    param([string]$Path)
    $result = @{ PreservedBlocks = @{}; BindSections = @{}; RawText = "" }
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { return $result }
    $text = [System.IO.File]::ReadAllText($Path, [System.Text.Encoding]::UTF8)
    $result.RawText = $text
    foreach ($key in @("customStates", "options", "ui", "dbver", "dever")) {
        $block = Get-DebounceBlock -Text $text -KeyName $key
        if ($block) { $result.PreservedBlocks[$key] = $block }
    }
    foreach ($key in @("GENERAL", "PALADIN", "ROGUE", "DEATHKNIGHT", "DEMONHUNTER", "DRUID", "EVOKER", "HUNTER", "MAGE", "MONK", "PRIEST", "SHAMAN", "WARLOCK", "WARRIOR")) {
        $block = Get-DebounceBlock -Text $text -KeyName $key
        if ($block) { $result.BindSections[$key] = $block }
    }
    return $result
}

# Build DebounceVars content from rows (all or for one section). Preserved: customStates, options, ui, dbver, dever.
function Export-DebounceFromRows {
    param(
        [string]$Path,
        [System.Collections.ArrayList]$Rows,
        [string]$OnlySection,
        [hashtable]$ExistingBindSections
    )
    if (-not $ExistingBindSections) { $ExistingBindSections = @{} }
    $preserved = @{}
    if (Test-Path -LiteralPath $Path -PathType Leaf) {
        $read = Read-DebounceFile -Path $Path
        $preserved = $read.PreservedBlocks
        foreach ($k in $read.BindSections.Keys) {
            if (-not $ExistingBindSections[$k]) { $ExistingBindSections[$k] = $read.BindSections[$k] }
        }
    }
    $defaultCustomStates = "{" + "`r`n{" + "`r`n[""value""] = false," + "`r`n[""mode""] = 0," + "`r`n}," + "`r`n{" + "`r`n[""value""] = false," + "`r`n[""mode""] = 0," + "`r`n}," + "`r`n{" + "`r`n[""value""] = false," + "`r`n[""mode""] = 0," + "`r`n}," + "`r`n{" + "`r`n[""value""] = false," + "`r`n[""mode""] = 0," + "`r`n}," + "`r`n{" + "`r`n[""value""] = false," + "`r`n[""mode""] = 0," + "`r`n}," + "`r`n}"
    $defaultOptions = "{" + "`r`n[""blizzframes""] = {" + "`r`n}," + "`r`n}"
    $defaultUi = "{" + "`r`n[""pos""] = {" + "`r`n[""y""] = 653.3333740234375," + "`r`n[""x""] = 1033.33349609375," + "`r`n}," + "`r`n}," + "`r`n}"
    if (-not $preserved["customStates"]) { $preserved["customStates"] = $defaultCustomStates }
    if (-not $preserved["options"]) { $preserved["options"] = $defaultOptions }
    if (-not $preserved["ui"]) { $preserved["ui"] = $defaultUi }
    if (-not $preserved["dbver"]) { $preserved["dbver"] = "2" }
    if (-not $preserved["dever"]) { $preserved["dever"] = "2" }

    $rowsToUse = [System.Collections.ArrayList]::new()
    if ($OnlySection) {
        foreach ($r in $Rows) { if ($r.Section -eq $OnlySection) { [void]$rowsToUse.Add($r) } }
    } else {
        foreach ($r in $Rows) { [void]$rowsToUse.Add($r) }
    }
    $generalActionNames = @{}
    foreach ($r in $Rows) {
        if ($r.Section -eq "General") { $generalActionNames[$r.ActionName] = $true }
    }
    $byClassTab = @{}
    foreach ($r in $rowsToUse) {
        $map = Get-DebounceSectionMapping -Section $r.Section
        if (-not $map) { continue }
        $ck = $map.Class
        $tab = $map.Tab
        if (-not $byClassTab[$ck]) { $byClassTab[$ck] = @{} }
        if (-not $byClassTab[$ck][$tab]) { $byClassTab[$ck][$tab] = [System.Collections.ArrayList]::new() }
        [void]$byClassTab[$ck][$tab].Add($r)
    }
    $newBindSections = @{}
    foreach ($ck in $byClassTab.Keys) {
        if ($ck -eq "GENERAL") {
            $allGeneral = [System.Collections.ArrayList]::new()
            foreach ($idx in ($byClassTab["GENERAL"].Keys | Sort-Object)) {
                foreach ($r in $byClassTab["GENERAL"][$idx]) { [void]$allGeneral.Add($r) }
            }
            $newBindSections["GENERAL"] = Get-DebounceGeneralSectionLua -Rows $allGeneral
            continue
        }
        $tabs = $byClassTab[$ck]
        $tabLuaByIndex = @{}
        foreach ($idx in $tabs.Keys) {
            $filtered = [System.Collections.ArrayList]::new()
            foreach ($r in $tabs[$idx]) {
                if (-not $generalActionNames.ContainsKey($r.ActionName)) { [void]$filtered.Add($r) }
            }
            $tabLuaByIndex[$idx] = Get-DebounceTabLua -Rows $filtered
        }
        $newBindSections[$ck] = Get-DebounceSectionTableLua -ClassKey $ck -TabsByIndex $tabLuaByIndex
    }

    $classOrder = @("PALADIN", "ROGUE", "DEATHKNIGHT", "DEMONHUNTER", "DRUID", "EVOKER", "HUNTER", "MAGE", "MONK", "PRIEST", "SHAMAN", "WARLOCK", "WARRIOR")
    $sb = [System.Text.StringBuilder]::new()
    [void]$sb.AppendLine("DebounceVars = {")
    [void]$sb.AppendLine('["customStates"] = ')
    [void]$sb.AppendLine($preserved["customStates"] + ",")
    $onlyClass = if ($OnlySection) { (Get-DebounceSectionMapping -Section $OnlySection).Class } else { $null }
    foreach ($c in $classOrder) {
        $useNew = $newBindSections[$c] -and (-not $OnlySection -or $onlyClass -eq $c)
        if ($useNew) {
            [void]$sb.AppendLine($newBindSections[$c])
        } elseif ($ExistingBindSections[$c]) {
            [void]$sb.AppendLine("[""$c""] = ")
            [void]$sb.AppendLine(($ExistingBindSections[$c].TrimEnd()) + ",")
        }
    }
    [void]$sb.AppendLine('["dbver"] = ' + $preserved["dbver"] + ',')
    $useNewGeneral = $newBindSections["GENERAL"] -and (-not $OnlySection -or $onlyClass -eq "GENERAL")
    if ($useNewGeneral) {
        [void]$sb.AppendLine($newBindSections["GENERAL"])
    } elseif ($ExistingBindSections["GENERAL"]) {
        [void]$sb.AppendLine('["GENERAL"] = ')
        [void]$sb.AppendLine(($ExistingBindSections["GENERAL"].TrimEnd()) + ",")
    }
    [void]$sb.AppendLine('["options"] = ')
    [void]$sb.AppendLine($preserved["options"] + ",")
    [void]$sb.AppendLine('["ui"] = ')
    [void]$sb.AppendLine($preserved["ui"] + ",")
    [void]$sb.AppendLine('["dever"] = ' + $preserved["dever"])
    [void]$sb.AppendLine("}")
    [System.IO.File]::WriteAllText($Path, $sb.ToString(), [System.Text.Encoding]::UTF8)
}

# ---------- GUI ----------
# Returns hashtable of "Section|Key" -> $true for keys that appear more than once in that section (empty keys ignored).
# Rows that inherit key from General (e.g. Target Member1-40 in a spec) are excluded so they don't trigger duplicate highlight.
function Get-DuplicateKeysBySection {
    param([System.Collections.ArrayList]$Rows)
    $generalActionNames = @{}
    foreach ($r in $Rows) {
        if ($r.Section -eq "General") { $generalActionNames[$r.ActionName] = $true }
    }
    $bySectionKey = @{}
    foreach ($r in $Rows) {
        if ($r.Section -ne "General" -and $generalActionNames.ContainsKey($r.ActionName)) { continue }
        $k = if ($r.Key) { $r.Key.ToString().Trim() } else { "" }
        if (-not $k) { continue }
        $sec = if ($r.Section) { $r.Section.ToString().Trim() } else { "" }
        if (-not $sec) { continue }
        $key = "$sec|$k"
        if (-not $bySectionKey[$key]) { $bySectionKey[$key] = 0 }
        $bySectionKey[$key]++
    }
    $duplicateSet = @{}
    foreach ($key in $bySectionKey.Keys) { if ($bySectionKey[$key] -gt 1) { $duplicateSet[$key] = $true } }
    return $duplicateSet
}

function Update-GridFromRows {
    param($Grid, [System.Collections.ArrayList]$Rows, [string]$FilterSection, [string]$SearchText)
    if ($null -eq $SearchText -and $Script:GridSearchText) { $SearchText = $Script:GridSearchText }
    $duplicateSet = Get-DuplicateKeysBySection -Rows $Rows
    $keyList = [System.Collections.ArrayList]::new()
    foreach ($k in (Get-FullKeyPool)) { [void]$keyList.Add($k) }
    # Include F1–F12 and common modifier combos so General target binds (F12, SHIFT-F12, CTRL-F12, etc.) always appear in the Key dropdown
    foreach ($fk in @("F1","F2","F3","F4","F5","F6","F7","F8","F9","F10","F11","F12")) {
        if ($keyList -notcontains $fk) { [void]$keyList.Add($fk) }
        foreach ($mod in @("SHIFT-","CTRL-","ALT-")) {
            $c = $mod + $fk
            if ($keyList -notcontains $c) { [void]$keyList.Add($c) }
        }
    }
    foreach ($r in $Rows) {
        $k = if ($r.Key) { $r.Key.Trim() } else { "" }
        if ($k -and $keyList -notcontains $k) { [void]$keyList.Add($k) }
    }
    $keyList.Sort()
    [void]$keyList.Insert(0, "")   # blank option at top to clear the key
    $keyCol = $Grid.Columns["Key"]
    if ($keyCol -and $keyCol -is [System.Windows.Forms.DataGridViewComboBoxColumn]) {
        $keyCol.DataSource = $null
        $keyCol.DataSource = [object[]]$keyList
    }
    $Grid.Rows.Clear()
    $generalActionNames = @{}
    $generalByActionName = @{}
    if ($FilterSection -and $FilterSection -ne "General") {
        foreach ($r in $Rows) {
            if ($r.Section -eq "General") {
                $generalActionNames[$r.ActionName] = $true
                $generalByActionName[$r.ActionName] = [PSCustomObject]@{ Key = $r.Key; MacroText = $r.MacroText }
            }
        }
    }
    $search = if ($SearchText) { $SearchText.Trim() } else { "" }
    foreach ($r in $Rows) {
        if ($FilterSection -and $r.Section -ne $FilterSection) { continue }
        $useGeneral = $FilterSection -and $FilterSection -ne "General" -and $generalByActionName.ContainsKey($r.ActionName)
        if ($useGeneral) {
            $g = $generalByActionName[$r.ActionName]
            $dispKey = if ($g.Key) { $g.Key } else { "" }
            $dispMacro = if ($g.MacroText) { $g.MacroText } else { "" }
        } else {
            $dispKey = $r.Key; $dispMacro = $r.MacroText
        }
        if ($search) {
            $s = if ($r.Section) { $r.Section } else { "" }
            $a = if ($r.ActionName) { $r.ActionName } else { "" }
            $combined = "$s $a $dispKey $dispMacro".ToUpperInvariant()
            if ($combined.IndexOf($search.ToUpperInvariant()) -lt 0) { continue }
        }
        [void]$Grid.Rows.Add($r.Section, $r.ActionName, $dispKey, $dispMacro)
    }
    # Highlight Key cell when duplicate within same section (red tint); General and cross-spec same key are not highlighted
    $normalBack = [System.Drawing.Color]::White
    $duplicateBack = [System.Drawing.Color]::FromArgb(255, 220, 220)
    foreach ($row in $Grid.Rows) {
        $sec = $row.Cells["Section"].Value; $keyVal = $row.Cells["Key"].Value
        $secStr = if ($null -ne $sec) { $sec.ToString().Trim() } else { "" }
        $keyStr = if ($null -ne $keyVal -and ($keyVal.ToString().Trim())) { $keyVal.ToString().Trim() } else { "" }
        if ($keyStr -and $duplicateSet["$secStr|$keyStr"]) {
            $row.Cells["Key"].Style.BackColor = $duplicateBack
        } else {
            $row.Cells["Key"].Style.BackColor = $normalBack
        }
    }
    $Script:CurrentFilterSection = $FilterSection
}

function Sync-GridToRows {
    param($Grid, [System.Collections.ArrayList]$Rows, [string]$FilterSection)
    $generalActionNames = @{}
    if ($FilterSection -and $FilterSection -ne "General") {
        foreach ($r in $Rows) { if ($r.Section -eq "General") { $generalActionNames[$r.ActionName] = $true } }
    }
    $key = 0
    foreach ($r in $Rows) {
        if ($FilterSection -and $r.Section -ne $FilterSection) { continue }
        $isGeneralAction = $FilterSection -and $FilterSection -ne "General" -and $generalActionNames.ContainsKey($r.ActionName)
        if ($key -lt $Grid.Rows.Count -and -not $isGeneralAction) {
            $cellKey = $Grid.Rows[$key].Cells["Key"].Value
            $cellMac = $Grid.Rows[$key].Cells["MacroText"].Value
            $r.Key = if ($null -ne $cellKey) { $cellKey.ToString().Trim() } else { "" }
            $r.MacroText = if ($null -ne $cellMac) { $cellMac.ToString().Trim() } else { "" }
        }
        $key++
    }
}

# GUI helpers: get current filter, refresh combos, set status. Used by event handlers.
function Get-CurrentFilterSection {
    param([System.Windows.Forms.ComboBox]$ComboFilter)
    if (-not $ComboFilter -or -not $ComboFilter.SelectedItem -or $ComboFilter.SelectedItem -eq "(All sections)") { return $null }
    return $ComboFilter.SelectedItem
}

function Show-DuplicateKeysPopup {
    param([System.Collections.ArrayList]$DuplicateList, [System.Windows.Forms.Form]$Owner = $null)
    if (-not $DuplicateList -or $DuplicateList.Count -eq 0) { return }
    $dupForm = New-Object System.Windows.Forms.Form
    $dupForm.Text = "Duplicate keys in same section"
    $dupForm.Size = New-Object System.Drawing.Size(600, 440)
    $dupForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
    if ($Owner) { $dupForm.Owner = $Owner }
    $dupForm.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $y = [int]10
    $lbl1 = New-Object System.Windows.Forms.Label
    $lbl1.Text = "What is a duplicate? One key (e.g. CTRL-K) is bound to two or more different actions. In-game only one of them will respond to that key."
    $lbl1.AutoSize = $true
    $lbl1.Location = New-Object System.Drawing.Point(12, $y)
    $lbl1.MaximumSize = New-Object System.Drawing.Size(560, 0)
    [void]$dupForm.Controls.Add($lbl1)
    $y += 28
    $lbl2 = New-Object System.Windows.Forms.Label
    $lbl2.Text = "How to fix: In the grid, find each action after the arrow and give one of them a different key so that key is used by only one action."
    $lbl2.AutoSize = $true
    $lbl2.Location = New-Object System.Drawing.Point(12, $y)
    $lbl2.MaximumSize = New-Object System.Drawing.Size(560, 0)
    [void]$dupForm.Controls.Add($lbl2)
    $y += 28
    $lbl3 = New-Object System.Windows.Forms.Label
    $lbl3.Text = "List (Section: Key → Action1, Action2, ...):"
    $lbl3.AutoSize = $true
    $lbl3.Location = New-Object System.Drawing.Point(12, $y)
    [void]$dupForm.Controls.Add($lbl3)
    $y += 22
    $lb = New-Object System.Windows.Forms.ListBox
    $lb.Location = New-Object System.Drawing.Point(12, $y)
    $lb.Size = New-Object System.Drawing.Size(560, 260)
    $lb.Font = New-Object System.Drawing.Font("Consolas", 9)
    foreach ($line in $DuplicateList) {
        [void]$lb.Items.Add($line)
    }
    [void]$dupForm.Controls.Add($lb)
    $lblCount = New-Object System.Windows.Forms.Label
    $lblCount.Text = "Total: $($DuplicateList.Count) duplicate key(s) in this section."
    $lblCount.AutoSize = $true
    $lblCount.Location = New-Object System.Drawing.Point(12, 354)
    [void]$dupForm.Controls.Add($lblCount)
    $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Text = "OK"
    $btnOk.Size = New-Object System.Drawing.Size(80, 26)
    $btnOk.Location = New-Object System.Drawing.Point(492, 350)
    $btnOk.Add_Click({ $dupForm.Close() })
    [void]$dupForm.Controls.Add($btnOk)
    $dupForm.AcceptButton = $btnOk
    $dupForm.ShowDialog() | Out-Null
}

function Update-SectionFilterCombo {
    param([System.Collections.ArrayList]$Rows, [System.Windows.Forms.ComboBox]$ComboFilter)
    $ComboFilter.Items.Clear()
    [void]$ComboFilter.Items.Add("(All sections)")
    if ($Rows -and $Rows.Count -gt 0) {
        $sections = $Rows | ForEach-Object { $_.Section } | Sort-Object -Unique
        foreach ($s in $sections) { [void]$ComboFilter.Items.Add($s) }
    }
    $ComboFilter.SelectedIndex = 0
}

function Update-AccountCombo {
    param([string]$ConfigPath, [System.Windows.Forms.ComboBox]$ComboAccount)
    $ComboAccount.Items.Clear()
    [void]$ComboAccount.Items.Add("(Pick account...)")
    $wtfDir = Get-WtfDirFromConfig -ConfigPath $ConfigPath
    if (-not $wtfDir) { $ComboAccount.SelectedIndex = 0; return }
    $accountDir = Join-Path $wtfDir "Account"
    if (-not (Test-Path -LiteralPath $accountDir -PathType Container)) { $ComboAccount.SelectedIndex = 0; return }
    Get-ChildItem -LiteralPath $accountDir -Directory | ForEach-Object { [void]$ComboAccount.Items.Add($_.Name) }
    $idx = 0
    $configWtfPath = Get-ConfigWtfPathFromIni -ConfigPath $ConfigPath
    $accountFromWtf = if ($configWtfPath) { Get-AccountNameFromConfigWtf -ConfigWtfPath $configWtfPath } else { $null }
    if ($accountFromWtf) {
        for ($i = 1; $i -lt $ComboAccount.Items.Count; $i++) {
            if ($ComboAccount.Items[$i] -eq $accountFromWtf) { $idx = $i; break }
        }
    } elseif ($ComboAccount.Items.Count -eq 2) { $idx = 1 }
    $ComboAccount.SelectedIndex = $idx
}

function Update-BackupCombo {
    param([System.Windows.Forms.ComboBox]$ComboBackup)
    $ComboBackup.Items.Clear()
    Get-BackupFolderList | ForEach-Object { [void]$ComboBackup.Items.Add($_.Name) }
    if ($ComboBackup.Items.Count -gt 0) { $ComboBackup.SelectedIndex = 0 }
}

# Show Exclusions dialog; saves to Script:ExcludedKeysPath and updates Script state. Uses Panel (no GroupBox) to avoid grey dropdown overlap.
function Show-ExclusionsDialog {
    param(
        [System.Windows.Forms.Label]$StatusLabel,
        [string]$ExcludedKeysPath
    )
    $allKeys = Get-FullKeyPool
    $excluded = Get-ExcludedKeys
    $exclForm = New-Object System.Windows.Forms.Form
    $exclForm.Text = "Excluded keys (unused by Randomize)"
    $exclForm.Size = New-Object System.Drawing.Size(380, 420)
    $exclForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
    $exclForm.FormBorderStyle = "Sizable"
    $exclForm.MinimumSize = New-Object System.Drawing.Size(320, 300)
    $exclForm.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $pad = 8
    $lblExcl = New-Object System.Windows.Forms.Label
    $lblExcl.Text = "Check keys to EXCLUDE from Randomize (leave unchecked to allow):"
    $lblExcl.Location = New-Object System.Drawing.Point($pad, $pad)
    $lblExcl.AutoSize = $true
    $lblExcl.MaximumSize = New-Object System.Drawing.Size(340, 0)
    [void]$exclForm.Controls.Add($lblExcl)
    $clb = New-Object System.Windows.Forms.CheckedListBox
    $clb.Location = New-Object System.Drawing.Point($pad, 36)
    $clb.Size = New-Object System.Drawing.Size(340, 300)
    $clb.Anchor = "Top,Bottom,Left,Right"
    $clb.CheckOnClick = $true
    $clb.Sorted = $true
    foreach ($k in $allKeys) { [void]$clb.Items.Add($k) }
    for ($i = 0; $i -lt $clb.Items.Count; $i++) {
        if ($excluded -contains $clb.Items[$i]) { $clb.SetItemChecked($i, $true) }
    }
    [void]$exclForm.Controls.Add($clb)
    $btnExclSave = New-Object System.Windows.Forms.Button
    $btnExclSave.Text = "Save to CSV"
    $btnExclSave.Location = New-Object System.Drawing.Point($pad, 346)
    $btnExclSave.Size = New-Object System.Drawing.Size(100, 26)
    $btnExclSave.Anchor = "Bottom,Left"
    $btnExclSave.Add_Click({
        try {
            $dir = [System.IO.Path]::GetDirectoryName($ExcludedKeysPath)
            if ($dir -and -not (Test-Path -LiteralPath $dir -PathType Container)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
            $sb = [System.Text.StringBuilder]::new()
            [void]$sb.AppendLine("Key")
            foreach ($item in $clb.CheckedItems) { [void]$sb.AppendLine($item) }
            [System.IO.File]::WriteAllText($ExcludedKeysPath, $sb.ToString(), [System.Text.Encoding]::UTF8)
            if ($StatusLabel) { $StatusLabel.Text = "Excluded keys saved to excluded_keys.csv ($($clb.CheckedItems.Count) excluded)." }
            $exclForm.Close()
        } catch {
            if ($StatusLabel) { $StatusLabel.Text = "Error saving exclusions: $_" }
        Write-CrimsonBindLog -Message "Save exclusions failed: $_" -Level "Error"
        }
    })
    [void]$exclForm.Controls.Add($btnExclSave)
    $btnExclCancel = New-Object System.Windows.Forms.Button
    $btnExclCancel.Text = "Cancel"
    $btnExclCancel.Location = New-Object System.Drawing.Point(116, 346)
    $btnExclCancel.Size = New-Object System.Drawing.Size(75, 26)
    $btnExclCancel.Anchor = "Bottom,Left"
    $btnExclCancel.Add_Click({ $exclForm.Close() })
    [void]$exclForm.Controls.Add($btnExclCancel)
    $exclForm.Add_Shown({
        $exclForm.Width = [Math]::Max(360, $exclForm.Width)
        $exclForm.Height = [Math]::Max(380, $exclForm.Height)
    })
    $exclForm.ShowDialog()
}

# ---------- GUI ----------
# Layout: TableLayoutPanel (Dock Fill) for main form; no GroupBox to avoid grey overlap on ComboBox dropdowns.
# Naming: lblConfigPath, txtConfigPath, btnBrowseConfig, comboFilter, lblStatus, etc.
# Helpers: Get-CurrentFilterSection, Update-SectionFilterCombo, Update-AccountCombo, Update-BackupCombo, Show-ExclusionsDialog.
# Event handlers are thin: get state -> call business logic -> update UI. Error handling: try/catch with status text and optional Write-CrimsonBindLog.
# Summary of changes: (1) TableLayoutPanel + FlowLayoutPanel for responsive layout and consistent padding. (2) Segoe UI 9pt. (3) Extracted GUI helpers and Exclusions dialog. (4) Input validation and user-facing error messages. (5) Optional logging when $Script:EnableLogging is $true.
# Future improvements: tooltips on buttons; "Open folder" for config/CSV path; single-window layout (grid in tab or splitter); configurable paths in a small settings file.
$form = New-Object System.Windows.Forms.Form
$form.Text = "Crimson Binds"
$form.Size = New-Object System.Drawing.Size(740, 360)
$form.StartPosition = "CenterScreen"
$form.MinimumSize = New-Object System.Drawing.Size(520, 340)
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9)

$padding = 8
$contentPanel = New-Object System.Windows.Forms.Panel
$contentPanel.Dock = "Fill"
$contentPanel.Padding = New-Object System.Windows.Forms.Padding($padding)
[void]$form.Controls.Add($contentPanel)

# Main layout: 7 rows, 3 columns (label | fill | button). Row heights AutoSize; column 0 AutoSize, 1 Percent 100, 2 AutoSize.
$mainTable = New-Object System.Windows.Forms.TableLayoutPanel
$mainTable.Dock = "Fill"
$mainTable.ColumnCount = 3
[void]$mainTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
[void]$mainTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
[void]$mainTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
$mainTable.RowCount = 7
foreach ($i in 0..6) { [void]$mainTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize))) }
$mainTable.Padding = New-Object System.Windows.Forms.Padding(0)
[void]$contentPanel.Controls.Add($mainTable)

$cellMargin = New-Object System.Windows.Forms.Padding($padding, 4, $padding, 4)

# Row 0: config.ini
$lblConfigPath = New-Object System.Windows.Forms.Label
$lblConfigPath.Text = "config.ini:"
$lblConfigPath.Margin = $cellMargin
$lblConfigPath.Anchor = "Top,Left"
$txtConfigPath = New-Object System.Windows.Forms.TextBox
$txtConfigPath.Height = 22
$txtConfigPath.ReadOnly = $true
$txtConfigPath.Margin = $cellMargin
$txtConfigPath.Anchor = "Top,Left,Right"
$btnBrowseConfig = New-Object System.Windows.Forms.Button
$btnBrowseConfig.Text = "Browse..."
$btnBrowseConfig.Size = New-Object System.Drawing.Size(80, 26)
$btnBrowseConfig.Margin = $cellMargin
$btnBrowseConfig.Anchor = "Top,Right"
$mainTable.Controls.Add($lblConfigPath, 0, 0)
$mainTable.Controls.Add($txtConfigPath, 1, 0)
$mainTable.Controls.Add($btnBrowseConfig, 2, 0)

# Row 1: CSV
$lblCsvPath = New-Object System.Windows.Forms.Label
$lblCsvPath.Text = "CSV (binds):"
$lblCsvPath.Margin = $cellMargin
$lblCsvPath.Anchor = "Top,Left"
$txtCsvPath = New-Object System.Windows.Forms.TextBox
$txtCsvPath.Height = 22
$txtCsvPath.ReadOnly = $true
$txtCsvPath.Margin = $cellMargin
$txtCsvPath.Anchor = "Top,Left,Right"
$btnBrowseCsv = New-Object System.Windows.Forms.Button
$btnBrowseCsv.Text = "Browse..."
$btnBrowseCsv.Size = New-Object System.Drawing.Size(80, 26)
$btnBrowseCsv.Margin = $cellMargin
$btnBrowseCsv.Anchor = "Top,Right"
$mainTable.Controls.Add($lblCsvPath, 0, 1)
$mainTable.Controls.Add($txtCsvPath, 1, 1)
$mainTable.Controls.Add($btnBrowseCsv, 2, 1)

# Row 2: Filter
$lblFilter = New-Object System.Windows.Forms.Label
$lblFilter.Text = "Filter:"
$lblFilter.Margin = $cellMargin
$lblFilter.Anchor = "Top,Left"
$comboFilter = New-Object System.Windows.Forms.ComboBox
$comboFilter.DropDownStyle = "DropDownList"
$comboFilter.DropDownHeight = 220
$comboFilter.MinimumSize = New-Object System.Drawing.Size(180, 24)
$comboFilter.Margin = $cellMargin
$comboFilter.Anchor = "Top,Left"
[void]$comboFilter.Items.Add("(All sections)")
$comboFilter.SelectedIndex = 0
$mainTable.Controls.Add($lblFilter, 0, 2)
$mainTable.Controls.Add($comboFilter, 1, 2)

# Row 3: Action buttons in a single cell (FlowLayoutPanel left-to-right; no GroupBox to avoid grey overlap)
$flowActions = New-Object System.Windows.Forms.FlowLayoutPanel
$flowActions.FlowDirection = "LeftToRight"
$flowActions.WrapContents = $false
$flowActions.AutoSize = $true
$flowActions.Margin = $cellMargin
$flowActions.Anchor = "Top,Left"
$btnExportConfig = New-Object System.Windows.Forms.Button
$btnExportConfig.Text = "Export config -> CSV"
$btnExportConfig.Size = New-Object System.Drawing.Size(130, 26)
$btnLoadCsv = New-Object System.Windows.Forms.Button
$btnLoadCsv.Text = "Load CSV"
$btnLoadCsv.Size = New-Object System.Drawing.Size(75, 26)
$btnRandomize = New-Object System.Windows.Forms.Button
$btnRandomize.Text = "Randomize keys"
$btnRandomize.Size = New-Object System.Drawing.Size(110, 26)
$btnUpdateConfig = New-Object System.Windows.Forms.Button
$btnUpdateConfig.Text = "Update config.ini"
$btnUpdateConfig.Size = New-Object System.Drawing.Size(120, 26)
$btnSaveCsv = New-Object System.Windows.Forms.Button
$btnSaveCsv.Text = "Save CSV"
$btnSaveCsv.Size = New-Object System.Drawing.Size(75, 26)
$btnUpdateDebounce = New-Object System.Windows.Forms.Button
$btnUpdateDebounce.Text = "Update Debounce"
$btnUpdateDebounce.Size = New-Object System.Drawing.Size(120, 26)
$flowActions.Controls.AddRange(@($btnExportConfig, $btnLoadCsv, $btnRandomize, $btnUpdateConfig, $btnSaveCsv, $btnUpdateDebounce))
$mainTable.Controls.Add($flowActions, 0, 3)
$mainTable.SetColumnSpan($flowActions, 3)

# Row 4: WoW account
$lblAccount = New-Object System.Windows.Forms.Label
$lblAccount.Text = "WoW Account"
$lblAccount.Margin = $cellMargin
$lblAccount.Anchor = "Top,Left"
$comboAccount = New-Object System.Windows.Forms.ComboBox
$comboAccount.DropDownStyle = "DropDownList"
$comboAccount.MinimumSize = New-Object System.Drawing.Size(120, 24)
$comboAccount.Margin = $cellMargin
$comboAccount.Anchor = "Top,Left"
[void]$comboAccount.Items.Add("(Pick account...)")
$comboAccount.SelectedIndex = 0
$mainTable.Controls.Add($lblAccount, 0, 4)
$mainTable.Controls.Add($comboAccount, 1, 4)

# Row 5: Backup / Restore
$btnBackup = New-Object System.Windows.Forms.Button
$btnBackup.Text = "Backup"
$btnBackup.Size = New-Object System.Drawing.Size(80, 26)
$btnBackup.BackColor = [System.Drawing.Color]::LightGreen
$btnBackup.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnBackup.Margin = $cellMargin
$btnBackup.Anchor = "Top,Left"
$lblRestore = New-Object System.Windows.Forms.Label
$lblRestore.Text = "Restore from:"
$lblRestore.Margin = $cellMargin
$lblRestore.Anchor = "Top,Left"
$lblRestore.BackColor = [System.Drawing.Color]::Transparent
$comboBackup = New-Object System.Windows.Forms.ComboBox
$comboBackup.DropDownStyle = "DropDownList"
$comboBackup.MinimumSize = New-Object System.Drawing.Size(120, 24)
$comboBackup.Margin = $cellMargin
$comboBackup.Anchor = "Top,Left"
$btnRestore = New-Object System.Windows.Forms.Button
$btnRestore.Text = "Restore from backup"
$btnRestore.Size = New-Object System.Drawing.Size(130, 26)
$btnRestore.Margin = $cellMargin
$btnRestore.Anchor = "Top,Left"
$mainTable.Controls.Add($btnBackup, 0, 5)
$flowBackup = New-Object System.Windows.Forms.FlowLayoutPanel
$flowBackup.FlowDirection = "LeftToRight"
$flowBackup.WrapContents = $false
$flowBackup.AutoSize = $true
$flowBackup.Margin = $cellMargin
$flowBackup.Controls.Add($lblRestore)
$flowBackup.Controls.Add($comboBackup)
$flowBackup.Controls.Add($btnRestore)
$mainTable.Controls.Add($flowBackup, 1, 5)
$mainTable.SetColumnSpan($flowBackup, 2)

# Row 6: Status
$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Text = "Browse to config.ini, then Export config -> CSV. Edit CSV in Excel if you like. Load CSV, filter by section, edit keys in grid. Randomize / Update config.ini / Update Debounce."
$lblStatus.Margin = $cellMargin
$lblStatus.Anchor = "Top,Left,Right,Bottom"
$lblStatus.AutoSize = $false
$lblStatus.MinimumSize = New-Object System.Drawing.Size(0, 36)
$mainTable.Controls.Add($lblStatus, 0, 6)
$mainTable.SetColumnSpan($lblStatus, 3)

# Data grid in a separate window (toolbar uses Panel; no GroupBox)
$gridForm = New-Object System.Windows.Forms.Form
$gridForm.Text = "Crimson Binds - Data Grid (Section / Action / Key / MacroText)"
$gridForm.Size = New-Object System.Drawing.Size(720, 420)
$gridForm.StartPosition = [System.Windows.Forms.FormStartPosition]::Manual
$gridForm.MinimumSize = New-Object System.Drawing.Size(400, 200)
$gridForm.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$dgv = New-Object System.Windows.Forms.DataGridView
$dgv.Dock = "Fill"
$dgv.AutoSizeColumnsMode = "Fill"
$dgv.AllowUserToAddRows = $false
$dgv.ColumnCount = 4
$dgv.Columns[0].Name = "Section"; $dgv.Columns[0].ReadOnly = $true
$dgv.Columns[1].Name = "Action"; $dgv.Columns[1].ReadOnly = $true
$dgv.Columns.RemoveAt(2)
$keyCol = New-Object System.Windows.Forms.DataGridViewComboBoxColumn
$keyCol.Name = "Key"
$keyCol.HeaderText = "Key"
$keyCol.ReadOnly = $false
$keyCol.DisplayStyle = [System.Windows.Forms.DataGridViewComboBoxDisplayStyle]::ComboBox
$keyCol.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$keyCol.DataSource = [object[]](Get-FullKeyPool)
$dgv.Columns.Insert(2, $keyCol)
$dgv.Columns[3].Name = "MacroText"; $dgv.Columns[3].ReadOnly = $false
$dgv.Add_DataError({
    param($sender, $e)
    if ($e.ColumnIndex -ge 0 -and $sender.Columns[$e.ColumnIndex].Name -eq "Key") { $e.ThrowException = $false }
})
$dgv.Add_CellValueChanged({
    param($s, $e)
    if (-not $Script:AllRows -or $e.ColumnIndex -lt 0) { return }
    if ($s.Columns[$e.ColumnIndex].Name -ne "Key" -and $s.Columns[$e.ColumnIndex].Name -ne "MacroText") { return }
    Sync-GridToRows -Grid $s -Rows $Script:AllRows -FilterSection $Script:CurrentFilterSection
    $duplicateSet = Get-DuplicateKeysBySection -Rows $Script:AllRows
    $normalBack = [System.Drawing.Color]::White
    $duplicateBack = [System.Drawing.Color]::FromArgb(255, 220, 220)
    foreach ($row in $s.Rows) {
        $sec = $row.Cells["Section"].Value; $keyVal = $row.Cells["Key"].Value
        $secStr = if ($null -ne $sec) { $sec.ToString().Trim() } else { "" }
        $keyStr = if ($null -ne $keyVal -and ($keyVal.ToString().Trim())) { $keyVal.ToString().Trim() } else { "" }
        if ($keyStr -and $duplicateSet["$secStr|$keyStr"]) {
            $row.Cells["Key"].Style.BackColor = $duplicateBack
        } else {
            $row.Cells["Key"].Style.BackColor = $normalBack
        }
    }
})
$gridPanel = New-Object System.Windows.Forms.Panel
$gridPanel.Dock = "Bottom"
$gridPanel.Height = 40
$gridPanel.Padding = New-Object System.Windows.Forms.Padding($padding, 6, $padding, 0)
$btnSaveFromGrid = New-Object System.Windows.Forms.Button
$btnSaveFromGrid.Text = "Save to CSV"
$btnSaveFromGrid.Size = New-Object System.Drawing.Size(100, 26)
$btnSaveFromGrid.Location = New-Object System.Drawing.Point($padding, 6)
$btnExclusions = New-Object System.Windows.Forms.Button
$btnExclusions.Text = "Exclusions"
$btnExclusions.Size = New-Object System.Drawing.Size(90, 26)
$btnExclusions.Location = New-Object System.Drawing.Point(114, 6)
$btnRandomizeSelected = New-Object System.Windows.Forms.Button
$btnRandomizeSelected.Text = "Randomize selected"
$btnRandomizeSelected.Size = New-Object System.Drawing.Size(110, 26)
$btnRandomizeSelected.Location = New-Object System.Drawing.Point(210, 6)
$lblSearch = New-Object System.Windows.Forms.Label
$lblSearch.Text = "Search:"
$lblSearch.AutoSize = $true
$lblSearch.Location = New-Object System.Drawing.Point(326, 10)
$txtGridSearch = New-Object System.Windows.Forms.TextBox
$txtGridSearch.Size = New-Object System.Drawing.Size(140, 22)
$txtGridSearch.Location = New-Object System.Drawing.Point(371, 7)
if ($Script:GridSearchText) { $txtGridSearch.Text = $Script:GridSearchText }
$txtGridSearch.Add_KeyDown({
    param($s, $e)
    if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Return) {
        $e.SuppressKeyPress = $true
        $Script:GridSearchText = $txtGridSearch.Text.Trim()
        Update-GridFromRows -Grid $dgv -Rows $Script:AllRows -FilterSection $Script:CurrentFilterSection -SearchText $Script:GridSearchText
    }
})
$btnSearch = New-Object System.Windows.Forms.Button
$btnSearch.Text = "Search"
$btnSearch.Size = New-Object System.Drawing.Size(60, 26)
$btnSearch.Location = New-Object System.Drawing.Point(516, 6)
$btnClearSearch = New-Object System.Windows.Forms.Button
$btnClearSearch.Text = "Clear"
$btnClearSearch.Size = New-Object System.Drawing.Size(55, 26)
$btnClearSearch.Location = New-Object System.Drawing.Point(581, 6)
[void]$gridPanel.Controls.Add($btnSaveFromGrid)
[void]$gridPanel.Controls.Add($btnExclusions)
[void]$gridPanel.Controls.Add($btnRandomizeSelected)
[void]$gridPanel.Controls.Add($lblSearch)
[void]$gridPanel.Controls.Add($txtGridSearch)
[void]$gridPanel.Controls.Add($btnSearch)
[void]$gridPanel.Controls.Add($btnClearSearch)
# Add Fill (grid) first, then Bottom (toolbar) so the grid client area stops above the bar — not under it.
[void]$gridForm.Controls.Add($dgv)
[void]$gridForm.Controls.Add($gridPanel)
$gridForm.Add_FormClosing({
    param($sender, $e)
    if ($form.Visible) { $e.Cancel = $true; $sender.Hide() }
})

# ---------- Event wiring (thin handlers: get state, call logic, update UI) ----------
$comboFilter.Add_SelectedIndexChanged({
    Sync-GridToRows -Grid $dgv -Rows $Script:AllRows -FilterSection $Script:CurrentFilterSection
    $filter = Get-CurrentFilterSection -ComboFilter $comboFilter
    Update-GridFromRows -Grid $dgv -Rows $Script:AllRows -FilterSection $filter
})

$btnBrowseConfig.Add_Click({
    $fd = New-Object System.Windows.Forms.OpenFileDialog
    $fd.Filter = "config.ini|config.ini|All (*.*)|*.*"
    if ($Script:ConfigPath) { $fd.InitialDirectory = [System.IO.Path]::GetDirectoryName($Script:ConfigPath) }
    if ($fd.ShowDialog() -eq "OK") {
        $Script:ConfigPath = $fd.FileName
        $txtConfigPath.Text = $Script:ConfigPath
        $Script:WtfDir = Get-WtfDirFromConfig -ConfigPath $Script:ConfigPath
        Update-AccountCombo -ConfigPath $Script:ConfigPath -ComboAccount $comboAccount
        Update-BackupCombo -ComboBackup $comboBackup
        Update-BackupRestoreState -btnBackup $btnBackup -comboBackup $comboBackup -btnRestore $btnRestore
    }
})

$btnBrowseCsv.Add_Click({
    $fd = New-Object System.Windows.Forms.OpenFileDialog
    $fd.Filter = "CSV|*.csv|All (*.*)|*.*"
    $fd.DefaultExt = "csv"
    $fd.InitialDirectory = $ToolDir
    $defaultCsv = Join-Path $ToolDir "bindpad_defaults.csv"
    if (Test-Path -LiteralPath $defaultCsv -PathType Leaf) { $fd.FileName = [System.IO.Path]::GetFileName($defaultCsv) }
    if ($fd.ShowDialog() -eq "OK") {
        $Script:CsvPath = $fd.FileName
        $txtCsvPath.Text = $Script:CsvPath
        try {
            $Script:AllRows = Import-CsvToRows -Path $Script:CsvPath
            Update-SectionFilterCombo -Rows $Script:AllRows -ComboFilter $comboFilter
            Update-GridFromRows -Grid $dgv -Rows $Script:AllRows -FilterSection $null
            $lblStatus.Text = "Loaded $($Script:AllRows.Count) rows from CSV."
        } catch {
            $lblStatus.Text = "Error loading CSV: $_"
            Write-CrimsonBindLog -Message "Import-CsvToRows failed: $_" -Level "Error"
        }
    }
})

$btnExportConfig.Add_Click({
    if (-not $Script:ConfigPath -or -not (Test-Path -LiteralPath $Script:ConfigPath -PathType Leaf)) { $lblStatus.Text = "Select config.ini first."; return }
    $outPath = $Script:CsvPath
    if (-not $outPath) { $outPath = Join-Path $ToolDir "binds_export.csv" }
    try {
        $rows = Get-ConfigIniSectionsAndBinds -ConfigPath $Script:ConfigPath
        # Reverse-translate config.ini key format (+sc21_67701769) to human CSV keys (e.g. SHIFT-F) for editing; Update config/Debounce will translate back.
        foreach ($r in $rows) {
            if ($r.Key) { $r.Key = ConvertFrom-ConfigIniKey -ConfigKey $r.Key }
        }
        if ($Script:AllRows -and $Script:AllRows.Count -gt 0 -and $Script:CsvPath) {
            $existingByKey = @{}
            foreach ($r in $Script:AllRows) { $existingByKey["$($r.Section)|$($r.ActionName)"] = $r }
            $parsedByKey = @{}
            foreach ($r in $rows) { $parsedByKey["$($r.Section)|$($r.ActionName)"] = $r }
            $merged = [System.Collections.ArrayList]::new()
            foreach ($r in $Script:AllRows) {
                $k = "$($r.Section)|$($r.ActionName)"
                if ($parsedByKey.ContainsKey($k)) { [void]$merged.Add($parsedByKey[$k]) } else { [void]$merged.Add($r) }
            }
            foreach ($r in $rows) {
                $k = "$($r.Section)|$($r.ActionName)"
                if (-not $existingByKey.ContainsKey($k)) { [void]$merged.Add($r) }
            }
            $rows = $merged
            $filtered = [System.Collections.ArrayList]::new()
            foreach ($r in $rows) {
                if ($r.Section -eq "General" -or $r.ActionName -notmatch '^Target Member\d+$') { [void]$filtered.Add($r) }
            }
            $rows = $filtered
        }
        Export-RowsToCsv -Path $outPath -Rows $rows
        $Script:AllRows = $rows
        $Script:CsvPath = $outPath
        $txtCsvPath.Text = $outPath
        Update-SectionFilterCombo -Rows $Script:AllRows -ComboFilter $comboFilter
        Update-GridFromRows -Grid $dgv -Rows $Script:AllRows -FilterSection $null
        $lblStatus.Text = "Exported $($rows.Count) rows to $outPath"
    } catch {
        $lblStatus.Text = "Error exporting: $_"
        Write-CrimsonBindLog -Message "Export config to CSV failed: $_" -Level "Error"
    }
})

$btnLoadCsv.Add_Click({
    if (-not $Script:CsvPath -or -not (Test-Path -LiteralPath $Script:CsvPath -PathType Leaf)) { $lblStatus.Text = "Select a CSV file first."; return }
    try {
        $Script:AllRows = Import-CsvToRows -Path $Script:CsvPath
        Update-SectionFilterCombo -Rows $Script:AllRows -ComboFilter $comboFilter
        Update-GridFromRows -Grid $dgv -Rows $Script:AllRows -FilterSection $null
        if ($gridForm -and -not $gridForm.IsDisposed -and -not $gridForm.Visible) { $gridForm.Show() }
        $lblStatus.Text = "Loaded $($Script:AllRows.Count) rows."
    } catch {
        $lblStatus.Text = "Error loading CSV: $_"
        Write-CrimsonBindLog -Message "Import-CsvToRows (Load CSV button) failed: $_" -Level "Error"
    }
})

$btnRandomize.Add_Click({
    $filter = Get-CurrentFilterSection -ComboFilter $comboFilter
    Sync-GridToRows -Grid $dgv -Rows $Script:AllRows -FilterSection $filter
    Invoke-RandomizeKeys -Rows $Script:AllRows -OnlySection $filter
    Update-GridFromRows -Grid $dgv -Rows $Script:AllRows -FilterSection $filter
    $scope = if ($filter) { " for section: $filter" } else { "" }
    $lblStatus.Text = "Randomized keys$scope. Click Update config.ini or Update Debounce to write files."
})

$btnUpdateConfig.Add_Click({
    $filter = Get-CurrentFilterSection -ComboFilter $comboFilter
    Sync-GridToRows -Grid $dgv -Rows $Script:AllRows -FilterSection $filter
    if (-not $Script:ConfigPath -or -not (Test-Path -LiteralPath $Script:ConfigPath -PathType Leaf)) { $lblStatus.Text = "Select config.ini first."; return }
    try {
        $n = Set-ConfigIniFromRows -ConfigPath $Script:ConfigPath -Rows $Script:AllRows -OnlySection $filter
        $val = Get-ConfigIniUpdateValidation -ConfigPath $Script:ConfigPath -Rows $Script:AllRows -OnlySection $filter
        $scope = if ($filter) { " (section: $filter)" } else { "" }
        $msg = "Updated $n keybinds in config.ini$scope. Validation: $($val.Match) match"
        if ($val.Match -gt $n) { $msg += " ($($val.Match - $n) already correct)" }
        if ($val.Mismatch -gt 0) { $msg += ", $($val.Mismatch) mismatch" }
        if ($val.Skipped -gt 0) { $msg += ", $($val.Skipped) skipped (not in config)" }
        if ($val.NoKeyInCsv -gt 0) { $msg += ", $($val.NoKeyInCsv) empty keybind in CSV" }
        $msg += "."
        if ($val.DuplicateKeyInSection -and $val.DuplicateKeyInSection.Count -gt 0) {
            $msg += " Duplicate key in same section (only one action gets the key): $($val.DuplicateKeyInSection.Count). Change one in CSV to fix."
            Write-CrimsonBindLog -Message "Duplicate key in section: $($val.DuplicateKeyInSection -join '; ')" -Level "Warning"
            Show-DuplicateKeysPopup -DuplicateList $val.DuplicateKeyInSection -Owner $form
        }
        $lblStatus.Text = $msg
        if ($val.Mismatch -gt 0 -and $val.MismatchDetails.Count -gt 0) {
            Write-CrimsonBindLog -Message "Config.ini validation mismatches: $($val.MismatchDetails -join '; ')" -Level "Warning"
        }
    } catch {
        $lblStatus.Text = "Error updating config.ini: $_"
        Write-CrimsonBindLog -Message "Set-ConfigIniFromRows failed: $_" -Level "Error"
    }
})

$btnSaveCsv.Add_Click({
    $filter = Get-CurrentFilterSection -ComboFilter $comboFilter
    Sync-GridToRows -Grid $dgv -Rows $Script:AllRows -FilterSection $filter
    $path = $Script:CsvPath
    if (-not $path) { $path = Join-Path $ToolDir "binds_export.csv" }
    try {
        Export-RowsToCsv -Path $path -Rows $Script:AllRows
        $lblStatus.Text = "Saved $($Script:AllRows.Count) rows to $path"
    } catch {
        $lblStatus.Text = "Error saving CSV: $_"
        Write-CrimsonBindLog -Message "Export-RowsToCsv failed: $_" -Level "Error"
    }
})

$btnUpdateDebounce.Add_Click({
    if (-not $Script:WtfDir) { $lblStatus.Text = "Set config.ini with WTFPath first."; return }
    $account = $null
    if ($comboAccount.SelectedItem -and $comboAccount.SelectedItem -ne "(Pick account...)") { $account = $comboAccount.SelectedItem }
    if (-not $account) { $lblStatus.Text = "Pick a WoW account for Debounce path."; return }
    if (-not $Script:AllRows -or $Script:AllRows.Count -eq 0) { $lblStatus.Text = "Load CSV first (or Export config -> CSV)."; return }
    $filter = Get-CurrentFilterSection -ComboFilter $comboFilter
    $path = Join-Path $Script:WtfDir "Account\$account\SavedVariables\Debounce.lua"
    $dir = [System.IO.Path]::GetDirectoryName($path)
    if (-not (Test-Path -LiteralPath $dir -PathType Container)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    try {
        Export-DebounceFromRows -Path $path -Rows $Script:AllRows -OnlySection $filter -ExistingBindSections $null
        $scope = if ($filter) { "section '$filter'" } else { "all sections" }
        $lblStatus.Text = "Updated Debounce.lua ($scope) at $path"
    } catch {
        $lblStatus.Text = "Debounce error: $_"
        Write-CrimsonBindLog -Message "Export-DebounceFromRows failed: $_" -Level "Error"
    }
})

$btnBackup.Add_Click({
    if (-not $Script:ConfigPath -or -not (Test-Path -LiteralPath $Script:ConfigPath -PathType Leaf)) { $lblStatus.Text = "Select config.ini first."; return }
    if (-not $Script:WtfDir) { $lblStatus.Text = "config.ini must have a valid WTFPath so we can find the WTF folder."; return }
    $account = $null
    if ($comboAccount.SelectedItem -and $comboAccount.SelectedItem -ne "(Pick account...)") { $account = $comboAccount.SelectedItem }
    try {
        $result = New-CrimsonBackup -ConfigPath $Script:ConfigPath -WtfDir $Script:WtfDir -AccountName $account -CsvPath $Script:CsvPath
        if ($result) {
            Update-BackupCombo -ComboBackup $comboBackup
            Update-BackupRestoreState -btnBackup $btnBackup -comboBackup $comboBackup -btnRestore $btnRestore
            $lblStatus.Text = "Backed up $($result.FileCount) file(s) to $($result.Path)"
        } else { $lblStatus.Text = "Backup failed." }
    } catch {
        $lblStatus.Text = "Backup error: $_"
        Write-CrimsonBindLog -Message "New-CrimsonBackup failed: $_" -Level "Error"
    }
})

$btnRestore.Add_Click({
    $sel = $comboBackup.SelectedItem
    if (-not $sel) { $lblStatus.Text = "Select a backup from the list."; return }
    if (-not $Script:ConfigPath) { $lblStatus.Text = "Select config.ini first."; return }
    $account = $null
    if ($comboAccount.SelectedItem -and $comboAccount.SelectedItem -ne "(Pick account...)") { $account = $comboAccount.SelectedItem }
    $backupDir = Join-Path $Script:BackupRoot $sel
    if (-not (Test-Path -LiteralPath $backupDir -PathType Container)) { $lblStatus.Text = "Backup folder not found."; return }
    try {
        $n = Restore-CrimsonBackup -BackupFolderPath $backupDir -ConfigPath $Script:ConfigPath -WtfDir $Script:WtfDir -AccountName $account
        if ($Script:WtfDir -and $account) {
            $lblStatus.Text = "Restored $n file(s). config.ini and SavedVariables (Debounce) overwritten."
        } else {
            $suffix = if (-not $Script:WtfDir) { " (SavedVariables skipped - no WTF path set.)" } else { " (Pick WoW account to restore SavedVariables.)" }
            $lblStatus.Text = "Restored $n file(s). config.ini restored." + $suffix
        }
    } catch {
        $lblStatus.Text = "Restore error: $_"
        Write-CrimsonBindLog -Message "Restore-CrimsonBackup failed: $_" -Level "Error"
    }
})

$comboBackup.Add_SelectedIndexChanged({
    $btnRestore.Enabled = ($null -ne $comboBackup.SelectedItem) -and $Script:ConfigPath
})

$btnSaveFromGrid.Add_Click({
    $filter = Get-CurrentFilterSection -ComboFilter $comboFilter
    Sync-GridToRows -Grid $dgv -Rows $Script:AllRows -FilterSection $filter
    $path = $Script:CsvPath
    if (-not $path) { $path = Join-Path $ToolDir "binds_export.csv" }
    try {
        Export-RowsToCsv -Path $path -Rows $Script:AllRows
        $lblStatus.Text = "Saved $($Script:AllRows.Count) rows to CSV from grid."
    } catch {
        $lblStatus.Text = "Save error: $_"
        Write-CrimsonBindLog -Message "Export-RowsToCsv (from grid) failed: $_" -Level "Error"
    }
})

$btnExclusions.Add_Click({
    Show-ExclusionsDialog -StatusLabel $lblStatus -ExcludedKeysPath $Script:ExcludedKeysPath
})

$btnRandomizeSelected.Add_Click({
    $filter = Get-CurrentFilterSection -ComboFilter $comboFilter
    Sync-GridToRows -Grid $dgv -Rows $Script:AllRows -FilterSection $filter
    $n = Invoke-RandomizeSelectedGridRows -Grid $dgv -Rows $Script:AllRows -FilterSection $filter
    Update-GridFromRows -Grid $dgv -Rows $Script:AllRows -FilterSection $filter -SearchText $Script:GridSearchText
    if ($n -gt 0) {
        $lblStatus.Text = "Randomized $n selected row(s) to unique keys (no duplicates in section; General keys excluded for specs)."
    } else {
        $lblStatus.Text = "No rows randomized. Select one or more rows in the grid (skip rows that inherit key from General), then click Randomize selected."
    }
})

$btnSearch.Add_Click({
    $Script:GridSearchText = $txtGridSearch.Text.Trim()
    Update-GridFromRows -Grid $dgv -Rows $Script:AllRows -FilterSection $Script:CurrentFilterSection -SearchText $Script:GridSearchText
})

$btnClearSearch.Add_Click({
    $txtGridSearch.Text = ""
    $Script:GridSearchText = ""
    Update-GridFromRows -Grid $dgv -Rows $Script:AllRows -FilterSection $Script:CurrentFilterSection
})

$form.Add_Shown({
    if ($Script:ConfigPath) {
        $Script:WtfDir = Get-WtfDirFromConfig -ConfigPath $Script:ConfigPath
        Update-AccountCombo -ConfigPath $Script:ConfigPath -ComboAccount $comboAccount
    }
    Update-BackupCombo -ComboBackup $comboBackup
    Update-BackupRestoreState -btnBackup $btnBackup -comboBackup $comboBackup -btnRestore $btnRestore
    $form.Add_FormClosing({ if ($gridForm -and $gridForm.Visible) { $gridForm.Close() } })
    $gap = 12
    $gridForm.Location = New-Object System.Drawing.Point(($form.Left + $form.Width + $gap), $form.Top)
    $gridForm.Show()
})

$form.ShowDialog() | Out-Null
if ($gridForm -and -not $gridForm.IsDisposed) { $gridForm.Close() }
