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
            $currentSection = $Matches[1].Trim()
            continue
        }
        if ($null -ne $currentSection -and $strip -match '^([^=]+)=(.*)$') {
            $actionName = $Matches[1].Trim()
            $value = $Matches[2].Trim()
            if (-not $actionName) { continue }
            if ($actionName -match '^START\s') { continue }
            $keyPart = $value
            $afterSemi = ""
            if ($value -match '^([^;]*);(.*)$') {
                $keyPart = $Matches[1].Trim()
                $afterSemi = $Matches[2].Trim()
            }
            [void]$out.Add([PSCustomObject]@{
                Section    = $currentSection
                ActionName = $actionName
                Key        = $keyPart
                MacroText  = $afterSemi
                TextureID  = "132089"
            })
        }
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
    foreach ($line in $lines) {
        $strip = $line.TrimEnd("`r", "`n")
        if ($strip -match '^\[(.+)\]$') {
            $currentSection = $Matches[1].Trim()
            [void]$newLines.Add($line)
            continue
        }
        if ($null -ne $currentSection -and $strip -match '^([^=]+)=(.*)$') {
            $actionName = $Matches[1].Trim()
            $value = $Matches[2].Trim()
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
    $outText = ($newLines -join "`r`n") + "`r`n"
    [System.IO.File]::WriteAllText($ConfigPath, $outText, [System.Text.Encoding]::Unicode)
    return $replaced
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
        $sec = if ($r.Section) { $r.Section.Trim() } else { "" }
        $act = if ($r.ActionName) { $r.ActionName.Trim() } else { "" }
        $key = if ($r.BindPadKey) { $r.BindPadKey.Trim() } else { if ($r.Key) { $r.Key.Trim() } else { "" } }
        $mac = if ($r.MacroText) { $r.MacroText.Trim() } else { "" }
        $tex = if ($r.TextureID) { $r.TextureID.Trim() } else { "132089" }
        if ($sec -and $act) {
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
    "A","S","D","E","W","Q"
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
    param([System.Collections.ArrayList]$Rows)
    $fullPool = Get-FullKeyPoolForRandomize
    $rng = [System.Random]::new()
    $shuffled = [System.Collections.ArrayList]::new()
    foreach ($k in ($fullPool | Sort-Object { $rng.Next() })) { [void]$shuffled.Add($k) }
    $generalList = [System.Collections.ArrayList]::new()
    $bySection = @{}
    foreach ($r in $Rows) {
        if ($r.Section -eq "General") { [void]$generalList.Add($r) }
        else {
            if (-not $bySection[$r.Section]) { $bySection[$r.Section] = [System.Collections.ArrayList]::new() }
            [void]$bySection[$r.Section].Add($r)
        }
    }
    $idx = 0
    foreach ($r in $generalList) {
        if ($idx -ge $shuffled.Count) { break }
        $r.Key = $shuffled[$idx]; $idx++
    }
    foreach ($sec in $bySection.Keys) {
        foreach ($r in $bySection[$sec]) {
            if ($idx -ge $shuffled.Count) { $idx = 0 }
            $r.Key = $shuffled[$idx]; $idx++
        }
    }
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
    param([string]$ConfigPath, [string]$WtfDir, [string]$AccountName)
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

# Build Lua for one bind entry (macrotext with name, value, icon, key).
function Get-DebounceBindLua {
    param($Row)
    $name = Get-LuaEscaped -s $Row.ActionName
    $value = Get-LuaEscaped -s $Row.MacroText
    $icon = 132089
    if ($Row.TextureID -match '^\d+$') { $icon = [int]$Row.TextureID }
    $key = Get-LuaEscaped -s $Row.Key
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

# Build full class table: tabs 1,2,3 and optionally [0]. Do not use for GENERAL.
function Get-DebounceSectionTableLua {
    param([string]$ClassKey, [hashtable]$TabsByIndex)
    $sb = [System.Text.StringBuilder]::new()
    [void]$sb.Append("[""$ClassKey""] = {")
    $sb.AppendLine()
    $tabIndices = @($TabsByIndex.Keys | Where-Object { $_ -ne 0 } | Sort-Object)
    foreach ($idx in $tabIndices) {
        $tabLua = $TabsByIndex[$idx]
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
function Update-GridFromRows {
    param($Grid, [System.Collections.ArrayList]$Rows, [string]$FilterSection, [string]$SearchText)
    if ($null -eq $SearchText -and $Script:GridSearchText) { $SearchText = $Script:GridSearchText }
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
    $keyCol = $Grid.Columns["Key"]
    if ($keyCol -and $keyCol -is [System.Windows.Forms.DataGridViewComboBoxColumn]) {
        $keyCol.DataSource = $null
        $keyCol.DataSource = [object[]]$keyList
    }
    $Grid.Rows.Clear()
    $generalActionNames = @{}
    if ($FilterSection -and $FilterSection -ne "General") {
        foreach ($r in $Rows) { if ($r.Section -eq "General") { $generalActionNames[$r.ActionName] = $true } }
    }
    $search = if ($SearchText) { $SearchText.Trim() } else { "" }
    foreach ($r in $Rows) {
        if ($FilterSection -and $r.Section -ne $FilterSection) { continue }
        if ($FilterSection -and $FilterSection -ne "General" -and $generalActionNames.ContainsKey($r.ActionName)) { continue }
        if ($search) {
            $s = if ($r.Section) { $r.Section } else { "" }
            $a = if ($r.ActionName) { $r.ActionName } else { "" }
            $k = if ($r.Key) { $r.Key } else { "" }
            $m = if ($r.MacroText) { $r.MacroText } else { "" }
            $combined = "$s $a $k $m".ToUpperInvariant()
            if ($combined.IndexOf($search.ToUpperInvariant()) -lt 0) { continue }
        }
        [void]$Grid.Rows.Add($r.Section, $r.ActionName, $r.Key, $r.MacroText)
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
        if ($FilterSection -and $FilterSection -ne "General" -and $generalActionNames.ContainsKey($r.ActionName)) { continue }
        if ($key -lt $Grid.Rows.Count) {
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
$lblSearch = New-Object System.Windows.Forms.Label
$lblSearch.Text = "Search:"
$lblSearch.AutoSize = $true
$lblSearch.Location = New-Object System.Drawing.Point(210, 10)
$txtGridSearch = New-Object System.Windows.Forms.TextBox
$txtGridSearch.Size = New-Object System.Drawing.Size(140, 22)
$txtGridSearch.Location = New-Object System.Drawing.Point(255, 7)
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
$btnSearch.Location = New-Object System.Drawing.Point(400, 6)
$btnClearSearch = New-Object System.Windows.Forms.Button
$btnClearSearch.Text = "Clear"
$btnClearSearch.Size = New-Object System.Drawing.Size(55, 26)
$btnClearSearch.Location = New-Object System.Drawing.Point(465, 6)
[void]$gridPanel.Controls.Add($btnSaveFromGrid)
[void]$gridPanel.Controls.Add($btnExclusions)
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
    Invoke-RandomizeKeys -Rows $Script:AllRows
    Update-GridFromRows -Grid $dgv -Rows $Script:AllRows -FilterSection $filter
    $lblStatus.Text = "Randomized keys. Click Update config.ini or Update Debounce to write files."
})

$btnUpdateConfig.Add_Click({
    $filter = Get-CurrentFilterSection -ComboFilter $comboFilter
    Sync-GridToRows -Grid $dgv -Rows $Script:AllRows -FilterSection $filter
    if (-not $Script:ConfigPath -or -not (Test-Path -LiteralPath $Script:ConfigPath -PathType Leaf)) { $lblStatus.Text = "Select config.ini first."; return }
    try {
        $n = Set-ConfigIniFromRows -ConfigPath $Script:ConfigPath -Rows $Script:AllRows -OnlySection $filter
        if ($filter) { $lblStatus.Text = "Updated $n keybinds in config.ini (section: $filter)." } else { $lblStatus.Text = "Updated $n keybinds in config.ini." }
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
        $result = New-CrimsonBackup -ConfigPath $Script:ConfigPath -WtfDir $Script:WtfDir -AccountName $account
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
