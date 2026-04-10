#Requires -Version 5.1
<#
.SYNOPSIS
  Standalone tool: config.ini keybind randomization + push to CrimsonBind WoW addon (SavedVariables).

.DESCRIPTION
  Companion to the CrimsonBind addon. Reads your hotkey profile config.ini (UTF-8 or UTF-16; encoding is preserved on save), optional bindable_keys.ini pool,
  randomizes keys (optional section filter), writes config.ini back, writes WTF\...\SavedVariables\CrimsonBind.lua.
  Import from WoW merges pendingEdits from SavedVariables back into config.ini.
  Load / Randomize / Sync merge macro bodies from CrimsonBind_binds.csv for every row that exists in the CSV (same Section + ActionName): CSV MacroText wins, including empty cells (clears profile placeholder text after ';' in config.ini that was mistaken for a macro).
  Target MemberN rows with empty MacroText get a default /focus raid/party macro on CSV import (same pattern as Paladin Holy). Non-healer specs: randomizer assigns keys for Target Member 1–10 only (11+ stay unbound); config.ini lines for those binds are updated only when a non-empty key exists (avoids wiping Debounce lines).
  Import CSV to WoW / GUI Import and Write-CrimsonBindCsvKeysFromSyncRows re-save the CSV with empty Target Member macros filled and all four columns quoted (Excel list-separator ';' no longer splits /focus lines into extra columns).
  Set-ConfigIniFromRows never writes ActionName=;macro (empty encoded key before ';') — that breaks Debounce/ggloader; rows with no Key in the sync set keep the previous config.ini line until a key exists.
  For [General] only: if the line on disk already has no key token before the first ';' (e.g. TargetMouseOver=; /target …), that line is left unchanged even when the CSV/sync rows have a Key, so a clean profile stays meta-slot-only. Pass -ForceGeneralIniKeys to Set-ConfigIniFromRows to push CSV keys into those lines anyway.
  Randomize for a class/spec excludes keys used under [General] from loaded rows and from non-empty General Key cells in the CSV (so single-section randomize stays consistent after reload).
  The key pool excludes any combo whose base is numpad-related (NUMPAD0–9, other NUMPAD* tokens, and nav synonyms like NUMPADDOWN), even if those appear under [NumpadKeys] or [ExplicitKeys].
  Modifier lines like CTRL-ALT- are joined to bases with an extra dash (CTRL-ALT-PAGEUP), not CTRL-ALTPAGEUP. INSERT/HOME/PAGEUP/arrow bases are excluded: they share Windows VKs with numpad navigation, so Debounce would label them NUMPADINS/NUMPADPGUP even without NUMPAD in the pool token.
  Clear-KeyIfNumpadRelated runs the full decode pipeline (vk/sc tokens, compact ^!+ forms, NUMPADHOME→NUMPAD7, etc.) then drops keys whose base is still numpad. It is used for rows, CSV import, General key reservation, ConvertTo-ConfigIniKey, and canonical comparison so encoded profile keys cannot round-trip back as numpad binds.
  For the same ActionName as a [General] row, class/spec rows use General's key (same rule as config.ini writes). After randomizing one section only, those rows are re-copied from General so they do not keep stale keys. Load / Sync / Import align in-memory spec keys to General before writing CrimsonBind.lua.
  Export CSV from config.ini writes Section, ActionName, decoded Key, and leaves MacroText empty (profile macro bodies after ';' are not dumped—avoid overwriting your real macro CSV). Use -IncludeProfileMacros on Export-CrimsonBindCsvFromConfigIni to embed stripped config macros again. Import CSV writes CrimsonBind.lua from CSV (macros = source of truth) and updates config.ini key tokens only (macro tails after ';' unchanged).
  After Sync to WoW and after Randomize+Sync, current bind keys are written into CrimsonBind_binds.csv (Key column only; MacroText unchanged) for rows whose Section+ActionName match, so [General] keys in the CSV stay accurate and spec randomize excludes them via Get-GeneralReservedKeyCanonSetFromCsv.
  Sync also rewrites config.ini bind lines (all sections) so compact/numpad profile keys stripped in memory are persisted — Debounce/other readers see the same keys as CrimsonBind.lua.
  [Main] WTFPath= in config.ini (relative paths are relative to the config.ini folder) is used to fill the
  WoW SavedVariables path as WTF\Account\<account>\SavedVariables (account from Config.wtf, or inferred).

.EXAMPLE
  .\Sync-CrimsonBinds.ps1

.EXAMPLE
  .\Sync-CrimsonBinds.ps1 -ImportCsvToWow
  Headless: read sync_paths.ini (or overrides), import CrimsonBind_binds.csv to WoW SavedVariables\CrimsonBind.lua and refresh config.ini keys.

#>

param(
    [switch] $ImportCsvToWow,
    [switch] $NoBackup,
    [string] $ConfigIni = "",
    [string] $WoWSavedVariables = "",
    [string] $CrimsonBindCsv = ""
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$ToolDir = $PSScriptRoot
$Script:SyncPathsIni = Join-Path $ToolDir "sync_paths.ini"
$Script:DefaultKeyPoolPath = Join-Path $ToolDir "bindable_keys.ini"
$Script:ConfigIniKeySuffix = "67701769"
# Placeholder vk token when Target Member lines are `=; /focus...` (ggloader ERROR on empty hotkey before ';').
$Script:ConfigIniTargetMemberEmptyKeyPlaceholder = "vk13_67701769"
$Script:SyncRows = $null
$Script:ConfigIniNonBindSections = @{ "Main" = $true; "SourceSettings" = $true; "PhysInput" = $true; "Druid_Settings" = $true }
$Script:ConfigIniKeybindValuePattern = '^[\^!+]*(vk|sc)[0-9a-fA-F]+_[0-9a-fA-F]+$'
$Script:ConfigIniDescriptionMacroPattern = '^For instructions on how to set up this hotkey'
$Script:ConfigIniExcludeActionNames = @{
    "Human Racial" = $true; "Stoneform" = $true; "Shadowmeld" = $true; "Escape Artist" = $true
    "Gift of the Naaru" = $true; "Darkflight" = $true; "Blood Fury" = $true; "Will of the Forsaken" = $true
    "War Stomp" = $true; "Berserking" = $true; "Arcane Torrent" = $true; "Rocket Jump" = $true
    "Rocket Barrage" = $true; "Quaking Palm" = $true; "Spatial Rift" = $true; "Light's Judgment" = $true
    "Fireblood" = $true; "Arcane Pulse" = $true; "Bull Rush" = $true; "Ancestral Call" = $true
    "Haymaker" = $true; "Regeneratin" = $true; "Bag of Tricks" = $true; "Hyper Organic Light Originator" = $true
    "Azerite Surge" = $true; "Sharpen Blade" = $true
}

$Script:ConfigIniKeyCodeMap = @{
    "F1"="vk70"; "F2"="vk71"; "F3"="vk72"; "F4"="vk73"; "F5"="vk74"; "F6"="vk75"; "F7"="vk76"; "F8"="vk77"; "F9"="vk78"; "F10"="vk79"; "F11"="vk7a"; "F12"="vk7b"
    "0"="vk30"; "1"="vk31"; "2"="vk32"; "3"="vk33"; "4"="vk34"; "5"="vk35"; "6"="vk36"; "7"="vk37"; "8"="vk38"; "9"="vk39"
    "MINUS"="vkbd"; "EQUALS"="vkbb"
    "A"="sc1e"; "B"="sc30"; "C"="sc2e"; "D"="sc20"; "E"="sc12"; "F"="sc21"; "G"="sc22"; "H"="sc23"; "I"="sc17"; "J"="sc24"; "K"="sc25"; "L"="sc26"; "M"="sc32"; "N"="sc31"; "O"="sc18"; "P"="sc19"; "Q"="sc10"; "R"="sc13"; "S"="sc1f"; "T"="sc14"; "U"="sc16"; "V"="sc2f"; "W"="sc11"; "X"="sc2d"; "Y"="sc15"; "Z"="sc2c"
    "LBRACKET"="vkdb"; "RBRACKET"="vkdd"; "BACKSLASH"="vkdc"; "SEMICOLON"="vkba"; "APOSTROPHE"="vkde"; "COMMA"="vkbc"; "PERIOD"="vkbe"; "SLASH"="vkbf"
    "NUMPAD0"="vk60"; "NUMPAD1"="sc4f"; "NUMPAD2"="sc50"; "NUMPAD3"="sc51"; "NUMPAD4"="vk64"; "NUMPAD5"="vk65"; "NUMPAD6"="vk66"; "NUMPAD7"="sc47"; "NUMPAD8"="vk68"; "NUMPAD9"="sc49"
    "NUMPADMULTIPLY"="vk6a"; "NUMPADPLUS"="vk6b"; "NUMPADMINUS"="vk6d"; "NUMPADDIVIDE"="vk6f"
    "TAB"="sc0f"
    "INSERT"="vk2d"; "HOME"="vk24"; "PAGEUP"="vk21"; "PAGEDOWN"="vk22"; "END"="vk23"
    "UP"="vk26"; "DOWN"="vk28"; "LEFT"="vk25"; "RIGHT"="vk27"
    "SC152"="sc152"
}
$Script:ConfigIniCodeToKeyName = @{}
foreach ($entry in $Script:ConfigIniKeyCodeMap.GetEnumerator()) { $Script:ConfigIniCodeToKeyName[$entry.Value] = $entry.Key }
$Script:IniKeyTokenToFriendlyChar = @{
    "LBRACKET" = "["; "RBRACKET" = "]"; "BACKSLASH" = "\"; "SEMICOLON" = ";"; "APOSTROPHE" = "'"; "COMMA" = ","; "PERIOD" = "."; "SLASH" = "/"
}
$Script:FriendlyCharToIniKeyToken = @{}
foreach ($kv in $Script:IniKeyTokenToFriendlyChar.GetEnumerator()) { $Script:FriendlyCharToIniKeyToken[$kv.Value] = $kv.Key }

# Profile/config may use numpad nav names (Num Lock off); WoW SetBinding expects NUMPAD0–NUMPAD9.
$Script:NumpadNavSynonymToNumpad = @{
    "NUMPADHOME"     = "NUMPAD7"
    "NUMPADEND"      = "NUMPAD1"
    "NUMPADUP"       = "NUMPAD8"
    "NUMPADDOWN"     = "NUMPAD2"
    "NUMPADLEFT"     = "NUMPAD4"
    "NUMPADRIGHT"    = "NUMPAD6"
    "NUMPADPGUP"     = "NUMPAD9"
    "NUMPADPGDN"     = "NUMPAD3"
    "NUMPADPAGEUP"   = "NUMPAD9"
    "NUMPADPAGEDOWN" = "NUMPAD3"
    "NUMPADINS"      = "NUMPAD0"
    "NUMPADINSERT"   = "NUMPAD0"
}

# INSERT/HOME/PAGEUP/etc. use the same Windows VKs as numpad navigation (Num Lock off); Debounce labels them NUMPADINS/NUMPADPGUP even when the pool meant the main keyboard.
$Script:RandomPoolExcludeBasesNumpadVkAlias = @{
    INSERT = $true; DELETE = $true; HOME = $true; END = $true
    PAGEUP = $true; PAGEDOWN = $true
    UP = $true; DOWN = $true; LEFT = $true; RIGHT = $true
}

# Numpad excluded from presets / randomizer pool (use [ExplicitKeys] in bindable_keys.ini if needed).
$Script:KeyPoolNumpadNumbers = @()
$Script:KeyPoolFull = @(
    "6", "7", "8", "9", "0", "MINUS", "EQUALS",
    "T", "Y", "U", "I", "O", "P", "LBRACKET", "RBRACKET", "BACKSLASH",
    "G", "H", "J", "K", "L", "SEMICOLON", "APOSTROPHE",
    "Z", "X", "C", "V", "B", "N", "M", "COMMA", "PERIOD", "SLASH",
    "TAB",
    "INSERT", "HOME", "PAGEUP", "PAGEDOWN", "END", "UP", "DOWN", "LEFT", "RIGHT"
)
$Script:ModsFull = @("CTRL-", "ALT-", "SHIFT-", "CTRL-ALT-", "CTRL-SHIFT-", "CTRL-ALT-SHIFT-")
$Script:KeyPoolRestoRightHand = @(
    "6", "7", "8", "9", "0", "MINUS", "EQUALS",
    "T", "Y", "U", "I", "O", "P", "LBRACKET", "RBRACKET", "BACKSLASH",
    "G", "H", "J", "K", "L", "SEMICOLON", "APOSTROPHE",
    "V", "B", "N", "M", "COMMA", "PERIOD", "SLASH",
    "INSERT", "HOME", "PAGEUP", "PAGEDOWN", "END", "UP", "DOWN", "LEFT", "RIGHT", "SC152"
)
$Script:ModsResto = @("CTRL-", "ALT-", "SHIFT-", "CTRL-ALT-", "CTRL-SHIFT-")

function Read-SyncPathsIni {
    $r = @{ ConfigIni = ""; WoWSavedVariables = ""; KeyPoolFile = ""; CrimsonBindCsv = "" }
    if (-not (Test-Path -LiteralPath $Script:SyncPathsIni -PathType Leaf)) { return $r }
    foreach ($line in (Get-Content -LiteralPath $Script:SyncPathsIni -Encoding UTF8 -ErrorAction SilentlyContinue)) {
        if ($null -eq $line) { continue }
        $t = $line.Trim()
        if (-not $t -or $t.StartsWith('#')) { continue }
        $eq = $t.IndexOf('=')
        if ($eq -lt 1) { continue }
        $k = $t.Substring(0, $eq).Trim()
        $v = $t.Substring($eq + 1).Trim()
        switch ($k) {
            "ConfigIni" { $r.ConfigIni = $v }
            "WoWSavedVariables" { $r.WoWSavedVariables = $v }
            "KeyPoolFile" { $r.KeyPoolFile = $v }
            "CrimsonBindCsv" { $r.CrimsonBindCsv = $v }
        }
    }
    return $r
}

function Write-SyncPathsIni {
    param([string]$ConfigIni, [string]$WoWSavedVariables, [string]$KeyPoolFile, [string]$CrimsonBindCsv)
    $sb = [System.Text.StringBuilder]::new()
    [void]$sb.AppendLine('# Sync-CrimsonBinds paths (safe to delete)')
    if ($ConfigIni) { [void]$sb.AppendLine("ConfigIni=$ConfigIni") }
    if ($WoWSavedVariables) { [void]$sb.AppendLine("WoWSavedVariables=$WoWSavedVariables") }
    if ($KeyPoolFile) { [void]$sb.AppendLine("KeyPoolFile=$KeyPoolFile") }
    if ($CrimsonBindCsv) { [void]$sb.AppendLine("CrimsonBindCsv=$CrimsonBindCsv") }
    [System.IO.File]::WriteAllText($Script:SyncPathsIni, $sb.ToString(), [System.Text.UTF8Encoding]::new($false))
}

function New-CrimsonBindBackup {
    param(
        [string]$ConfigPath,
        [string]$CsvPath,
        [string]$KeyPoolPath,
        [string]$SavedVarsDir
    )
    $backupsRoot = Join-Path $Script:ToolDir "BACKUPS"
    if (-not (Test-Path -LiteralPath $backupsRoot)) {
        [void](New-Item -Path $backupsRoot -ItemType Directory -Force)
    }
    $stamp = [DateTime]::Now.ToString("yyyy-MM-dd_HHmmss")
    $dir = Join-Path $backupsRoot $stamp
    $suffix = 2
    while (Test-Path -LiteralPath $dir) {
        $dir = Join-Path $backupsRoot "${stamp}_$suffix"
        $suffix++
    }
    [void](New-Item -Path $dir -ItemType Directory -Force)
    $manifest = @{}
    $copied = 0
    $filesToBackup = @(
        @{ Name = "config.ini";            Source = $ConfigPath }
        @{ Name = "CrimsonBind_binds.csv"; Source = $CsvPath }
        @{ Name = "bindable_keys.ini";     Source = $KeyPoolPath }
    )
    if ($SavedVarsDir) {
        $filesToBackup += @{ Name = "CrimsonBind.lua"; Source = (Join-Path $SavedVarsDir "CrimsonBind.lua") }
        $filesToBackup += @{ Name = "Debounce.lua";    Source = (Join-Path $SavedVarsDir "Debounce.lua") }
        $filesToBackup += @{ Name = "BindPad.lua";     Source = (Join-Path $SavedVarsDir "BindPad.lua") }
    }
    foreach ($f in $filesToBackup) {
        $src = $f.Source
        if (-not $src -or -not (Test-Path -LiteralPath $src -PathType Leaf)) { continue }
        $dest = Join-Path $dir $f.Name
        Copy-Item -LiteralPath $src -Destination $dest -Force
        $manifest[$f.Name] = [System.IO.Path]::GetFullPath($src)
        $copied++
    }
    if ($copied -eq 0) {
        Remove-Item -LiteralPath $dir -Recurse -Force -ErrorAction SilentlyContinue
        return $null
    }
    $json = $manifest | ConvertTo-Json -Depth 1
    [System.IO.File]::WriteAllText((Join-Path $dir "manifest.json"), $json, [System.Text.UTF8Encoding]::new($false))
    return $dir
}

function Restore-CrimsonBindBackup {
    param(
        [Parameter(Mandatory)][string]$BackupDir,
        [string[]]$FileNames
    )
    $mfPath = Join-Path $BackupDir "manifest.json"
    if (-not (Test-Path -LiteralPath $mfPath -PathType Leaf)) {
        throw "No manifest.json in $BackupDir"
    }
    $manifest = Get-Content -LiteralPath $mfPath -Raw | ConvertFrom-Json
    $restored = 0
    foreach ($name in $FileNames) {
        $src = Join-Path $BackupDir $name
        if (-not (Test-Path -LiteralPath $src -PathType Leaf)) { continue }
        $dest = $null
        foreach ($p in $manifest.PSObject.Properties) {
            if ($p.Name -eq $name) { $dest = $p.Value; break }
        }
        if (-not $dest) { continue }
        $destDir = [System.IO.Path]::GetDirectoryName($dest)
        if ($destDir -and -not (Test-Path -LiteralPath $destDir)) {
            [void](New-Item -Path $destDir -ItemType Directory -Force)
        }
        Copy-Item -LiteralPath $src -Destination $dest -Force
        $restored++
    }
    return $restored
}

function Show-RestoreBackupDialog {
    param([string]$BackupsRoot, [System.Windows.Forms.Label]$StatusLabel)
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = "Restore Backup"
    $dlg.Size = [System.Drawing.Size]::new(560, 380)
    $dlg.StartPosition = "CenterParent"
    $dlg.FormBorderStyle = "FixedDialog"
    $dlg.MaximizeBox = $false
    $dlg.MinimizeBox = $false

    $lblFolder = New-Object System.Windows.Forms.Label
    $lblFolder.Text = "Backup folder:"
    $lblFolder.Location = [System.Drawing.Point]::new(12, 14)
    $lblFolder.Size = [System.Drawing.Size]::new(85, 20)
    $dlg.Controls.Add($lblFolder)

    $tbFolder = New-Object System.Windows.Forms.TextBox
    $tbFolder.Location = [System.Drawing.Point]::new(100, 12)
    $tbFolder.Size = [System.Drawing.Size]::new(350, 22)
    $tbFolder.ReadOnly = $true
    $dlg.Controls.Add($tbFolder)

    $btnBrowse = New-Object System.Windows.Forms.Button
    $btnBrowse.Text = "..."
    $btnBrowse.Location = [System.Drawing.Point]::new(456, 10)
    $btnBrowse.Size = [System.Drawing.Size]::new(75, 26)
    $dlg.Controls.Add($btnBrowse)

    $pnlChecks = New-Object System.Windows.Forms.Panel
    $pnlChecks.Location = [System.Drawing.Point]::new(12, 46)
    $pnlChecks.Size = [System.Drawing.Size]::new(520, 200)
    $pnlChecks.AutoScroll = $true
    $dlg.Controls.Add($pnlChecks)

    $btnRestore = New-Object System.Windows.Forms.Button
    $btnRestore.Text = "Restore Selected"
    $btnRestore.Location = [System.Drawing.Point]::new(280, 258)
    $btnRestore.Size = [System.Drawing.Size]::new(120, 30)
    $btnRestore.Enabled = $false
    $dlg.Controls.Add($btnRestore)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Cancel"
    $btnCancel.Location = [System.Drawing.Point]::new(410, 258)
    $btnCancel.Size = [System.Drawing.Size]::new(80, 30)
    $dlg.Controls.Add($btnCancel)

    $lblInfo = New-Object System.Windows.Forms.Label
    $lblInfo.Location = [System.Drawing.Point]::new(12, 298)
    $lblInfo.Size = [System.Drawing.Size]::new(520, 40)
    $lblInfo.Text = "Select a backup folder to see available files."
    $dlg.Controls.Add($lblInfo)

    $script:checkBoxes = @()

    $populateChecks = {
        param([string]$dir)
        $pnlChecks.Controls.Clear()
        $script:checkBoxes = @()
        $btnRestore.Enabled = $false
        $mfPath = Join-Path $dir "manifest.json"
        if (-not (Test-Path -LiteralPath $mfPath -PathType Leaf)) {
            $lblInfo.Text = "No manifest.json found in selected folder."
            return
        }
        try { $mf = Get-Content -LiteralPath $mfPath -Raw | ConvertFrom-Json } catch {
            $lblInfo.Text = "Failed to read manifest: $($_.Exception.Message)"
            return
        }
        $y = 4
        foreach ($p in $mf.PSObject.Properties) {
            $name = $p.Name
            $origPath = $p.Value
            $backupFile = Join-Path $dir $name
            $exists = Test-Path -LiteralPath $backupFile -PathType Leaf
            $cb = New-Object System.Windows.Forms.CheckBox
            $cb.Text = "$name  ->  $origPath"
            $cb.Location = [System.Drawing.Point]::new(4, $y)
            $cb.Size = [System.Drawing.Size]::new(500, 22)
            $cb.Checked = $exists
            $cb.Enabled = $exists
            $cb.Tag = $name
            $pnlChecks.Controls.Add($cb)
            $script:checkBoxes += $cb
            $y += 26
        }
        $btnRestore.Enabled = ($script:checkBoxes | Where-Object { $_.Checked }).Count -gt 0
        $lblInfo.Text = "Select files to restore. A new backup of current files will be created first."
    }

    $btnBrowse.Add_Click({
        $d = New-Object System.Windows.Forms.FolderBrowserDialog
        $d.Description = "Select a backup folder inside BACKUPS"
        if ($BackupsRoot -and (Test-Path -LiteralPath $BackupsRoot)) {
            $d.SelectedPath = $BackupsRoot
        }
        if ($d.ShowDialog() -eq "OK") {
            $tbFolder.Text = $d.SelectedPath
            & $populateChecks $d.SelectedPath
        }
    })

    $btnCancel.Add_Click({ $dlg.Close() })

    $btnRestore.Add_Click({
        $dir = $tbFolder.Text.Trim()
        if (-not $dir) { return }
        $selected = @()
        foreach ($cb in $script:checkBoxes) {
            if ($cb.Checked -and $cb.Enabled) { $selected += $cb.Tag }
        }
        if ($selected.Count -eq 0) {
            $lblInfo.Text = "No files selected."
            return
        }
        try {
            $mfPath2 = Join-Path $dir "manifest.json"
            $mf2 = Get-Content -LiteralPath $mfPath2 -Raw | ConvertFrom-Json
            $cfgSrc = $null; $csvSrc = $null; $poolSrc = $null; $svDir = $null
            foreach ($p in $mf2.PSObject.Properties) {
                switch ($p.Name) {
                    "config.ini"            { $cfgSrc = $p.Value }
                    "CrimsonBind_binds.csv" { $csvSrc = $p.Value }
                    "bindable_keys.ini"     { $poolSrc = $p.Value }
                    "Debounce.lua"          { $svDir = [System.IO.Path]::GetDirectoryName($p.Value) }
                    "BindPad.lua"           { if (-not $svDir) { $svDir = [System.IO.Path]::GetDirectoryName($p.Value) } }
                }
            }
            $preBackup = New-CrimsonBindBackup -ConfigPath $cfgSrc -CsvPath $csvSrc -KeyPoolPath $poolSrc -SavedVarsDir $svDir
            $preNote = if ($preBackup) { "Pre-restore backup: $(Split-Path -Leaf $preBackup). " } else { "" }
            $count = Restore-CrimsonBindBackup -BackupDir $dir -FileNames $selected
            $lblInfo.Text = "${preNote}Restored $count file(s)."
            if ($StatusLabel) { $StatusLabel.Text = "${preNote}Restored $count file(s) from $(Split-Path -Leaf $dir)." }
        } catch {
            $lblInfo.Text = "Restore failed: $($_.Exception.Message)"
        }
    })

    [void]$dlg.ShowDialog()
}

function Get-CrimsonBindCsvDefaultPath {
    param([string]$ConfigPath)
    if ($ConfigPath -and (Test-Path -LiteralPath $ConfigPath -PathType Leaf)) {
        $dir = [System.IO.Path]::GetDirectoryName($ConfigPath)
        if ($dir) {
            return [System.IO.Path]::GetFullPath([System.IO.Path]::Combine($dir, "CrimsonBind_binds.csv"))
        }
    }
    return [System.IO.Path]::GetFullPath((Join-Path $ToolDir "CrimsonBind_binds.csv"))
}

function Get-MacroTextForCrimsonBindCsvExport {
    param([string]$Raw, [switch]$EmptyOnly)
    if ($EmptyOnly) { return "" }
    if (-not $Raw) { return "" }
    $t = $Raw.Trim()
    if ($t -match '§§§') {
        $t = ($t -split '§§§', 2)[0].Trim()
    }
    if ($t -match '§§') {
        $t = ($t -split '§§', 2)[0].Trim()
    }
    return $t
}

# config.ini / CSV often store multi-line macros as §-separated lines; WoW needs real newlines. Strip §§§ then §§ note tails (§§Note: …) before § → newline.
function Normalize-MacroTextForCrimsonBindLua {
    param([string]$Raw)
    if ($null -eq $Raw -or $Raw -eq "") { return "" }
    $t = $Raw.Trim()
    if ($t -match '§§§') {
        $t = ($t -split '§§§', 2)[0].Trim()
    }
    if ($t -match '§§') {
        $t = ($t -split '§§', 2)[0].Trim()
    }
    $t = $t -replace '§', "`n"
    $t = $t -replace "`r`n", "`n" -replace "`r", "`n"
    return $t.Trim()
}

# CSV often leaves Target MemberN blank for non-healer specs; CrimsonBind skips empty macros in getBindsToApply, so those rows never bind.
# Same /focus pattern as Paladin Holy rows in CrimsonBind_binds.csv (party branch for 1–9, TM10 uses partypet4).
function Get-DefaultMacroTextForTargetMemberIndex {
    param([int]$Index)
    if ($Index -lt 1) { return "" }
    if ($Index -le 9) {
        return "/focus [mod:ctrl]raidpet$Index; [mod:alt]party$Index; raid$Index"
    }
    if ($Index -eq 10) {
        return "/focus [mod:ctrl]raidpet10; [mod:alt]partypet4; raid10"
    }
    return "/focus [mod:ctrl]raidpet$Index; raid$Index"
}

function Export-CrimsonBindCsvFromConfigIni {
    param(
        [Parameter(Mandatory)][string]$ConfigPath,
        [Parameter(Mandatory)][string]$CsvPath,
        [switch]$IncludeProfileMacros
    )
    $rows = Get-SyncConfigIniRows -ConfigPath $ConfigPath
    if ($rows.Count -eq 0) {
        throw "No bind rows in config.ini (nothing to export)."
    }
    Update-SyncRowsDecodedKeys -Rows $rows
    $export = foreach ($r in @($rows)) {
        if ($null -eq $r) { continue }
        $k = if ($null -eq $r.Key) { "" } else { $r.Key.ToString().Trim() }
        $mt = if ($IncludeProfileMacros) {
            Get-MacroTextForCrimsonBindCsvExport -Raw ($r.MacroText)
        } else {
            ""
        }
        [PSCustomObject]@{
            Section    = $r.Section
            ActionName = $r.ActionName
            MacroText  = $mt
            Key        = $k
        }
    }
    $dir = [System.IO.Path]::GetDirectoryName($CsvPath)
    if ($dir -and -not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    $export | Select-Object Section, ActionName, MacroText, Key | Export-Csv -LiteralPath $CsvPath -NoTypeInformation -Encoding UTF8
}

function Import-CsvRowProperty {
    param([psobject]$Row, [string[]]$Names)
    foreach ($n in $Names) {
        foreach ($p in $Row.PSObject.Properties) {
            if ($p.Name -ieq $n) {
                if ($null -eq $p.Value) { return "" }
                return $p.Value.ToString()
            }
        }
    }
    return ""
}

function Import-CrimsonBindCsvToRows {
    param([Parameter(Mandatory)][string]$CsvPath)
    if (-not (Test-Path -LiteralPath $CsvPath -PathType Leaf)) {
        throw "CSV not found: $CsvPath"
    }
    $tbl = @(Import-Csv -LiteralPath $CsvPath -Encoding UTF8)
    $out = [System.Collections.ArrayList]::new()
    foreach ($row in $tbl) {
        if ($null -eq $row) { continue }
        $sec = (Import-CsvRowProperty -Row $row -Names @('Section')).Trim()
        $an = (Import-CsvRowProperty -Row $row -Names @('ActionName')).Trim()
        $mtRaw = Import-CsvRowProperty -Row $row -Names @('MacroText')
        $mt = Normalize-MacroTextForCrimsonBindLua -Raw $mtRaw
        $key = Clear-KeyIfNumpadRelated -Key ((Import-CsvRowProperty -Row $row -Names @('Key')).Trim())
        if (-not $an) { continue }
        if ($an -match '^Target Member(\d+)$' -and [string]::IsNullOrWhiteSpace($mt)) {
            $idx = 0
            if ([int]::TryParse($Matches[1], [ref]$idx)) {
                $mt = Get-DefaultMacroTextForTargetMemberIndex -Index $idx
            }
        }
        [void]$out.Add([PSCustomObject]@{
            Section = $sec; ActionName = $an; Key = $key; MacroText = $mt; TextureID = "132089"
        })
    }
    return $out
}

# PowerShell Export-Csv often leaves fields unquoted; MacroText with ';' then breaks Excel (and some validators) when list separator is ';'. Always quote all four columns.
function Export-CrimsonBindCsvTableUtf8Quoted {
    param(
        [Parameter(Mandatory)] $Rows,
        [Parameter(Mandatory)][string]$Path
    )
    function Escape-CsvField([string]$s) {
        if ($null -eq $s) { $s = "" }
        return '"' + ($s.ToString() -replace '"', '""') + '"'
    }
    $enc = [System.Text.UTF8Encoding]::new($false)
    $sb = [System.Text.StringBuilder]::new()
    [void]$sb.AppendLine(@(
            (Escape-CsvField 'Section')
            (Escape-CsvField 'ActionName')
            (Escape-CsvField 'MacroText')
            (Escape-CsvField 'Key')
        ) -join ',')
    foreach ($row in @($Rows)) {
        if ($null -eq $row) { continue }
        $sec = Import-CsvRowProperty -Row $row -Names @('Section')
        $an = Import-CsvRowProperty -Row $row -Names @('ActionName')
        $mt = Import-CsvRowProperty -Row $row -Names @('MacroText')
        $ky = Import-CsvRowProperty -Row $row -Names @('Key')
        [void]$sb.AppendLine(@(
                (Escape-CsvField $sec)
                (Escape-CsvField $an)
                (Escape-CsvField $mt)
                (Escape-CsvField $ky)
            ) -join ',')
    }
    [System.IO.File]::WriteAllText($Path, $sb.ToString().TrimEnd() + "`r`n", $enc)
}

function Apply-TargetMemberDefaultMacrosToCsvTable {
    param([Parameter(Mandatory)] $Tbl)
    $n = 0
    foreach ($row in @($Tbl)) {
        if ($null -eq $row) { continue }
        $an = (Import-CsvRowProperty -Row $row -Names @('ActionName')).Trim()
        if ($an -notmatch '^Target Member(\d+)$') { continue }
        $mtRaw = Import-CsvRowProperty -Row $row -Names @('MacroText')
        if (-not [string]::IsNullOrWhiteSpace($mtRaw)) { continue }
        $idx = 0
        if (-not [int]::TryParse($Matches[1], [ref]$idx)) { continue }
        $newM = Get-DefaultMacroTextForTargetMemberIndex -Index $idx
        $set = $false
        foreach ($p in $row.PSObject.Properties) {
            if ($p.Name -ieq 'MacroText') {
                $p.Value = $newM
                $set = $true
                break
            }
        }
        if (-not $set) {
            $row | Add-Member -NotePropertyName MacroText -NotePropertyValue $newM -Force
        }
        $n++
    }
    return $n
}

function Update-CrimsonBindCsvFileTargetMemberDefaultMacros {
    param([Parameter(Mandatory)][string]$CsvPath)
    if (-not (Test-Path -LiteralPath $CsvPath -PathType Leaf)) { return 0 }
    try {
        $tbl = @(Import-Csv -LiteralPath $CsvPath -Encoding UTF8 -ErrorAction Stop)
    } catch {
        return 0
    }
    if ($tbl.Count -eq 0) { return 0 }
    $n = Apply-TargetMemberDefaultMacrosToCsvTable -Tbl $tbl
    Export-CrimsonBindCsvTableUtf8Quoted -Rows $tbl -Path $CsvPath
    return $n
}

# CSV is the macro source of truth for any bind that appears in the CSV: overwrite row MacroText to match (including empty, to drop profile help text loaded from config.ini after ';').
function Merge-CrimsonBindCsvMacrosIntoRows {
    param(
        [System.Collections.ArrayList]$Rows,
        [string]$CsvPath
    )
    if (-not $Rows -or $Rows.Count -eq 0) { return 0 }
    if (-not $CsvPath -or -not (Test-Path -LiteralPath $CsvPath -PathType Leaf)) { return 0 }
    try {
        $csvRows = @(Import-CrimsonBindCsvToRows -CsvPath $CsvPath)
    } catch {
        return 0
    }
    if ($csvRows.Count -eq 0) { return 0 }
    $bySecAct = @{}
    foreach ($cr in $csvRows) {
        if ($null -eq $cr) { continue }
        $s = if ($null -eq $cr.Section) { "" } else { $cr.Section.ToString().Trim() }
        $a = if ($null -eq $cr.ActionName) { "" } else { $cr.ActionName.ToString().Trim() }
        if (-not $a) { continue }
        $mt = if ($null -eq $cr.MacroText) { "" } else { $cr.MacroText.ToString().Trim() }
        $bySecAct["$s|$a"] = $mt
    }
    $changed = 0
    foreach ($r in $Rows) {
        if ($null -eq $r) { continue }
        $s = if ($null -eq $r.Section) { "" } else { $r.Section.ToString().Trim() }
        $a = if ($null -eq $r.ActionName) { "" } else { $r.ActionName.ToString().Trim() }
        $lk = "$s|$a"
        if (-not $bySecAct.ContainsKey($lk)) { continue }
        $newM = $bySecAct[$lk]
        $cur = if ($null -eq $r.MacroText) { "" } else { $r.MacroText.ToString() }
        if ($cur -ceq $newM) { continue }
        $r.MacroText = $newM
        $changed++
    }
    return $changed
}

# Overwrite Key column in CSV from in-memory sync rows (Section+ActionName match). Fills empty Target Member* MacroText on disk (same defaults as Import-CrimsonBindCsvToWow). Writes UTF-8 CSV with all fields quoted (safe for Excel when MacroText contains ';').
function Write-CrimsonBindCsvKeysFromSyncRows {
    param(
        [System.Collections.ArrayList]$Rows,
        [Parameter(Mandatory)][string]$CsvPath
    )
    if (-not $Rows -or $Rows.Count -eq 0) { return 0 }
    if (-not (Test-Path -LiteralPath $CsvPath -PathType Leaf)) { return 0 }
    $map = @{}
    foreach ($r in @($Rows)) {
        if ($null -eq $r) { continue }
        $s = if ($null -eq $r.Section) { "" } else { $r.Section.ToString().Trim() }
        $a = if ($null -eq $r.ActionName) { "" } else { $r.ActionName.ToString().Trim() }
        if (-not $a) { continue }
        $k = if ($null -eq $r.Key) { "" } else { $r.Key.ToString().Trim() }
        $map["$s|$a"] = $k
    }
    try {
        $tbl = @(Import-Csv -LiteralPath $CsvPath -Encoding UTF8 -ErrorAction Stop)
    } catch {
        return 0
    }
    if ($tbl.Count -eq 0) { return 0 }
    $changed = 0
    if ($map.Count -gt 0) {
        foreach ($row in $tbl) {
            if ($null -eq $row) { continue }
            $sec = (Import-CsvRowProperty -Row $row -Names @('Section')).Trim()
            $an = (Import-CsvRowProperty -Row $row -Names @('ActionName')).Trim()
            if (-not $an) { continue }
            $lk = "$sec|$an"
            if (-not $map.ContainsKey($lk)) { continue }
            $newK = $map[$lk]
            if ($null -eq $newK) { $newK = "" }
            $oldK = (Import-CsvRowProperty -Row $row -Names @('Key')).Trim()
            if ($oldK -ceq $newK) { continue }
            $set = $false
            foreach ($p in $row.PSObject.Properties) {
                if ($p.Name -ieq 'Key') {
                    $p.Value = $newK
                    $set = $true
                    break
                }
            }
            if (-not $set) {
                $row | Add-Member -NotePropertyName Key -NotePropertyValue $newK -Force
            }
            $changed++
        }
    }
    $tmFill = Apply-TargetMemberDefaultMacrosToCsvTable -Tbl $tbl
    if ($changed -gt 0 -or $tmFill -gt 0) {
        Export-CrimsonBindCsvTableUtf8Quoted -Rows $tbl -Path $CsvPath
    }
    return $changed
}

# Resolve WTFPath= from [Main] (same idea as CrimsonBind.ps1). Relative paths are relative to config.ini's folder.
function Resolve-WtfPathRelativeToConfigIni {
    param([string]$ConfigPath, [string]$RawWtfPath)
    if (-not $RawWtfPath) { return $null }
    $t = $RawWtfPath.Trim().Trim('"')
    if (-not $t) { return $null }
    $cfgDir = [System.IO.Path]::GetDirectoryName($ConfigPath)
    if (-not $cfgDir) { return $null }
    if ([System.IO.Path]::IsPathRooted($t)) {
        return [System.IO.Path]::GetFullPath($t)
    }
    return [System.IO.Path]::GetFullPath((Join-Path $cfgDir $t))
}

function Get-WtfDirFromConfigIni {
    param([string]$ConfigPath)
    if (-not (Test-Path -LiteralPath $ConfigPath -PathType Leaf)) { return $null }
    try {
        $content = (Read-ConfigIniFileText -Path $ConfigPath) -split "`r?`n"
    } catch { return $null }
    $inMain = $false
    foreach ($line in $content) {
        if ($null -eq $line) { continue }
        $t = $line.Trim()
        if ($t -eq "[Main]") { $inMain = $true; continue }
        if ($inMain -and $t -match '^\[') { break }
        if ($inMain -and $t -match '^WTFPath=(.+)') {
            $full = Resolve-WtfPathRelativeToConfigIni -ConfigPath $ConfigPath -RawWtfPath $Matches[1].Trim()
            if (-not $full) { return $null }
            if ([System.IO.Path]::GetExtension($full) -ieq '.wtf') {
                return [System.IO.Path]::GetDirectoryName($full)
            }
            return $full.TrimEnd([char[]]@('\', '/'))
        }
    }
    return $null
}

function Get-AccountNameFromConfigWtf {
    param([string]$ConfigWtfPath)
    if (-not $ConfigWtfPath -or -not (Test-Path -LiteralPath $ConfigWtfPath -PathType Leaf)) { return $null }
    try {
        $content = Get-Content -LiteralPath $ConfigWtfPath -Encoding Unicode -ErrorAction Stop
    } catch { return $null }
    foreach ($line in $content) {
        if ($null -eq $line) { continue }
        if ($line -match 'SET\s+accountName\s+"([^"]*)"') {
            $name = $Matches[1].Trim()
            if ($name) { return $name }
        }
    }
    return $null
}

# WTF\Account\<name>\SavedVariables — account from Config.wtf, else single Account subfolder, else first with CrimsonBind.lua.
function Get-SavedVariablesDirFromConfigIni {
    param([string]$ConfigPath)
    $wtfDir = Get-WtfDirFromConfigIni -ConfigPath $ConfigPath
    if (-not $wtfDir -or -not (Test-Path -LiteralPath $wtfDir -PathType Container)) { return $null }
    $configWtf = Join-Path $wtfDir "Config.wtf"
    $acct = Get-AccountNameFromConfigWtf -ConfigWtfPath $configWtf
    $acctRoot = Join-Path $wtfDir "Account"
    if (-not (Test-Path -LiteralPath $acctRoot -PathType Container)) { return $null }
    if (-not $acct) {
        $dirs = @(Get-ChildItem -LiteralPath $acctRoot -Directory -ErrorAction SilentlyContinue | Sort-Object Name)
        if ($dirs.Count -eq 0) { return $null }
        if ($dirs.Count -eq 1) {
            $acct = $dirs[0].Name
        } else {
            $withCb = $dirs | Where-Object { Test-Path -LiteralPath (Join-Path $_.FullName "SavedVariables\CrimsonBind.lua") } | Select-Object -First 1
            $acct = if ($withCb) { $withCb.Name } else { $dirs[0].Name }
        }
    }
    $sv = Join-Path $wtfDir "Account\$acct\SavedVariables"
    if (Test-Path -LiteralPath $sv -PathType Container) { return $sv }
    return $null
}

# config.ini key tokens use ^ ! + mod prefixes (no CTRL-/ALT- dashes) before vk/sc codes or key names.
function Expand-ConfigIniCompactMods {
    param([string]$Key)
    $k = ($Key -replace "^\s+|\s+$", "")
    if (-not $k) { return $Key }
    if ($k.IndexOf('-') -ge 0) { return $Key }
    if ($k -match $Script:ConfigIniKeybindValuePattern) { return $Key }
    $i = 0
    $mods = [System.Collections.ArrayList]::new()
    $len = $k.Length
    while ($i -lt $len) {
        $c = $k[$i]
        if ($c -eq [char]'^') { [void]$mods.Add('CTRL'); $i++; continue }
        if ($c -eq [char]'!') { [void]$mods.Add('ALT'); $i++; continue }
        if ($c -eq [char]'+') { [void]$mods.Add('SHIFT'); $i++; continue }
        break
    }
    if ($i -eq 0) { return $Key }
    $base = $k.Substring($i).Trim()
    if (-not $base) { return $Key }
    if ($base -ceq '=') { $base = 'EQUALS' }
    elseif ($base -ceq '-') { $base = 'MINUS' }
    $sorted = @($mods | Sort-Object)
    return (Join-BindKeyFromModsAndBase -Mods $sorted -Base $base)
}

function Normalize-NumpadNavSynonymsInBindKey {
    param([string]$Key)
    if (-not $Key) { return $Key }
    $pb = Split-BindKeyIntoModsAndBase -Key $Key
    $base = if ($pb.Base) { $pb.Base.Trim() } else { "" }
    if (-not $base) { return $Key }
    $u = $base.ToUpperInvariant()
    $repl = $Script:NumpadNavSynonymToNumpad[$u]
    if (-not $repl) { return $Key }
    return (Join-BindKeyFromModsAndBase -Mods $pb.Mods -Base $repl)
}

function Normalize-BindKeyPipeline {
    param([string]$Key)
    $kp = ($Key -replace "^\s+|\s+$", "")
    if (-not $kp) { return $kp }
    if ($kp -match $Script:ConfigIniKeybindValuePattern) {
        $kp = ConvertFrom-ConfigIniKey -ConfigKey $kp
    }
    $kp = Expand-ConfigIniCompactMods -Key $kp
    $kp = Normalize-NumpadNavSynonymsInBindKey -Key $kp
    return $kp
}

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
        if ($hex.Length -eq 1 -and ($prefix -eq "vk" -or $prefix -eq "sc")) {
            $code = $prefix + "0" + $hex
            $keyName = $Script:ConfigIniCodeToKeyName[$code]
        }
    }
    if (-not $keyName) { return $k }
    $modStr = ""
    if ($modPrefix -match '\^') { $modStr += "CTRL-" }
    if ($modPrefix -match '!') { $modStr += "ALT-" }
    if ($modPrefix -match '\+') { $modStr += "SHIFT-" }
    return $modStr + $keyName
}

function Split-BindKeyIntoModsAndBase {
    param([string]$Key)
    $k = ($Key -replace "^\s+|\s+$", "")
    if (-not $k) { return @{ Mods = @(); Base = "" } }
    $parts = $k -split '-'
    $mods = [System.Collections.ArrayList]::new()
    $i = 0
    for (; $i -lt $parts.Count; $i++) {
        $seg = $parts[$i]
        if ($null -eq $seg) { continue }
        $t = $seg.Trim().ToUpperInvariant()
        if ($t -eq 'CTRL' -or $t -eq 'ALT' -or $t -eq 'SHIFT') { [void]$mods.Add($t) }
        else { break }
    }
    $base = if ($i -lt $parts.Count) { ($parts[$i..($parts.Count - 1)] -join '-').Trim() } else { "" }
    return @{ Mods = @($mods); Base = $base }
}

function Join-BindKeyFromModsAndBase {
    param([object[]]$Mods, [string]$Base)
    $mb = if ($Base) { $Base.Trim() } else { "" }
    if ($Mods -and $Mods.Count -gt 0) {
        $m = ($Mods | ForEach-Object { $_.ToString() }) -join '-'
        if ($mb) { return "$m-$mb" }
        return $m
    }
    return $mb
}

function ConvertTo-ConfigIniKey {
    param([string]$Key, [string]$Suffix = $Script:ConfigIniKeySuffix)
    if ($null -eq $Key) { return "" }
    $k = ($Key -replace "^\s+|\s+$", "")
    if (-not $k) { return "" }
    $k = Clear-KeyIfNumpadRelated -Key $k
    if (-not $k) { return "" }
    $parts = $k -split '-'
    $hasCtrl = $false; $hasAlt = $false; $hasShift = $false
    $base = $null
    foreach ($p in $parts) {
        if ($null -eq $p) { continue }
        $t = $p.Trim().ToUpperInvariant()
        if ($t -eq 'CTRL') { $hasCtrl = $true } elseif ($t -eq 'ALT') { $hasAlt = $true } elseif ($t -eq 'SHIFT') { $hasShift = $true } else { $base = $p.Trim(); break }
    }
    if (-not $base -and $parts.Count -gt 0) {
        $lastSeg = $parts[$parts.Count - 1]
        if ($null -ne $lastSeg) { $base = $lastSeg.ToString().Trim() } else { $base = "" }
    }
    if (-not $base) { return $k }
    $lookupBase = $base.ToUpperInvariant()
    $navRepl = $Script:NumpadNavSynonymToNumpad[$lookupBase]
    if ($navRepl) {
        $lookupBase = $navRepl.ToUpperInvariant()
        $base = $navRepl
    }
    if ($base.Length -eq 1 -and $Script:FriendlyCharToIniKeyToken.ContainsKey($base)) {
        $lookupBase = $Script:FriendlyCharToIniKeyToken[$base]
    }
    $code = $Script:ConfigIniKeyCodeMap[$lookupBase]
    if (-not $code) { $code = $Script:ConfigIniKeyCodeMap[$base] }
    if (-not $code) { $code = $Script:ConfigIniKeyCodeMap[$base.ToUpperInvariant()] }
    if (-not $code) { return $k }
    $modStr = ""; if ($hasCtrl) { $modStr += "^" }; if ($hasAlt) { $modStr += "!" }; if ($hasShift) { $modStr += "+" }
    return $modStr + $code + "_" + $Suffix
}

function Get-BindKeyCanonicalConfigForm {
    param([string]$Key)
    if (-not $Key) { return "" }
    $t = Clear-KeyIfNumpadRelated -Key $Key.Trim()
    if (-not $t) { return "" }
    $enc = ConvertTo-ConfigIniKey -Key $t -Suffix $Script:ConfigIniKeySuffix
    if ($enc -match '_[0-9a-fA-F]+$' -and ($enc -match 'vk[0-9a-fA-F]+' -or $enc -match 'sc[0-9a-fA-F]+')) {
        return $enc.ToLowerInvariant()
    }
    return $t.ToUpperInvariant()
}

function Test-BindKeyUsesAltShift {
    param([string]$Key)
    if (-not $Key) { return $false }
    $t = $Key.Trim()
    if (-not $t) { return $false }
    if ($t -match '_[0-9a-fA-F]+$' -and ($t -match 'vk[0-9a-fA-F]+' -or $t -match 'sc[0-9a-fA-F]+')) {
        if ($t -match '^([\^!+]*)(vk|sc)') {
            $mods = if ($Matches.Count -gt 1 -and $null -ne $Matches[1]) { [string]$Matches[1] } else { "" }
            return ($mods.IndexOf([char]'!') -ge 0) -and ($mods.IndexOf([char]'+') -ge 0)
        }
        return $false
    }
    $hasAlt = $false; $hasShift = $false
    foreach ($p in ($t -split '-')) {
        if ($null -eq $p) { continue }
        $u = $p.Trim().ToUpperInvariant()
        if ($u -eq 'ALT') { $hasAlt = $true }
        if ($u -eq 'SHIFT') { $hasShift = $true }
    }
    return $hasAlt -and $hasShift
}

function Test-BindKeyUsesCtrlOrAlt {
    param([string]$Key)
    if (-not $Key) { return $false }
    $t = $Key.Trim()
    if (-not $t) { return $false }
    if ($t -match '_[0-9a-fA-F]+$' -and ($t -match 'vk[0-9a-fA-F]+' -or $t -match 'sc[0-9a-fA-F]+')) {
        if ($t -match '^([\^!+]*)(vk|sc)') {
            $mods = if ($Matches.Count -gt 1 -and $null -ne $Matches[1]) { [string]$Matches[1] } else { "" }
            return ($mods.IndexOf([char]'^') -ge 0) -or ($mods.IndexOf([char]'!') -ge 0)
        }
        return $false
    }
    foreach ($p in ($t -split '-')) {
        if ($null -eq $p) { continue }
        $u = $p.Trim().ToUpperInvariant()
        if ($u -eq 'CTRL' -or $u -eq 'ALT') { return $true }
    }
    return $false
}

function Test-ActionNameIsTargetMember1Through10 {
    param([string]$ActionName)
    if (-not $ActionName) { return $false }
    return $ActionName.Trim() -match '^Target Member(10|[1-9])$'
}

function Test-SectionStringIsHealerHealingSpec {
    param([string]$Section)
    if (-not $Section) { return $false }
    $s = $Section.Trim()
    if ($s -ieq "General") { return $false }
    if ($s -notmatch '^\s*(.+?)\s+-\s+(.+?)\s*$') { return $false }
    $class = ($Matches[1].Trim() -replace '\s+', '').ToUpperInvariant()
    $spec = $Matches[2].Trim()
    switch ($class) {
        "PALADIN" { return $spec -ieq "Holy" }
        "PRIEST" { return $spec -imatch '^(Discipline|Holy)$' }
        "DRUID" { return $spec -ieq "Restoration" }
        "SHAMAN" { return $spec -ieq "Restoration" }
        "MONK" { return $spec -ieq "Mistweaver" }
        "EVOKER" { return $spec -ieq "Preservation" }
        default { return $false }
    }
}

function Normalize-BindKeyDisplayFromTokens {
    param([string]$Key)
    if (-not $Key) { return $Key }
    $t = Clear-KeyIfNumpadRelated -Key $Key.Trim()
    if (-not $t) { return "" }
    $enc = ConvertTo-ConfigIniKey -Key $t -Suffix $Script:ConfigIniKeySuffix
    if ($enc -match '_[0-9a-fA-F]+$' -and ($enc -match 'vk[0-9a-fA-F]+' -or $enc -match 'sc[0-9a-fA-F]+')) {
        $h = ConvertFrom-ConfigIniKey -ConfigKey $enc
        if ($h) { return $h }
    }
    return $t
}

function Get-ConfigIniMacroSuffixFromValue {
    param([AllowNull()][string]$Value)
    if ($null -eq $Value -or $Value -eq "") { return "" }
    $i = $Value.IndexOf(';')
    if ($i -lt 0) { return "" }
    return $Value.Substring($i)
}

# Debounce / disk copies are often UTF-8; older flow assumed UTF-16 LE. Wrong encoding breaks [Section] parsing and row load.
function Read-ConfigIniFileText {
    param([Parameter(Mandatory)][string]$Path)
    $bytes = [System.IO.File]::ReadAllBytes($Path)
    if ($bytes.Length -eq 0) { return "" }
    if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) {
        return [System.Text.Encoding]::UTF8.GetString($bytes, 3, $bytes.Length - 3)
    }
    if ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE) {
        return [System.Text.Encoding]::Unicode.GetString($bytes, 2, $bytes.Length - 2)
    }
    if ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFE -and $bytes[1] -eq 0xFF) {
        return [System.Text.Encoding]::BigEndianUnicode.GetString($bytes, 2, $bytes.Length - 2)
    }
    $lim = [Math]::Min($bytes.Length, 8192)
    $nulEven = 0
    for ($i = 1; $i -lt $lim; $i += 2) {
        if ($bytes[$i] -eq 0) { $nulEven++ }
    }
    if ($lim -ge 8 -and $nulEven -gt ($lim / 4)) {
        return [System.Text.Encoding]::Unicode.GetString($bytes)
    }
    return [System.Text.Encoding]::UTF8.GetString($bytes)
}

# Match Read-ConfigIniFileText so re-saved config.ini uses the same encoding ggloader / editors expect.
function Get-ConfigIniFileDiskEncoding {
    param([Parameter(Mandatory)][string]$Path)
    $bytes = [System.IO.File]::ReadAllBytes($Path)
    if ($bytes.Length -eq 0) {
        return [System.Text.UTF8Encoding]::new($false)
    }
    if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) {
        return [System.Text.UTF8Encoding]::new($true)
    }
    if ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE) {
        return [System.Text.Encoding]::Unicode
    }
    if ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFE -and $bytes[1] -eq 0xFF) {
        return [System.Text.Encoding]::BigEndianUnicode
    }
    $lim = [Math]::Min($bytes.Length, 8192)
    $nulEven = 0
    for ($i = 1; $i -lt $lim; $i += 2) {
        if ($bytes[$i] -eq 0) { $nulEven++ }
    }
    if ($lim -ge 8 -and $nulEven -gt ($lim / 4)) {
        return [System.Text.Encoding]::Unicode
    }
    return [System.Text.UTF8Encoding]::new($false)
}

function Repair-ConfigIniTargetMemberEmptyHotkeyToken {
    param([string]$Line)
    if ($Line -notmatch '^(Target Member\d+)=(;\s+.+)$') { return $Line }
    return "$($Matches[1])=$($Script:ConfigIniTargetMemberEmptyKeyPlaceholder)$($Matches[2])"
}

function Get-SyncConfigIniRows {
    param([string]$ConfigPath)
    $out = [System.Collections.ArrayList]::new()
    if (-not (Test-Path -LiteralPath $ConfigPath -PathType Leaf)) { return $out }
    $content = Read-ConfigIniFileText -Path $ConfigPath
    $lines = $content -split "`r?`n"
    $currentSection = $null
    foreach ($line in $lines) {
        if ($null -eq $line) { continue }
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
        if ($Script:ConfigIniExcludeActionNames.ContainsKey($actionName)) { continue }
        if ($actionName -match '^Universal\d+ Unit\d+$') { continue }
        $keyPart = $value
        $afterSemi = ""
        if ($value -match '^([^;]*);(.*)$') {
            $keyPart = $Matches[1].Trim()
            $afterSemi = $Matches[2].Trim()
        }
        if ($afterSemi -and $afterSemi -match $Script:ConfigIniDescriptionMacroPattern) { continue }
        # Include Debounce-style compact keys (^ ! + prefixes) and obvious numpad token names, not only vk/sc_* lines.
        # Otherwise rows never load, Clear-KeyIfNumpadRelated never runs, and config.ini keeps !NUMPADUP forever.
        $compactOrNumpadToken = ($keyPart -match '^[\^!+]') -or ($keyPart -match '(?i)NUMPAD')
        if ($keyPart -and $keyPart -notmatch $Script:ConfigIniKeybindValuePattern -and -not $afterSemi -and -not $compactOrNumpadToken) { continue }
        [void]$out.Add([PSCustomObject]@{
            Section    = $currentSection
            ActionName = $actionName
            Key        = $keyPart
            MacroText  = (Normalize-MacroTextForCrimsonBindLua -Raw $afterSemi)
            TextureID  = "132089"
        })
    }
    return $out
}

# All [Section] headers in config.ini that are bind sections (excludes Main, SourceSettings, etc.).
# The row parser skips lines that do not look like hotkeys, so SyncRows alone can omit entire sections from the dropdown.
function Get-SyncConfigIniBindSectionNames {
    param([string]$ConfigPath)
    $set = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    if (-not $ConfigPath -or -not (Test-Path -LiteralPath $ConfigPath -PathType Leaf)) { return @() }
    $content = Read-ConfigIniFileText -Path $ConfigPath
    $lines = $content -split "`r?`n"
    foreach ($line in $lines) {
        if ($null -eq $line) { continue }
        $strip = $line.Trim()
        if ($strip -notmatch '^\[(.+)\]$') { continue }
        $raw = $Matches[1].Trim()
        $sec = if ($raw.Length -gt 0 -and [int][char]$raw[0] -eq 0xFEFF) { $raw.Substring(1).Trim() } else { $raw }
        if (-not $sec) { continue }
        if ($Script:ConfigIniNonBindSections.ContainsKey($sec)) { continue }
        [void]$set.Add($sec)
    }
    return @($set | Sort-Object)
}

function Update-SyncRowsDecodedKeys {
    param([System.Collections.ArrayList]$Rows)
    foreach ($r in @($Rows)) {
        if ($null -eq $r) { continue }
        $kp = if ($null -eq $r.Key) { "" } else { $r.Key.ToString().Trim() }
        $r.Key = Clear-KeyIfNumpadRelated -Key $kp
    }
}

# Matches Set-ConfigIniFromRows: any non-General row whose ActionName exists under [General] uses General's Key (WoW LUA and profile stay consistent).
function Sync-SpecRowKeysToMatchGeneral {
    param([System.Collections.ArrayList]$Rows)
    if (-not $Rows) { return }
    $gen = @{}
    foreach ($r in @($Rows)) {
        if ($null -eq $r) { continue }
        if ($r.Section -ne "General") { continue }
        $a = if ($null -eq $r.ActionName) { "" } else { $r.ActionName.ToString().Trim() }
        if (-not $a) { continue }
        $gen[$a] = $r.Key
    }
    if ($gen.Count -eq 0) { return }
    foreach ($r in @($Rows)) {
        if ($null -eq $r) { continue }
        $sn = if ($null -eq $r.Section) { "" } else { $r.Section.ToString().Trim() }
        # Skip General (already the source), empty sections, and CUSTOM (never mirrors General keys)
        if ($sn -eq "General" -or -not $sn -or ($sn -ieq 'CUSTOM')) { continue }
        $a = if ($null -eq $r.ActionName) { "" } else { $r.ActionName.ToString().Trim() }
        if (-not $a) { continue }
        if (-not $gen.ContainsKey($a)) { continue }
        $r.Key = $gen[$a]
    }
}

# Returns $true when a row Section is the special CUSTOM user-bind section (case-insensitive).
function Test-IsCrimsonBindCustomSection {
    param([string]$Section)
    return ($Section -ieq 'CUSTOM')
}

# Appends CSV rows whose Section is CUSTOM to SyncRows if not already present (by Section|ActionName).
# Call this after Get-SyncConfigIniRows so user-created custom binds ride along the normal pipeline
# without needing to exist in config.ini. Returns the count of rows appended.
function Merge-CrimsonBindCsvCustomRowsIntoSyncRows {
    param(
        [System.Collections.ArrayList]$SyncRows,
        [string]$CsvPath
    )
    if (-not $SyncRows) { return 0 }
    if (-not $CsvPath -or -not (Test-Path -LiteralPath $CsvPath -PathType Leaf)) { return 0 }
    try {
        $csvRows = @(Import-Csv -LiteralPath $CsvPath -Encoding UTF8 -ErrorAction Stop)
    } catch { return 0 }
    $existingKeys = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($r in $SyncRows) {
        if ($null -eq $r) { continue }
        $s = if ($null -eq $r.Section) { "" } else { $r.Section.ToString().Trim() }
        $a = if ($null -eq $r.ActionName) { "" } else { $r.ActionName.ToString().Trim() }
        [void]$existingKeys.Add("$s|$a")
    }
    $added = 0
    foreach ($row in $csvRows) {
        if ($null -eq $row) { continue }
        $sec = (Import-CsvRowProperty -Row $row -Names @('Section')).Trim()
        if (-not ($sec -ieq 'CUSTOM')) { continue }
        $an = (Import-CsvRowProperty -Row $row -Names @('ActionName')).Trim()
        if (-not $an) { continue }
        $lk = "$sec|$an"
        if ($existingKeys.Contains($lk)) { continue }
        $mt = Normalize-MacroTextForCrimsonBindLua -Raw (Import-CsvRowProperty -Row $row -Names @('MacroText'))
        $key = Clear-KeyIfNumpadRelated -Key ((Import-CsvRowProperty -Row $row -Names @('Key')).Trim())
        [void]$SyncRows.Add([PSCustomObject]@{
            Section = $sec; ActionName = $an; Key = $key; MacroText = $mt; TextureID = "132089"
        })
        [void]$existingKeys.Add($lk)
        $added++
    }
    return $added
}

# Like Get-GeneralReservedKeyCanonSetFromCsv but collects only CUSTOM-section keys.
function Get-CustomReservedKeyCanonSetFromCsv {
    param([string]$CsvPath)
    $set = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    if (-not $CsvPath -or -not (Test-Path -LiteralPath $CsvPath -PathType Leaf)) { return $set }
    try {
        $tbl = @(Import-Csv -LiteralPath $CsvPath -Encoding UTF8 -ErrorAction Stop)
    } catch { return $set }
    foreach ($row in $tbl) {
        if ($null -eq $row) { continue }
        $sec = (Import-CsvRowProperty -Row $row -Names @('Section')).Trim()
        if (-not ($sec -ieq 'CUSTOM')) { continue }
        $kx = Clear-KeyIfNumpadRelated -Key ((Import-CsvRowProperty -Row $row -Names @('Key')).Trim())
        if (-not $kx) { continue }
        $cc = Get-BindKeyCanonicalConfigForm -Key $kx
        if ($cc) { [void]$set.Add($cc) }
    }
    return $set
}

# WoW / profile loaders / editors often keep config.ini open; in-place WriteAllText fails with "locked a portion of the file".
# Write a temp file then Replace (atomic on NTFS) with retries on IOException.
function Write-ConfigIniFileRobust {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Content,
        [Parameter(Mandatory)][System.Text.Encoding]$Encoding
    )
    $enc = $Encoding
    $dir = [System.IO.Path]::GetDirectoryName($Path)
    if (-not $dir) { $dir = [System.IO.Path]::GetFullPath('.') }
    $name = [System.IO.Path]::GetFileName($Path)
    $tmp = Join-Path $dir ($name + '.crimsonbind-' + $PID + '-' + [Guid]::NewGuid().ToString('N').Substring(0, 8) + '.tmp')
    $maxAttempts = 15
    for ($attempt = 1; $attempt -le $maxAttempts; $attempt++) {
        if (Test-Path -LiteralPath $tmp) { Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue }
        try {
            [System.IO.File]::WriteAllText($tmp, $Content, $enc)
            if (Test-Path -LiteralPath $Path) {
                $bak = $Path + '.bak'
                [System.IO.File]::Replace($tmp, $Path, $bak, $true)
            } else {
                [System.IO.File]::Move($tmp, $Path)
            }
            return
        } catch {
            if (Test-Path -LiteralPath $tmp) { Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue }
            $ex = $_.Exception
            $retry = ($ex -is [System.IO.IOException]) -or ($ex -is [UnauthorizedAccessException])
            if (-not $retry -and $ex.InnerException) {
                $retry = $ex.InnerException -is [System.IO.IOException]
            }
            if (-not $retry) { throw }
            if ($attempt -eq $maxAttempts) {
                throw [System.IO.IOException]::new(
                    "Could not save config.ini after $maxAttempts attempts. Close WoW, Debounce, or any editor with this file open, then retry.",
                    $ex
                )
            }
            Start-Sleep -Milliseconds ([Math]::Min(100 * $attempt, 2000))
        }
    }
}

# Writes encoded key tokens only; preserves everything after the first ';' on each existing bind line (profile macro bodies). Macro text for CrimsonBind comes from CSV → CrimsonBind.lua, not from pushing CSV macros into config.ini.
function Set-ConfigIniFromRows {
    param(
        [string]$ConfigPath,
        [System.Collections.ArrayList]$Rows,
        [string]$OnlySection = $null,
        [switch]$ForceGeneralIniKeys
    )
    if (-not (Test-Path -LiteralPath $ConfigPath -PathType Leaf)) { return -1 }
    $content = Read-ConfigIniFileText -Path $ConfigPath
    $generalByActionName = @{}
    foreach ($r in @($Rows)) {
        if ($null -eq $r) { continue }
        if ($r.Section -eq "General" -and $null -ne $r.ActionName) {
            $generalByActionName[$r.ActionName.ToString().Trim()] = [PSCustomObject]@{ Key = $r.Key; MacroText = $r.MacroText }
        }
    }
    $keyMap = @{}
    foreach ($r in @($Rows)) {
        if ($null -eq $r) { continue }
        $secN = if ($null -eq $r.Section) { "" } else { $r.Section.ToString().Trim() }
        # CUSTOM rows never touch config.ini — they live only in CSV/SavedVariables
        if ($secN -ieq 'CUSTOM') { continue }
        $actN = if ($null -eq $r.ActionName) { "" } else { $r.ActionName.ToString().Trim() }
        $k = "$secN|$actN"
        if ($r.Section -eq "General") {
            $keyMap[$k] = [PSCustomObject]@{ Key = $r.Key; MacroText = $r.MacroText }
        } elseif ($actN -and $generalByActionName.ContainsKey($actN)) {
            $keyMap[$k] = $generalByActionName[$actN]
        } else {
            $keyMap[$k] = [PSCustomObject]@{ Key = $r.Key; MacroText = $r.MacroText }
        }
    }
    if ($OnlySection -and $OnlySection -ne "General") {
        foreach ($actionName in $generalByActionName.Keys) {
            $keyMap["$OnlySection|$actionName"] = $generalByActionName[$actionName]
        }
    }
    $lines = $content -split "`r?`n"
    $currentSection = $null
    $newLines = [System.Collections.ArrayList]::new()
    $replaced = 0
    $added = 0
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
            $rawK = if ($null -eq $info.Key) { "" } else { $info.Key.ToString().Trim() }
            $configKey = ConvertTo-ConfigIniKey -Key $rawK -Suffix $Script:ConfigIniKeySuffix
            if ([string]::IsNullOrWhiteSpace($rawK) -or [string]::IsNullOrWhiteSpace($configKey)) { continue }
            [void]$targetNewLines.Add("$act=$configKey")
            $count++
        }
        return $count
    }

    foreach ($line in $lines) {
        if ($null -eq $line) { continue }
        $strip = $line.TrimEnd("`r", "`n")
        if ($strip -match '^\[(.+)\]$') {
            $nextSection = $Matches[1].Trim()
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
            $lineNamePart = if ($null -ne $Matches[1]) { $Matches[1] } else { "" }
            $lineValPart  = if ($null -ne $Matches[2]) { $Matches[2] } else { "" }
            $actionName = $lineNamePart.Trim()
            if ($actionName) { $writtenInSection[$currentSection][$actionName] = $true }
            $key = "$currentSection|$actionName"
            $inTargetSection = -not $OnlySection -or ($currentSection -eq $OnlySection)
            $skipTm = $false
            if ($actionName -match '^Target Member\d+$' -and $currentSection -ne 'General' -and -not (Test-SectionStringIsHealerHealingSpec -Section $currentSection)) {
                $rawSk = ""
                if ($keyMap.ContainsKey($key)) {
                    $infSk = $keyMap[$key]
                    $rawSk = if ($null -eq $infSk.Key) { "" } else { $infSk.Key.ToString().Trim() }
                }
                if ([string]::IsNullOrWhiteSpace($rawSk)) { $skipTm = $true }
            }
            if ($inTargetSection -and $actionName -and -not $skipTm -and $keyMap.ContainsKey($key)) {
                $info = $keyMap[$key]
                $rawK = if ($null -eq $info.Key) { "" } else { $info.Key.ToString().Trim() }
                $configKey = ConvertTo-ConfigIniKey -Key $rawK -Suffix $Script:ConfigIniKeySuffix
                $oldVal = $lineValPart
                if (-not $ForceGeneralIniKeys -and $currentSection -eq 'General') {
                    $semiIdx = $oldVal.IndexOf([char]';')
                    $diskKeyPart = if ($semiIdx -lt 0) { $oldVal.Trim() } else { $oldVal.Substring(0, $semiIdx).Trim() }
                    if ([string]::IsNullOrWhiteSpace($diskKeyPart)) {
                        [void]$newLines.Add($line)
                        continue
                    }
                }
                if ([string]::IsNullOrWhiteSpace($rawK) -or [string]::IsNullOrWhiteSpace($configKey)) {
                    [void]$newLines.Add($line)
                } else {
                    $oldValTrim = $oldVal.Trim()
                    $suffix = Get-ConfigIniMacroSuffixFromValue -Value $oldValTrim
                    $newVal = $configKey + $suffix
                    [void]$newLines.Add($lineNamePart + "=" + $newVal)
                    $replaced++
                }
            } else { [void]$newLines.Add($line) }
        } else { [void]$newLines.Add($line) }
    }
    if ($null -ne $currentSection) {
        $n = Add-MissingKeysForSection -section $currentSection -keyMap $keyMap -targetNewLines $newLines
        if ($n -gt 0) { $added += $n }
    }
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
    $repairedLines = [System.Collections.ArrayList]::new()
    foreach ($ln in $newLines) {
        [void]$repairedLines.Add((Repair-ConfigIniTargetMemberEmptyHotkeyToken -Line $ln))
    }
    $outText = ($repairedLines -join "`r`n") + "`r`n"
    $writeEnc = Get-ConfigIniFileDiskEncoding -Path $ConfigPath
    Write-ConfigIniFileRobust -Path $ConfigPath -Content $outText -Encoding $writeEnc
    return ($replaced + $added)
}

# --- Key pool from bindable_keys.ini ---

function Read-BindableKeysIniSection {
    param([string]$Path, [string]$SectionName)
    $list = [System.Collections.ArrayList]::new()
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { return $list }
    $inSection = $false
    foreach ($line in (Get-Content -LiteralPath $Path -Encoding UTF8 -ErrorAction SilentlyContinue)) {
        if ($null -eq $line) { continue }
        $t = $line.Trim()
        if (-not $t -or $t.StartsWith('#')) { continue }
        if ($t -match '^\[(.+)\]$') {
            $inSection = ($Matches[1].Trim() -eq $SectionName)
            continue
        }
        if ($inSection) { [void]$list.Add($t) }
    }
    return $list
}

function Test-BindKeyBaseInExcludedSet {
    param([string]$Key, [System.Collections.Generic.HashSet[string]]$ExcludedBases)
    if (-not $Key -or -not $ExcludedBases -or $ExcludedBases.Count -eq 0) { return $false }
    $pb = Split-BindKeyIntoModsAndBase -Key $Key
    $b = if ($pb.Base) { $pb.Base.Trim().ToUpperInvariant() } else { "" }
    return $b -and $ExcludedBases.Contains($b)
}

function Test-BindKeyIsShiftWithNumpadDigit {
    param([string]$Key)
    if (-not $Key) { return $false }
    $pb = Split-BindKeyIntoModsAndBase -Key $Key
    $hasShift = $false
    foreach ($x in $pb.Mods) {
        if ($null -eq $x) { continue }
        if ($x.Trim().ToUpperInvariant() -eq 'SHIFT') { $hasShift = $true; break }
    }
    if (-not $hasShift) { return $false }
    $b = if ($pb.Base) { $pb.Base.Trim().ToUpperInvariant() } else { "" }
    return ($b -match '^NUMPAD[0-9]$')
}

function Test-BindKeyBaseIsNumpadRelated {
    param([string]$Key)
    if (-not $Key) { return $false }
    $pb = Split-BindKeyIntoModsAndBase -Key $Key
    $b = if ($pb.Base) { $pb.Base.Trim().ToUpperInvariant() } else { "" }
    if (-not $b) { return $false }
    if ($b -match '^NUMPAD\d$') { return $true }
    if ($Script:NumpadNavSynonymToNumpad.ContainsKey($b)) { return $true }
    if ($b -like 'NUMPAD*') { return $true }
    return $false
}

function Clear-KeyIfNumpadRelated {
    param([string]$Key)
    if ($null -eq $Key) { return "" }
    $kp = ($Key -replace "^\s+|\s+$", "")
    if (-not $kp) { return "" }
    $norm = Normalize-BindKeyPipeline -Key $kp
    if (-not $norm) { return "" }
    if (Test-BindKeyBaseIsNumpadRelated -Key $norm) { return "" }
    if ($norm -match 'NUMPAD') { return "" }
    return $norm
}

# bindable_keys.ini lines like "CTRL-ALT-" must become CTRL-ALT-PAGEUP, not CTRL-ALTPAGEUP (which breaks Split-BindKeyIntoModsAndBase).
function Join-BindKeyModifierPrefixToBase {
    param([string]$ModPrefix, [string]$BaseName)
    $m = ($ModPrefix -replace '^\s+|\s+$', "")
    $b = ($BaseName -replace '^\s+|\s+$', "")
    if (-not $m) { return $b }
    if (-not $b) { return $m }
    if ($m.EndsWith('-')) { return ($m.TrimEnd('-').Trim() + '-' + $b) }
    return ($m + '-' + $b)
}

function Build-KeyPoolFromIni {
    param([string]$Path)
    $mods = Read-BindableKeysIniSection -Path $Path -SectionName "Modifiers"
    $base = Read-BindableKeysIniSection -Path $Path -SectionName "BaseKeys"
    $numpad = Read-BindableKeysIniSection -Path $Path -SectionName "NumpadKeys"
    $explicit = Read-BindableKeysIniSection -Path $Path -SectionName "ExplicitKeys"
    $excluded = Read-BindableKeysIniSection -Path $Path -SectionName "ExcludedKeys"
    $exBaseLines = Read-BindableKeysIniSection -Path $Path -SectionName "ExcludedBaseKeys"
    $excludedBases = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($line in $exBaseLines) {
        if ($null -eq $line) { continue }
        $t = $line.Trim()
        if ($t) { [void]$excludedBases.Add($t.ToUpperInvariant()) }
    }
    $pool = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    # No unmodified BaseKeys or NumpadKeys — only MOD+key combos (and [ExplicitKeys]). Avoids bare R, E, 7, NUMPAD8, etc.
    foreach ($m in $mods) {
        if (-not $m -or -not $m.Trim()) { continue }
        foreach ($bk in ($base + $numpad)) {
            if (-not $bk -or -not $bk.Trim()) { continue }
            $bkU = $bk.Trim().ToUpperInvariant()
            if ($excludedBases.Contains($bkU)) { continue }
            if ($Script:RandomPoolExcludeBasesNumpadVkAlias.ContainsKey($bkU)) { continue }
            $combo = Join-BindKeyModifierPrefixToBase -ModPrefix $m -BaseName $bk
            if (Test-BindKeyIsShiftWithNumpadDigit -Key $combo) { continue }
            if (Test-BindKeyBaseIsNumpadRelated -Key $combo) { continue }
            if ($combo -ne "CTRL-C" -and $combo -ne "CTRL-V") { [void]$pool.Add($combo) }
        }
    }
    foreach ($k in $explicit) {
        if (-not $k) { continue }
        if (Test-BindKeyIsShiftWithNumpadDigit -Key $k) { continue }
        if (Test-BindKeyBaseIsNumpadRelated -Key $k) { continue }
        if (Test-BindKeyBaseInExcludedSet -Key $k -ExcludedBases $excludedBases) { continue }
        $eb = Split-BindKeyIntoModsAndBase -Key $k
        $ebu = if ($eb.Base) { $eb.Base.Trim().ToUpperInvariant() } else { "" }
        if ($ebu -and $Script:RandomPoolExcludeBasesNumpadVkAlias.ContainsKey($ebu)) { continue }
        [void]$pool.Add($k)
    }
    $exclSet = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($e in $excluded) {
        $c = Get-BindKeyCanonicalConfigForm -Key $e
        if ($c) { [void]$exclSet.Add($c) }
    }
    $out = [System.Collections.ArrayList]::new()
    foreach ($k in $pool) {
        if (Test-BindKeyIsShiftWithNumpadDigit -Key $k) { continue }
        if (Test-BindKeyBaseIsNumpadRelated -Key $k) { continue }
        if (Test-BindKeyUsesAltShift -Key $k) { continue }
        if (Test-BindKeyBaseInExcludedSet -Key $k -ExcludedBases $excludedBases) { continue }
        $cc = Get-BindKeyCanonicalConfigForm -Key $k
        if ($exclSet.Contains($cc)) { continue }
        [void]$out.Add($k)
    }
    $out.Sort()
    return $out
}

function Write-BindableKeysIniPreset {
    param([ValidateSet("Full", "RestoRightHand", "Minimal")][string]$Preset, [string]$Path)
    $sb = [System.Text.StringBuilder]::new()
    [void]$sb.AppendLine("# Auto-generated preset: $Preset")
    [void]$sb.AppendLine()
    [void]$sb.AppendLine("[Modifiers]")
    if ($Preset -eq "Minimal") {
        foreach ($x in @("CTRL-", "ALT-", "SHIFT-")) { [void]$sb.AppendLine($x) }
    } elseif ($Preset -eq "RestoRightHand") {
        foreach ($x in $Script:ModsResto) { [void]$sb.AppendLine($x) }
    } else {
        foreach ($x in $Script:ModsFull) { [void]$sb.AppendLine($x) }
    }
    [void]$sb.AppendLine()
    [void]$sb.AppendLine("[BaseKeys]")
    $keys = if ($Preset -eq "RestoRightHand") { $Script:KeyPoolRestoRightHand } else { $Script:KeyPoolFull }
    foreach ($k in $keys) {
        if ($k -like "NUMPAD*") { continue }
        [void]$sb.AppendLine($k)
    }
    [void]$sb.AppendLine()
    [void]$sb.AppendLine("[NumpadKeys]")
    foreach ($k in $Script:KeyPoolNumpadNumbers) { [void]$sb.AppendLine($k) }
    [void]$sb.AppendLine()
    [void]$sb.AppendLine("[ExplicitKeys]")
    [void]$sb.AppendLine()
    [void]$sb.AppendLine("[ExcludedKeys]")
    [void]$sb.AppendLine("CTRL-C")
    [void]$sb.AppendLine("CTRL-V")
    [void]$sb.AppendLine()
    [void]$sb.AppendLine("[ExcludedBaseKeys]")
    foreach ($xb in @("Q", "W", "E", "A", "S", "D", "R", "F", "1", "2", "3", "4", "5")) {
        [void]$sb.AppendLine($xb)
    }
    [System.IO.File]::WriteAllText($Path, $sb.ToString(), [System.Text.UTF8Encoding]::new($true))
}

function Get-GeneralReservedKeyCanonSetFromCsv {
    param([string]$CsvPath)
    $set = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    if (-not $CsvPath -or -not (Test-Path -LiteralPath $CsvPath -PathType Leaf)) { return $set }
    try {
        $tbl = @(Import-Csv -LiteralPath $CsvPath -Encoding UTF8 -ErrorAction Stop)
    } catch {
        return $set
    }
    foreach ($row in $tbl) {
        if ($null -eq $row) { continue }
        $sec = (Import-CsvRowProperty -Row $row -Names @('Section')).Trim()
        # Reserve both General keys and CUSTOM keys so neither pool can collide with user-defined binds
        if ($sec -ine 'General' -and $sec -ine 'CUSTOM') { continue }
        $kx = Clear-KeyIfNumpadRelated -Key ((Import-CsvRowProperty -Row $row -Names @('Key')).Trim())
        if (-not $kx) { continue }
        $cc = Get-BindKeyCanonicalConfigForm -Key $kx
        if ($cc) { [void]$set.Add($cc) }
    }
    return $set
}

function Get-SyncSpecPoolExcludingGeneral {
    param(
        [System.Collections.ArrayList]$FullPool,
        [System.Collections.ArrayList]$Rows,
        [string]$CrimsonBindCsvPath = $null
    )
    if (-not $FullPool) { return [System.Collections.ArrayList]::new() }
    $canonSet = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($r in @($Rows)) {
        $sn = if ($r.Section) { $r.Section.ToString().Trim() } else { "" }
        # Exclude both General keys (already reserved) and CUSTOM keys (user-pinned, never randomized)
        if ($sn -ine "General" -and $sn -ine "CUSTOM") { continue }
        $kx = if ($r.Key) { $r.Key.ToString().Trim() } else { "" }
        if (-not $kx) { continue }
        [void]$canonSet.Add((Get-BindKeyCanonicalConfigForm -Key $kx))
    }
    if ($CrimsonBindCsvPath) {
        foreach ($c in (Get-GeneralReservedKeyCanonSetFromCsv -CsvPath $CrimsonBindCsvPath)) {
            [void]$canonSet.Add($c)
        }
    }
    $out = [System.Collections.ArrayList]::new()
    foreach ($k in $FullPool) {
        $gc = Get-BindKeyCanonicalConfigForm -Key $k
        if ($canonSet.Contains($gc)) { continue }
        [void]$out.Add($k)
    }
    return $out
}

function Invoke-RandomizeAssignUniqueKeysToRows {
    param(
        [System.Collections.ArrayList]$TargetRows,
        [System.Collections.ArrayList]$RowsAll,
        [string]$SectionName,
        [System.Collections.ArrayList]$ShuffledPool,
        [hashtable]$GeneralActionNames,
        [System.Random]$Rng
    )
    if (-not $TargetRows -or $TargetRows.Count -eq 0) { return }
    if (-not $Rng) { $Rng = [System.Random]::new() }
    $targetIds = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    foreach ($tr in $TargetRows) {
        if ($null -eq $tr) { continue }
        $ts = if ($null -eq $tr.Section) { "" } else { $tr.Section.ToString().Trim() }
        $ta = if ($null -eq $tr.ActionName) { "" } else { $tr.ActionName.ToString().Trim() }
        [void]$targetIds.Add("$ts|$ta")
    }
    $usedCanon = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($r in @($RowsAll)) {
        if ($null -eq $r) { continue }
        if ($r.Section -ne $SectionName) { continue }
        $actName = if ($null -eq $r.ActionName) { "" } else { $r.ActionName.ToString().Trim() }
        if ($SectionName -ne "General" -and $GeneralActionNames -and $actName -and $GeneralActionNames.ContainsKey($actName)) { continue }
        $id = "$(if ($null -eq $r.Section) { '' } else { $r.Section.ToString().Trim() })|$actName"
        if ($targetIds.Contains($id)) { continue }
        $kx = if ($null -eq $r.Key) { "" } else { $r.Key.ToString().Trim() }
        if (-not $kx) { continue }
        [void]$usedCanon.Add((Get-BindKeyCanonicalConfigForm -Key $kx))
    }
    $tmFirst = [System.Collections.ArrayList]::new()
    $rest = [System.Collections.ArrayList]::new()
    foreach ($x in ($TargetRows | Sort-Object { $Rng.Next() })) {
        if ($null -eq $x) { continue }
        if (Test-ActionNameIsTargetMember1Through10 -ActionName $x.ActionName) { [void]$tmFirst.Add($x) } else { [void]$rest.Add($x) }
    }
    $order = [System.Collections.ArrayList]::new()
    foreach ($x in $tmFirst) { [void]$order.Add($x) }
    foreach ($x in $rest) { [void]$order.Add($x) }
    $tmExtraShuffled = [System.Collections.ArrayList]::new()
    $needTmExtra = $false
    foreach ($tr in $TargetRows) {
        if ($null -eq $tr) { continue }
        $tn = if ($null -eq $tr.ActionName) { "" } else { $tr.ActionName.ToString() }
        if (Test-ActionNameIsTargetMember1Through10 -ActionName $tn) { $needTmExtra = $true; break }
    }
    if ($needTmExtra) {
        $extras = [System.Collections.ArrayList]::new()
        foreach ($fk in @("F1", "F2", "F3", "F4", "F5", "F6", "F7", "F8", "F9", "F10", "F11", "F12")) {
            foreach ($cand in @($fk, "SHIFT-$fk")) {
                if (Test-BindKeyUsesCtrlOrAlt -Key $cand) { continue }
                [void]$extras.Add($cand)
            }
        }
        foreach ($x in ($extras | Sort-Object { $Rng.Next() })) { [void]$tmExtraShuffled.Add($x) }
    }
    foreach ($r in $order) {
        if ($null -eq $r) { continue }
        $rn = if ($null -eq $r.ActionName) { "" } else { $r.ActionName.ToString() }
        $restrictTm = Test-ActionNameIsTargetMember1Through10 -ActionName $rn
        $candidates = if ($restrictTm) { [object[]](@($ShuffledPool) + @($tmExtraShuffled)) } else { [object[]]$ShuffledPool }
        foreach ($cand in $candidates) {
            if ($null -eq $cand) { continue }
            if (Test-BindKeyUsesAltShift -Key $cand) { continue }
            if (Test-BindKeyBaseIsNumpadRelated -Key $cand) { continue }
            $pbx = Split-BindKeyIntoModsAndBase -Key $cand
            $bxu = if ($pbx.Base) { $pbx.Base.Trim().ToUpperInvariant() } else { "" }
            if ($bxu -and $Script:RandomPoolExcludeBasesNumpadVkAlias.ContainsKey($bxu)) { continue }
            if ($restrictTm -and (Test-BindKeyUsesCtrlOrAlt -Key $cand)) { continue }
            $cc = Get-BindKeyCanonicalConfigForm -Key $cand
            if (-not $cc) { continue }
            if ($usedCanon.Contains($cc)) { continue }
            $r.Key = Normalize-BindKeyDisplayFromTokens -Key $cand
            [void]$usedCanon.Add($cc)
            break
        }
    }
}

function Invoke-SyncRandomizeKeys {
    param(
        [System.Collections.ArrayList]$Rows,
        [string[]]$OnlySections = $null,
        [string]$KeyPoolPath,
        [string]$CrimsonBindCsvPath = $null
    )
    if (-not $Rows) { throw "No bind rows loaded (load config.ini first)." }
    $fullPool = Build-KeyPoolFromIni -Path $KeyPoolPath
    if ($fullPool.Count -eq 0) {
        throw "Key pool is empty. Check bindable_keys.ini at: $KeyPoolPath"
    }
    $rng = [System.Random]::new()
    # Pool for spec sections: excludes both General-reserved keys and CUSTOM-pinned keys
    $poolWithoutGeneral = Get-SyncSpecPoolExcludingGeneral -FullPool $fullPool -Rows $Rows -CrimsonBindCsvPath $CrimsonBindCsvPath
    # Pool for General section: excludes only CUSTOM-pinned keys (General keys belong to General)
    $customCanonSet = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($r in @($Rows)) {
        if ($null -eq $r) { continue }
        $sn = if ($null -eq $r.Section) { "" } else { $r.Section.ToString().Trim() }
        if (-not ($sn -ieq 'CUSTOM')) { continue }
        $kx = if ($null -eq $r.Key) { "" } else { $r.Key.ToString().Trim() }
        if ($kx) { [void]$customCanonSet.Add((Get-BindKeyCanonicalConfigForm -Key $kx)) }
    }
    if ($CrimsonBindCsvPath) {
        foreach ($c in (Get-CustomReservedKeyCanonSetFromCsv -CsvPath $CrimsonBindCsvPath)) { [void]$customCanonSet.Add($c) }
    }
    $poolForGeneral = [System.Collections.ArrayList]::new()
    foreach ($k in $fullPool) {
        if ($customCanonSet.Count -gt 0 -and $customCanonSet.Contains((Get-BindKeyCanonicalConfigForm -Key $k))) { continue }
        [void]$poolForGeneral.Add($k)
    }
    if ($OnlySections -and $OnlySections.Count -gt 0) {
        $generalActionNames = @{}
        foreach ($r in @($Rows)) {
            if ($null -eq $r) { continue }
            if ($r.Section -eq "General" -and $null -ne $r.ActionName) { $generalActionNames[$r.ActionName.ToString().Trim()] = $true }
        }
        # Randomize General first (if selected) so spec pool reflects new General keys
        $generalSelected = $false
        foreach ($secName in $OnlySections) {
            if ($secName -eq "General") { $generalSelected = $true; break }
        }
        if ($generalSelected) {
            $targetRows = [System.Collections.ArrayList]::new()
            foreach ($r in @($Rows)) {
                if ($null -eq $r) { continue }
                if ($r.Section -ne "General") { continue }
                $actN = if ($null -eq $r.ActionName) { "" } else { $r.ActionName.ToString().Trim() }
                if ($actN -match '^Target Member(\d+)$') {
                    $tmi = 0
                    if (-not [int]::TryParse($Matches[1], [ref]$tmi) -or $tmi -gt 10) { continue }
                }
                [void]$targetRows.Add($r)
            }
            $shuffled = [System.Collections.ArrayList]::new()
            foreach ($k in ($poolForGeneral | Sort-Object { $rng.Next() })) { [void]$shuffled.Add($k) }
            Invoke-RandomizeAssignUniqueKeysToRows -TargetRows $targetRows -RowsAll $Rows -SectionName "General" -ShuffledPool $shuffled -GeneralActionNames $generalActionNames -Rng $rng
            # Recompute spec pool with General's newly assigned keys
            $poolWithoutGeneral = Get-SyncSpecPoolExcludingGeneral -FullPool $fullPool -Rows $Rows -CrimsonBindCsvPath $CrimsonBindCsvPath
        }
        foreach ($secName in $OnlySections) {
            if ($secName -ieq 'CUSTOM') { continue }
            if ($secName -eq "General") { continue }  # Already handled above
            $targetRows = [System.Collections.ArrayList]::new()
            foreach ($r in @($Rows)) {
                if ($null -eq $r) { continue }
                if ($r.Section -ne $secName) { continue }
                $actN = if ($null -eq $r.ActionName) { "" } else { $r.ActionName.ToString().Trim() }
                if ($actN -and $generalActionNames.ContainsKey($actN)) { continue }
                if ($actN -match '^Target Member(\d+)$' -and -not (Test-SectionStringIsHealerHealingSpec -Section $secName)) {
                    $tmi = 0
                    if (-not [int]::TryParse($Matches[1], [ref]$tmi) -or $tmi -gt 10) { continue }
                }
                [void]$targetRows.Add($r)
            }
            $shuffled = [System.Collections.ArrayList]::new()
            foreach ($k in ($poolWithoutGeneral | Sort-Object { $rng.Next() })) { [void]$shuffled.Add($k) }
            Invoke-RandomizeAssignUniqueKeysToRows -TargetRows $targetRows -RowsAll $Rows -SectionName $secName -ShuffledPool $shuffled -GeneralActionNames $generalActionNames -Rng $rng
        }
        Sync-SpecRowKeysToMatchGeneral -Rows $Rows
        return
    }
    $generalList = [System.Collections.ArrayList]::new()
    $bySection = @{}
    foreach ($r in @($Rows)) {
        if ($null -eq $r) { continue }
        # CUSTOM rows are never randomized — they keep their user-pinned keys
        if (Test-IsCrimsonBindCustomSection -Section $r.Section) { continue }
        if ($r.Section -eq "General") { [void]$generalList.Add($r) }
        else {
            $secN = if ($null -eq $r.Section) { "" } else { $r.Section.ToString().Trim() }
            $actN2 = if ($null -eq $r.ActionName) { "" } else { $r.ActionName.ToString().Trim() }
            if ($actN2 -match '^Target Member(\d+)$' -and -not (Test-SectionStringIsHealerHealingSpec -Section $secN)) {
                $tmi2 = 0
                if (-not [int]::TryParse($Matches[1], [ref]$tmi2) -or $tmi2 -gt 10) { continue }
            }
            if (-not $secN) { continue }
            if (-not $bySection[$secN]) { $bySection[$secN] = [System.Collections.ArrayList]::new() }
            [void]$bySection[$secN].Add($r)
        }
    }
    $generalActionNames = @{}
    foreach ($r in @($Rows)) {
        if ($null -eq $r) { continue }
        if ($r.Section -eq "General" -and $null -ne $r.ActionName) { $generalActionNames[$r.ActionName.ToString().Trim()] = $true }
    }
    $rowsRef = $Rows
    $assignSection = {
        param([string]$secName, [System.Collections.ArrayList]$sectionRows, [System.Collections.ArrayList]$pool, [System.Random]$RngLocal)
        if (-not $RngLocal) { $RngLocal = [System.Random]::new() }
        if (-not $pool) { return }
        if (-not $sectionRows -or $sectionRows.Count -eq 0) { return }
        $shuffled = [System.Collections.ArrayList]::new()
        foreach ($k in ($pool | Sort-Object { $RngLocal.Next() })) { [void]$shuffled.Add($k) }
        Invoke-RandomizeAssignUniqueKeysToRows -TargetRows $sectionRows -RowsAll $rowsRef -SectionName $secName -ShuffledPool $shuffled -GeneralActionNames $generalActionNames -Rng $RngLocal
    }
    & $assignSection "General" $generalList $poolForGeneral $rng
    $poolWithoutGeneral2 = Get-SyncSpecPoolExcludingGeneral -FullPool $fullPool -Rows $Rows -CrimsonBindCsvPath $CrimsonBindCsvPath
    foreach ($sec in @($bySection.Keys)) {
        $secRows = $bySection[$sec]
        if ($secRows) {
            & $assignSection $sec $secRows $poolWithoutGeneral2 $rng
        }
    }
    Sync-SpecRowKeysToMatchGeneral -Rows $Rows
}

function Escape-LuaString {
    param([string]$s)
    if ($null -eq $s) { return "" }
    $t = $s.Replace('\', '\\').Replace('"', '\"').Replace("`r`n", '\n').Replace("`n", '\n').Replace("`r", '\n').Replace("`t", '\t')
    return $t
}

function Write-CrimsonBindVarsLuaFile {
    param([string]$OutPath, [System.Collections.ArrayList]$Rows)
    $manifest = @{}
    foreach ($r in @($Rows)) {
        if ($null -eq $r) { continue }
        $sec = if ($null -eq $r.Section) { "" } else { $r.Section.ToString().Trim() }
        $an = if ($null -eq $r.ActionName) { "" } else { $r.ActionName.ToString().Trim() }
        if (-not $manifest[$sec]) { $manifest[$sec] = [System.Collections.ArrayList]::new() }
        if ($an) { [void]$manifest[$sec].Add($an) }
    }
    $sb = [System.Text.StringBuilder]::new()
    [void]$sb.AppendLine("CrimsonBindVars = {")
    [void]$sb.AppendLine("  version = 1,")
    [void]$sb.AppendLine(('  sourceTimestamp = "' + (Get-Date -Format "yyyy-MM-ddTHH:mm:ss") + '",'))
    [void]$sb.AppendLine("  binds = {")
    $i = 0
    foreach ($r in @($Rows)) {
        if ($null -eq $r) { continue }
        $i++
        $mac = Escape-LuaString -s (Normalize-MacroTextForCrimsonBindLua -Raw $r.MacroText)
        $act = if ($null -eq $r.ActionName) { "" } else { $r.ActionName.ToString() }
        $sec = if ($null -eq $r.Section) { "" } else { $r.Section.ToString() }
        $key = if ($null -eq $r.Key) { "" } else { $r.Key.ToString().Trim() }
        $at = "macro"
        if ($mac -match '^\s*/cast') { $at = "spell" }
        elseif ($mac -match '^\s*/use') { $at = "item" }
        $src = if ($sec -ieq 'CUSTOM') { "custom" } else { "config" }
        [void]$sb.AppendLine("    [$i] = {")
        [void]$sb.AppendLine(('      section = "' + (Escape-LuaString -s $sec) + '",'))
        [void]$sb.AppendLine(('      actionName = "' + (Escape-LuaString -s $act) + '",'))
        [void]$sb.AppendLine(('      key = "' + (Escape-LuaString -s $key) + '",'))
        [void]$sb.AppendLine(('      macroText = "' + $mac + '",'))
        [void]$sb.AppendLine("      textureId = 0,")
        [void]$sb.AppendLine(('      actionType = "' + $at + '",'))
        [void]$sb.AppendLine(('      source = "' + $src + '",'))
        [void]$sb.AppendLine("    },")
    }
    [void]$sb.AppendLine("  },")
    [void]$sb.AppendLine("  abilityManifest = {")
    foreach ($sec in ($manifest.Keys | Sort-Object)) {
        [void]$sb.AppendLine(('    ["' + (Escape-LuaString -s $sec) + '"] = {'))
        $seen = @{}
        foreach ($an in $manifest[$sec]) {
            if ($seen[$an]) { continue }
            $seen[$an] = $true
            [void]$sb.AppendLine(('      "' + (Escape-LuaString -s $an) + '",'))
        }
        [void]$sb.AppendLine("    },")
    }
    [void]$sb.AppendLine("  },")
    [void]$sb.AppendLine("  pendingEdits = {},")
    [void]$sb.AppendLine("}")
    $dir = [System.IO.Path]::GetDirectoryName($OutPath)
    if ($dir -and -not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    $bak = $OutPath + ".bak"
    if (Test-Path -LiteralPath $OutPath) { Copy-Item -LiteralPath $OutPath -Destination $bak -Force }
    [System.IO.File]::WriteAllText($OutPath, $sb.ToString(), [System.Text.UTF8Encoding]::new($false))
}

function Unescape-LuaStr {
    param([string]$x)
    if (-not $x) { return "" }
    return ($x -replace '\\n', "`n" -replace '\\r', "`r" -replace '\\"', '"' -replace '\\\\', '\')
}

# Parses CrimsonBind.lua SavedVariables and returns all bind entries where source="custom" or section="CUSTOM".
function Read-CrimsonBindCustomRowsFromLua {
    param([string]$LuaPath)
    $list = [System.Collections.ArrayList]::new()
    if (-not (Test-Path -LiteralPath $LuaPath -PathType Leaf)) { return $list }
    $text = [System.IO.File]::ReadAllText($LuaPath, [System.Text.UTF8Encoding]::new($false))
    $blockRe = [regex]::new('\{([^{}]*?)\}', [System.Text.RegularExpressions.RegexOptions]::Singleline)
    $kvRe = [regex]::new('(\w+)\s*=\s*"((?:\\.|[^"\\])*)"')
    foreach ($block in $blockRe.Matches($text)) {
        $inner = $block.Groups[1].Value
        $kv = @{}
        foreach ($m in $kvRe.Matches($inner)) { $kv[$m.Groups[1].Value] = Unescape-LuaStr $m.Groups[2].Value }
        $sec = if ($kv['section']) { $kv['section'] } else { "" }
        $src = if ($kv['source']) { $kv['source'] } else { "" }
        if ($sec -ine 'CUSTOM' -and $src -ine 'custom') { continue }
        $an  = if ($kv['actionName']) { $kv['actionName'] } else { "" }
        if (-not $an) { continue }
        [void]$list.Add([PSCustomObject]@{
            Section    = "CUSTOM"
            ActionName = $an
            Key        = if ($kv['key']) { $kv['key'] } else { "" }
            MacroText  = if ($kv['macroText']) { $kv['macroText'] } else { "" }
        })
    }
    return $list
}

# Upserts CUSTOM rows from CrimsonBind.lua into the CSV (match by ActionName within CUSTOM section).
# Appends new rows, updates Key/MacroText for existing ones. Returns count of rows upserted.
function Export-CrimsonBindCustomRowsToCsv {
    param(
        [Parameter(Mandatory)][string]$LuaPath,
        [Parameter(Mandatory)][string]$CsvPath
    )
    $luaRows = @(Read-CrimsonBindCustomRowsFromLua -LuaPath $LuaPath)
    if ($luaRows.Count -eq 0) { return 0 }
    $tbl = if (Test-Path -LiteralPath $CsvPath -PathType Leaf) {
        try { @(Import-Csv -LiteralPath $CsvPath -Encoding UTF8 -ErrorAction Stop) } catch { @() }
    } else { @() }
    $byName = @{}
    foreach ($row in $tbl) {
        if ($null -eq $row) { continue }
        $sec = (Import-CsvRowProperty -Row $row -Names @('Section')).Trim()
        $an  = (Import-CsvRowProperty -Row $row -Names @('ActionName')).Trim()
        if ($sec -ieq 'CUSTOM' -and $an) { $byName[$an] = $row }
    }
    $upserted = 0
    $newRows = [System.Collections.ArrayList]::new()
    foreach ($lr in $luaRows) {
        $an = $lr.ActionName.Trim()
        if ($byName.ContainsKey($an)) {
            $existing = $byName[$an]
            foreach ($p in $existing.PSObject.Properties) {
                if ($p.Name -ieq 'Key') { $p.Value = $lr.Key }
                elseif ($p.Name -ieq 'MacroText') { $p.Value = $lr.MacroText }
            }
        } else {
            [void]$newRows.Add([PSCustomObject]@{
                Section = "CUSTOM"; ActionName = $an; MacroText = $lr.MacroText; Key = $lr.Key
            })
        }
        $upserted++
    }
    $combined = [System.Collections.ArrayList]::new()
    foreach ($r in $tbl) { [void]$combined.Add($r) }
    foreach ($r in $newRows) { [void]$combined.Add($r) }
    Export-CrimsonBindCsvTableUtf8Quoted -Rows $combined -Path $CsvPath
    return $upserted
}

function Read-CrimsonBindPendingEditsFromLua {
    param([string]$LuaPath)
    $list = [System.Collections.ArrayList]::new()
    if (-not (Test-Path -LiteralPath $LuaPath -PathType Leaf)) { return $list }
    $text = [System.IO.File]::ReadAllText($LuaPath, [System.Text.UTF8Encoding]::new($false))
    $re = [regex]::new('section\s*=\s*"((?:\\.|[^"\\])*)"\s*,\s*actionName\s*=\s*"((?:\\.|[^"\\])*)"\s*,\s*oldKey\s*=\s*"((?:\\.|[^"\\])*)"\s*,\s*newKey\s*=\s*"((?:\\.|[^"\\])*)"(?:\s*,\s*editType\s*=\s*"((?:\\.|[^"\\])*)")?(?:\s*,\s*oldMacro\s*=\s*"((?:\\.|[^"\\])*)")?(?:\s*,\s*newMacro\s*=\s*"((?:\\.|[^"\\])*)")?', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    foreach ($m in $re.Matches($text)) {
        $et = if ($m.Groups[5].Success) { (Unescape-LuaStr $m.Groups[5].Value) } else { "key" }
        $om = if ($m.Groups[6].Success) { (Unescape-LuaStr $m.Groups[6].Value) } else { $null }
        $nm = if ($m.Groups[7].Success) { (Unescape-LuaStr $m.Groups[7].Value) } else { $null }
        [void]$list.Add([PSCustomObject]@{
            Section    = (Unescape-LuaStr $m.Groups[1].Value)
            ActionName = (Unescape-LuaStr $m.Groups[2].Value)
            OldKey     = (Unescape-LuaStr $m.Groups[3].Value)
            NewKey     = (Unescape-LuaStr $m.Groups[4].Value)
            EditType   = $et
            OldMacro   = $om
            NewMacro   = $nm
        })
    }
    return $list
}

function Merge-PendingEditsIntoRows {
    param([System.Collections.ArrayList]$Rows, [System.Collections.ArrayList]$Pending)
    $n = 0
    foreach ($p in $Pending) {
        foreach ($r in $Rows) {
            if ($r.Section -eq $p.Section -and $r.ActionName -eq $p.ActionName) {
                $r.Key = $p.NewKey
                $et = if ($null -ne $p.EditType -and $p.EditType) { $p.EditType.ToString().Trim().ToLowerInvariant() } else { "key" }
                if (($et -eq "macro" -or $et -eq "both") -and $null -ne $p.NewMacro) {
                    $r.MacroText = $p.NewMacro.ToString()
                }
                $n++
                break
            }
        }
    }
    return $n
}

# If config.ini has [Main] WTFPath=, set WoW SavedVariables textbox to Account\<name>\SavedVariables.
function Sync-UpdateWtfTextBoxFromConfig {
    param([string]$ConfigPath, [System.Windows.Forms.TextBox]$TbWtf)
    if (-not $TbWtf) { return }
    if (-not $ConfigPath -or -not (Test-Path -LiteralPath $ConfigPath -PathType Leaf)) { return }
    $guess = Get-SavedVariablesDirFromConfigIni -ConfigPath $ConfigPath
    if ($guess) { $TbWtf.Text = $guess }
}

# --- WinForms ---

function Get-CrimsonBindAddonPath {
    param([string]$SavedVarsDir)
    if (-not $SavedVarsDir) { return $null }
    $p = $SavedVarsDir.Trim()
    for ($i = 0; $i -lt 4; $i++) {
        $p = [System.IO.Path]::GetDirectoryName($p)
        if (-not $p) { return $null }
    }
    return [System.IO.Path]::Combine($p, "Interface", "AddOns", "CrimsonBind")
}

function Show-SettingsDialog {
    param(
        [System.Windows.Forms.TextBox]$TbPool,
        [System.Windows.Forms.TextBox]$TbCfg,
        [System.Windows.Forms.TextBox]$TbCsv,
        [System.Windows.Forms.TextBox]$TbWtf,
        [System.Windows.Forms.Label]$StatusLabel
    )
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = "Settings"
    $dlg.Size = [System.Drawing.Size]::new(520, 302)
    $dlg.StartPosition = "CenterParent"
    $dlg.FormBorderStyle = "FixedDialog"
    $dlg.MaximizeBox = $false
    $dlg.MinimizeBox = $false

    $grp = New-Object System.Windows.Forms.GroupBox
    $grp.Text = "Key pool preset (overwrites bindable_keys.ini)"
    $grp.Location = [System.Drawing.Point]::new(12, 12)
    $grp.Size = [System.Drawing.Size]::new(480, 58)
    $dlg.Controls.Add($grp)

    $rbFull = New-Object System.Windows.Forms.RadioButton
    $rbFull.Text = "Full"
    $rbFull.Location = [System.Drawing.Point]::new(12, 22)
    $rbFull.Checked = $true
    $grp.Controls.Add($rbFull)

    $rbResto = New-Object System.Windows.Forms.RadioButton
    $rbResto.Text = "Right-hand / Resto"
    $rbResto.Location = [System.Drawing.Point]::new(80, 22)
    $grp.Controls.Add($rbResto)

    $rbMin = New-Object System.Windows.Forms.RadioButton
    $rbMin.Text = "Minimal"
    $rbMin.Location = [System.Drawing.Point]::new(240, 22)
    $grp.Controls.Add($rbMin)

    $btnWritePreset = New-Object System.Windows.Forms.Button
    $btnWritePreset.Text = "Write preset file"
    $btnWritePreset.Location = [System.Drawing.Point]::new(12, 84)
    $btnWritePreset.Size = [System.Drawing.Size]::new(140, 28)
    $dlg.Controls.Add($btnWritePreset)

    $btnExportCsv = New-Object System.Windows.Forms.Button
    $btnExportCsv.Text = "Export CSV from config"
    $btnExportCsv.Location = [System.Drawing.Point]::new(12, 124)
    $btnExportCsv.Size = [System.Drawing.Size]::new(168, 28)
    $dlg.Controls.Add($btnExportCsv)

    $btnOpenPool = New-Object System.Windows.Forms.Button
    $btnOpenPool.Text = "Open key pool"
    $btnOpenPool.Location = [System.Drawing.Point]::new(190, 124)
    $btnOpenPool.Size = [System.Drawing.Size]::new(120, 28)
    $dlg.Controls.Add($btnOpenPool)

    $btnImportFromWow = New-Object System.Windows.Forms.Button
    $btnImportFromWow.Text = "Import from WoW"
    $btnImportFromWow.Location = [System.Drawing.Point]::new(12, 164)
    $btnImportFromWow.Size = [System.Drawing.Size]::new(140, 28)
    $dlg.Controls.Add($btnImportFromWow)

    $btnMergeCustom = New-Object System.Windows.Forms.Button
    $btnMergeCustom.Text = "Merge CUSTOM → CSV"
    $btnMergeCustom.Location = [System.Drawing.Point]::new(12, 204)
    $btnMergeCustom.Size = [System.Drawing.Size]::new(168, 28)
    $btnMergeCustom.BackColor = [System.Drawing.Color]::FromArgb(255, 220, 200)
    $dlg.Controls.Add($btnMergeCustom)

    $lblMergeCustom = New-Object System.Windows.Forms.Label
    $lblMergeCustom.Text = "Reads CUSTOM binds from CrimsonBind.lua and upserts them into the CSV."
    $lblMergeCustom.Location = [System.Drawing.Point]::new(190, 208)
    $lblMergeCustom.Size = [System.Drawing.Size]::new(310, 20)
    $dlg.Controls.Add($lblMergeCustom)

    $btnCloseSettings = New-Object System.Windows.Forms.Button
    $btnCloseSettings.Text = "Close"
    $btnCloseSettings.Location = [System.Drawing.Point]::new(416, 244)
    $btnCloseSettings.Size = [System.Drawing.Size]::new(76, 28)
    $dlg.Controls.Add($btnCloseSettings)

    $btnWritePreset.Add_Click({
        $p = if ($rbResto.Checked) { "RestoRightHand" } elseif ($rbMin.Checked) { "Minimal" } else { "Full" }
        $path = $TbPool.Text.Trim()
        if (-not $path) { $path = $Script:DefaultKeyPoolPath }
        try {
            Write-BindableKeysIniPreset -Preset $p -Path $path
            if ($StatusLabel) { $StatusLabel.Text = "Wrote preset $p to:`n$path" }
        } catch {
            if ($StatusLabel) { $StatusLabel.Text = "Preset error: $($_.Exception.Message)" }
        }
    })

    $btnExportCsv.Add_Click({
        $cfg = $TbCfg.Text.Trim()
        if (-not $cfg -or -not (Test-Path -LiteralPath $cfg)) {
            if ($StatusLabel) { $StatusLabel.Text = "Select a valid config.ini path, then Export CSV." }
            return
        }
        $csv = $TbCsv.Text.Trim()
        if (-not $csv) { $csv = Get-CrimsonBindCsvDefaultPath -ConfigPath $cfg }
        try {
            Export-CrimsonBindCsvFromConfigIni -ConfigPath $cfg -CsvPath $csv
            $TbCsv.Text = $csv
            if ($StatusLabel) { $StatusLabel.Text = "Exported CSV:`n$csv`nSection, ActionName, Key from config.ini. MacroText column left empty." }
        } catch {
            if ($StatusLabel) { $StatusLabel.Text = "Export CSV failed: $($_.Exception.Message)" }
        }
    })

    $btnOpenPool.Add_Click({
        $path = $TbPool.Text.Trim()
        if (-not $path) { $path = $Script:DefaultKeyPoolPath }
        if (-not (Test-Path -LiteralPath $path)) { Write-BindableKeysIniPreset -Preset "Full" -Path $path }
        Start-Process notepad.exe -ArgumentList "`"$path`""
    })

    $btnImportFromWow.Add_Click({
        $cfg = $TbCfg.Text.Trim()
        $wtf = $TbWtf.Text.Trim()
        if (-not $cfg -or -not $wtf) {
            if ($StatusLabel) { $StatusLabel.Text = "Set config.ini and WoW SavedVariables paths." }
            return
        }
        $lua = Join-Path $wtf "CrimsonBind.lua"
        if (-not (Test-Path -LiteralPath $lua)) {
            if ($StatusLabel) { $StatusLabel.Text = "CrimsonBind.lua not found in SavedVariables." }
            return
        }
        $csvBk = $TbCsv.Text.Trim(); if (-not $csvBk) { $csvBk = Get-CrimsonBindCsvDefaultPath -ConfigPath $cfg }
        $poolBk = $TbPool.Text.Trim(); if (-not $poolBk) { $poolBk = $Script:DefaultKeyPoolPath }
        New-CrimsonBindBackup -ConfigPath $cfg -CsvPath $csvBk -KeyPoolPath $poolBk -SavedVarsDir $wtf | Out-Null
        try {
            $pending = Read-CrimsonBindPendingEditsFromLua -LuaPath $lua
            if ($pending.Count -eq 0) {
                if ($StatusLabel) { $StatusLabel.Text = "No pendingEdits in CrimsonBind.lua." }
                return
            }
            if (-not $Script:SyncRows) {
                $rows = Get-SyncConfigIniRows -ConfigPath $cfg
                Update-SyncRowsDecodedKeys -Rows $rows
                $Script:SyncRows = $rows
            }
            $merged = Merge-PendingEditsIntoRows -Rows $Script:SyncRows -Pending $pending
            Set-ConfigIniFromRows -ConfigPath $cfg -Rows $Script:SyncRows -ForceGeneralIniKeys | Out-Null
            Write-CrimsonBindVarsLuaFile -OutPath $lua -Rows $Script:SyncRows
            if ($StatusLabel) { $StatusLabel.Text = "Imported $merged pending edit(s) into config.ini and refreshed CrimsonBind.lua." }
        } catch {
            if ($StatusLabel) { $StatusLabel.Text = "Import failed: $($_.Exception.Message)" }
        }
    })

    $btnMergeCustom.Add_Click({
        $wtf = $TbWtf.Text.Trim()
        $csv = $TbCsv.Text.Trim()
        $cfg = $TbCfg.Text.Trim()
        if (-not $wtf -or -not (Test-Path -LiteralPath $wtf)) {
            if ($StatusLabel) { $StatusLabel.Text = "Set WoW SavedVariables folder before merging CUSTOM binds." }
            return
        }
        $lua = Join-Path $wtf "CrimsonBind.lua"
        if (-not (Test-Path -LiteralPath $lua)) {
            if ($StatusLabel) { $StatusLabel.Text = "CrimsonBind.lua not found in: $wtf" }
            return
        }
        if (-not $csv -and $cfg) { $csv = Get-CrimsonBindCsvDefaultPath -ConfigPath $cfg }
        if (-not $csv) {
            if ($StatusLabel) { $StatusLabel.Text = "Set CSV path before merging CUSTOM binds." }
            return
        }
        $poolBk = $TbPool.Text.Trim(); if (-not $poolBk) { $poolBk = $Script:DefaultKeyPoolPath }
        New-CrimsonBindBackup -ConfigPath $cfg -CsvPath $csv -KeyPoolPath $poolBk -SavedVarsDir $wtf | Out-Null
        try {
            $n = Export-CrimsonBindCustomRowsToCsv -LuaPath $lua -CsvPath $csv
            if ($n -eq 0) {
                if ($StatusLabel) { $StatusLabel.Text = "No CUSTOM binds found in CrimsonBind.lua (source=""custom"" or section=""CUSTOM"")." }
            } else {
                if ($StatusLabel) { $StatusLabel.Text = "Merged $n CUSTOM bind(s) from CrimsonBind.lua into:`n$csv" }
            }
        } catch {
            if ($StatusLabel) { $StatusLabel.Text = "Merge CUSTOM failed: $($_.Exception.Message)" }
        }
    })

    $btnCloseSettings.Add_Click({ $dlg.Close() })
    [void]$dlg.ShowDialog()
}

function Show-SyncCrimsonBindsMainForm {
    $paths = Read-SyncPathsIni
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Sync-CrimsonBinds"
    $form.Size = [System.Drawing.Size]::new(720, 560)
    $form.StartPosition = "CenterScreen"

    # Use [Type]::new(...) so $layoutY - 2 is never parsed as (array op_Subtraction).
    [int]$layoutY = 12
    $rowH = 28

    # Addon status row (top, before config.ini)
    $lblAddonStatus = New-Object System.Windows.Forms.Label
    $lblAddonStatus.Location = [System.Drawing.Point]::new(12, $layoutY)
    $lblAddonStatus.Size = [System.Drawing.Size]::new(400, 24)
    $lblAddonStatus.Text = "CrimsonBind addon: not checked"
    $form.Controls.Add($lblAddonStatus)

    $btnInstallAddon = New-Object System.Windows.Forms.Button
    $btnInstallAddon.Text = "Install CrimsonBind"
    $btnInstallAddon.Location = [System.Drawing.Point]::new(420, ($layoutY - 2))
    $btnInstallAddon.Size = [System.Drawing.Size]::new(152, 28)
    $btnInstallAddon.Visible = $false
    $form.Controls.Add($btnInstallAddon)

    $layoutY += 32
    $lblCfg = New-Object System.Windows.Forms.Label
    $lblCfg.Text = "config.ini:"
    $lblCfg.Location = [System.Drawing.Point]::new(12, $layoutY)
    $lblCfg.Size = [System.Drawing.Size]::new(100, 22)
    $form.Controls.Add($lblCfg)

    $tbCfg = New-Object System.Windows.Forms.TextBox
    $tbCfg.Location = [System.Drawing.Point]::new(115, ($layoutY - 2))
    $tbCfg.Size = [System.Drawing.Size]::new(480, 22)
    $tbCfg.Text = $paths.ConfigIni
    $form.Controls.Add($tbCfg)

    $btnCfg = New-Object System.Windows.Forms.Button
    $btnCfg.Text = "Browse..."
    $btnCfg.Location = [System.Drawing.Point]::new(605, ($layoutY - 4))
    $btnCfg.Size = [System.Drawing.Size]::new(90, 26)
    $form.Controls.Add($btnCfg)

    $layoutY += $rowH + 6
    $lblWtf = New-Object System.Windows.Forms.Label
    $lblWtf.Text = "WoW SavedVariables:"
    $lblWtf.Location = [System.Drawing.Point]::new(12, $layoutY)
    $lblWtf.Size = [System.Drawing.Size]::new(110, 22)
    $form.Controls.Add($lblWtf)

    $tbWtf = New-Object System.Windows.Forms.TextBox
    $tbWtf.Location = [System.Drawing.Point]::new(125, ($layoutY - 2))
    $tbWtf.Size = [System.Drawing.Size]::new(470, 22)
    $tbWtf.Text = $paths.WoWSavedVariables
    $form.Controls.Add($tbWtf)
    if ($paths.ConfigIni -and (Test-Path -LiteralPath $paths.ConfigIni.Trim())) {
        Sync-UpdateWtfTextBoxFromConfig -ConfigPath $paths.ConfigIni.Trim() -TbWtf $tbWtf
    }

    $btnWtf = New-Object System.Windows.Forms.Button
    $btnWtf.Text = "Browse..."
    $btnWtf.Location = [System.Drawing.Point]::new(605, ($layoutY - 4))
    $btnWtf.Size = [System.Drawing.Size]::new(90, 26)
    $form.Controls.Add($btnWtf)

    $layoutY += $rowH + 6
    $lblPool = New-Object System.Windows.Forms.Label
    $lblPool.Text = "bindable_keys.ini:"
    $lblPool.Location = [System.Drawing.Point]::new(12, $layoutY)
    $lblPool.Size = [System.Drawing.Size]::new(110, 22)
    $form.Controls.Add($lblPool)

    $tbPool = New-Object System.Windows.Forms.TextBox
    $tbPool.Location = [System.Drawing.Point]::new(125, ($layoutY - 2))
    $tbPool.Size = [System.Drawing.Size]::new(470, 22)
    $tbPool.Text = if ($paths.KeyPoolFile) { $paths.KeyPoolFile } else { $Script:DefaultKeyPoolPath }
    $form.Controls.Add($tbPool)

    $btnPool = New-Object System.Windows.Forms.Button
    $btnPool.Text = "Browse..."
    $btnPool.Location = [System.Drawing.Point]::new(605, ($layoutY - 4))
    $btnPool.Size = [System.Drawing.Size]::new(90, 26)
    $form.Controls.Add($btnPool)

    $layoutY += $rowH + 6
    $lblCsv = New-Object System.Windows.Forms.Label
    $lblCsv.Text = "CrimsonBind CSV:"
    $lblCsv.Location = [System.Drawing.Point]::new(12, $layoutY)
    $lblCsv.Size = [System.Drawing.Size]::new(110, 22)
    $form.Controls.Add($lblCsv)

    $tbCsv = New-Object System.Windows.Forms.TextBox
    $tbCsv.Location = [System.Drawing.Point]::new(125, ($layoutY - 2))
    $tbCsv.Size = [System.Drawing.Size]::new(470, 22)
    $csvDefault = if ($paths.CrimsonBindCsv) {
        $paths.CrimsonBindCsv
    } elseif ($paths.ConfigIni -and (Test-Path -LiteralPath $paths.ConfigIni.Trim())) {
        Get-CrimsonBindCsvDefaultPath -ConfigPath $paths.ConfigIni.Trim()
    } else {
        [System.IO.Path]::GetFullPath((Join-Path $ToolDir "CrimsonBind_binds.csv"))
    }
    $tbCsv.Text = $csvDefault
    $form.Controls.Add($tbCsv)

    $btnCsvBrowse = New-Object System.Windows.Forms.Button
    $btnCsvBrowse.Text = "Browse..."
    $btnCsvBrowse.Location = [System.Drawing.Point]::new(605, ($layoutY - 4))
    $btnCsvBrowse.Size = [System.Drawing.Size]::new(90, 26)
    $form.Controls.Add($btnCsvBrowse)

    $layoutY += $rowH + 8

    $btnBackupNow = New-Object System.Windows.Forms.Button
    $btnBackupNow.Text = "Backup now"
    $btnBackupNow.Location = [System.Drawing.Point]::new(12, $layoutY)
    $btnBackupNow.Size = [System.Drawing.Size]::new(100, 26)
    $form.Controls.Add($btnBackupNow)

    $layoutY += 34
    $lblSec = New-Object System.Windows.Forms.Label
    $lblSec.Text = "Randomize sections:"
    $lblSec.Location = [System.Drawing.Point]::new(12, $layoutY)
    $lblSec.Size = [System.Drawing.Size]::new(130, 20)
    $form.Controls.Add($lblSec)

    $clbSec = New-Object System.Windows.Forms.CheckedListBox
    $clbSec.Location = [System.Drawing.Point]::new(12, ($layoutY + 22))
    $clbSec.Size = [System.Drawing.Size]::new(680, 120)
    $clbSec.CheckOnClick = $true
    [void]$clbSec.Items.Add("(All sections)")
    $clbSec.SetItemChecked(0, $true)
    $form.Controls.Add($clbSec)

    $Script:ClbSecUpdating = $false

    $layoutY += 22 + 120 + 12

    $btnImportCsv = New-Object System.Windows.Forms.Button
    $btnImportCsv.Text = "Import CSV to WoW"
    $btnImportCsv.Location = [System.Drawing.Point]::new(12, $layoutY)
    $btnImportCsv.Size = [System.Drawing.Size]::new(156, 28)
    $form.Controls.Add($btnImportCsv)

    $btnRestore = New-Object System.Windows.Forms.Button
    $btnRestore.Text = "Restore backup..."
    $btnRestore.Location = [System.Drawing.Point]::new(176, $layoutY)
    $btnRestore.Size = [System.Drawing.Size]::new(130, 28)
    $form.Controls.Add($btnRestore)

    $btnSettings = New-Object System.Windows.Forms.Button
    $btnSettings.Text = "Settings..."
    $btnSettings.Location = [System.Drawing.Point]::new(314, $layoutY)
    $btnSettings.Size = [System.Drawing.Size]::new(90, 28)
    $form.Controls.Add($btnSettings)

    $layoutY += 36
    $btnLoad = New-Object System.Windows.Forms.Button
    $btnLoad.Text = "Load config.ini"
    $btnLoad.Location = [System.Drawing.Point]::new(12, $layoutY)
    $btnLoad.Size = [System.Drawing.Size]::new(130, 30)
    $form.Controls.Add($btnLoad)

    $btnRand = New-Object System.Windows.Forms.Button
    $btnRand.Text = "Randomize + Sync"
    $btnRand.Location = [System.Drawing.Point]::new(150, $layoutY)
    $btnRand.Size = [System.Drawing.Size]::new(140, 30)
    $form.Controls.Add($btnRand)

    $layoutY += 40
    $lblStatus = New-Object System.Windows.Forms.Label
    $lblStatus.Location = [System.Drawing.Point]::new(12, $layoutY)
    $lblStatus.Size = [System.Drawing.Size]::new(680, 96)
    $lblStatus.Text = "Load config.ini first. Import CSV to WoW (close WoW first) to push binds. Paths saved to sync_paths.ini."
    $form.Controls.Add($lblStatus)

    function Update-SectionCombo {
        $Script:ClbSecUpdating = $true
        $clbSec.Items.Clear()
        [void]$clbSec.Items.Add("(All sections)")
        $merged = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        if ($Script:SyncRows) {
            foreach ($r in $Script:SyncRows) {
                if ($null -eq $r) { continue }
                $t = if ($null -eq $r.Section) { "" } else { $r.Section.ToString().Trim() }
                if ($t -and -not ($t -ieq 'CUSTOM')) { [void]$merged.Add($t) }  # CUSTOM is never randomized
            }
        }
        $cfgPath = $tbCfg.Text.Trim()
        if ($cfgPath -and (Test-Path -LiteralPath $cfgPath)) {
            foreach ($sn in (Get-SyncConfigIniBindSectionNames -ConfigPath $cfgPath)) {
                if ($sn) { [void]$merged.Add($sn) }
            }
        }
        # General always first, then remaining sections sorted alphabetically
        if ($merged.Contains("General")) {
            [void]$clbSec.Items.Add("General")
            [void]$merged.Remove("General")
        }
        foreach ($ss in (@($merged) | Sort-Object)) {
            [void]$clbSec.Items.Add($ss)
        }
        for ($i = 0; $i -lt $clbSec.Items.Count; $i++) { $clbSec.SetItemChecked($i, $true) }
        $Script:ClbSecUpdating = $false
    }

    $clbSec.Add_ItemCheck({
        param($sender, $e)
        if ($Script:ClbSecUpdating) { return }
        if ($e.Index -eq 0) {
            $Script:ClbSecUpdating = $true
            $willBeChecked = $e.NewValue -eq 'Checked'
            for ($i = 1; $i -lt $clbSec.Items.Count; $i++) { $clbSec.SetItemChecked($i, $willBeChecked) }
            $Script:ClbSecUpdating = $false
        }
    })

    $btnCfg.Add_Click({
        $d = New-Object System.Windows.Forms.OpenFileDialog
        $d.Filter = "INI|*.ini|All|*.*"
        if ($d.ShowDialog() -eq "OK") {
            $tbCfg.Text = $d.FileName
            Sync-UpdateWtfTextBoxFromConfig -ConfigPath $d.FileName -TbWtf $tbWtf
        }
    })
    $btnWtf.Add_Click({
        $d = New-Object System.Windows.Forms.FolderBrowserDialog
        $d.Description = "Select WoW Account SavedVariables folder (contains BindPad.lua, etc.)"
        if ($d.ShowDialog() -eq "OK") {
            $tbWtf.Text = $d.SelectedPath
            Update-AddonStatus
        }
    })
    $btnPool.Add_Click({
        $d = New-Object System.Windows.Forms.OpenFileDialog
        $d.Filter = "INI|*.ini|All|*.*"
        if ($d.ShowDialog() -eq "OK") { $tbPool.Text = $d.FileName }
    })
    $btnCsvBrowse.Add_Click({
        $d = New-Object System.Windows.Forms.OpenFileDialog
        $d.Filter = "CSV|*.csv|All|*.*"
        if ($tbCsv.Text.Trim()) { $d.FileName = [System.IO.Path]::GetFileName($tbCsv.Text.Trim()) }
        $d.InitialDirectory = if ($tbCfg.Text.Trim() -and (Test-Path -LiteralPath $tbCfg.Text.Trim())) {
            [System.IO.Path]::GetDirectoryName($tbCfg.Text.Trim())
        } else { $ToolDir }
        if ($d.ShowDialog() -eq "OK") { $tbCsv.Text = $d.FileName }
    })

    $btnLoad.Add_Click({
        $cfg = $tbCfg.Text.Trim()
        if (-not $cfg -or -not (Test-Path -LiteralPath $cfg)) {
            $lblStatus.Text = "Select a valid config.ini path."
            return
        }
        try {
            $rows = Get-SyncConfigIniRows -ConfigPath $cfg
            Update-SyncRowsDecodedKeys -Rows $rows
            $csvPath = $tbCsv.Text.Trim()
            if (-not $csvPath) { $csvPath = Get-CrimsonBindCsvDefaultPath -ConfigPath $cfg }
            $mf = Merge-CrimsonBindCsvMacrosIntoRows -Rows $rows -CsvPath $csvPath
            $mc = Merge-CrimsonBindCsvCustomRowsIntoSyncRows -SyncRows $rows -CsvPath $csvPath
            Sync-SpecRowKeysToMatchGeneral -Rows $rows
            $Script:SyncRows = $rows
            Update-SectionCombo
            Sync-UpdateWtfTextBoxFromConfig -ConfigPath $cfg -TbWtf $tbWtf
            Write-SyncPathsIni -ConfigIni $cfg -WoWSavedVariables $tbWtf.Text.Trim() -KeyPoolFile $tbPool.Text.Trim() -CrimsonBindCsv $tbCsv.Text.Trim()
            if ($rows.Count -eq 0) {
                $lblStatus.Text = "Loaded config.ini: 0 bind rows (check sections / filters). WTFPath SavedVariables: $($tbWtf.Text)"
            } else {
                $extra = if ($mf -gt 0) { " Merged $mf macro field(s) from CSV (CSV wins, including clears)." } else { "" }
                $customNote = if ($mc -gt 0) { " + $mc CUSTOM bind(s)." } else { "" }
                $lblStatus.Text = "Loaded $($rows.Count) bind rows.$extra$customNote SavedVariables: $($tbWtf.Text)"
            }
        } catch {
            $lblStatus.Text = "Load failed: $($_.Exception.Message)"
        }
    })

    $btnRand.Add_Click({
        $cfg = $tbCfg.Text.Trim()
        $poolPath = $tbPool.Text.Trim()
        if (-not $poolPath) { $poolPath = $Script:DefaultKeyPoolPath }
        if (-not $cfg -or -not (Test-Path -LiteralPath $cfg)) {
            $lblStatus.Text = "Select config.ini and Load first (or fix path)."
            return
        }
        if (-not (Test-Path -LiteralPath $poolPath)) {
            $lblStatus.Text = "bindable_keys.ini not found. Use Write preset file or Browse."
            return
        }
        Sync-UpdateWtfTextBoxFromConfig -ConfigPath $cfg -TbWtf $tbWtf
        $wtf = $tbWtf.Text.Trim()
        if (-not $wtf -or -not (Test-Path -LiteralPath $wtf)) {
            $lblStatus.Text = "Set WoW SavedVariables folder (or add WTFPath= under [Main] in config.ini)."
            return
        }
        $csvBk = $tbCsv.Text.Trim(); if (-not $csvBk) { $csvBk = Get-CrimsonBindCsvDefaultPath -ConfigPath $cfg }
        New-CrimsonBindBackup -ConfigPath $cfg -CsvPath $csvBk -KeyPoolPath $poolPath -SavedVarsDir $wtf | Out-Null
        try {
            if (-not $Script:SyncRows) {
                $rows = Get-SyncConfigIniRows -ConfigPath $cfg
                Update-SyncRowsDecodedKeys -Rows $rows
                $csvBoot = $tbCsv.Text.Trim()
                if (-not $csvBoot) { $csvBoot = Get-CrimsonBindCsvDefaultPath -ConfigPath $cfg }
                [void](Merge-CrimsonBindCsvMacrosIntoRows -Rows $rows -CsvPath $csvBoot)
                [void](Merge-CrimsonBindCsvCustomRowsIntoSyncRows -SyncRows $rows -CsvPath $csvBoot)
                $Script:SyncRows = $rows
                Update-SectionCombo
            }
            if ($Script:SyncRows.Count -eq 0) {
                $lblStatus.Text = "No bind rows to randomize. Click Load config.ini first or fix config content."
                return
            }
            # Compute which sections to randomize from CheckedListBox
            $onlySections = $null
            $checkedItems = @($clbSec.CheckedItems)
            if ($checkedItems.Count -gt 0 -and -not ($checkedItems -contains "(All sections)")) {
                $onlySections = [string[]]($checkedItems | Where-Object { $_ -ne "(All sections)" -and $_ -ine "CUSTOM" })
                # Auto-include General if specific specs selected and General exists in data
                $hasGeneralRows = @($Script:SyncRows | Where-Object { $_.Section -eq "General" }).Count -gt 0
                if ($hasGeneralRows -and $onlySections -notcontains "General") {
                    $onlySections = @("General") + $onlySections
                }
            }
            $rowsCopy = [System.Collections.ArrayList]::new()
            foreach ($x in $Script:SyncRows) {
                if ($null -eq $x) { continue }
                [void]$rowsCopy.Add([PSCustomObject]@{
                    Section = $x.Section; ActionName = $x.ActionName; Key = $x.Key; MacroText = $x.MacroText; TextureID = $x.TextureID
                })
            }
            $csvPath = $tbCsv.Text.Trim()
            if (-not $csvPath) { $csvPath = Get-CrimsonBindCsvDefaultPath -ConfigPath $cfg }
            $mf = Merge-CrimsonBindCsvMacrosIntoRows -Rows $rowsCopy -CsvPath $csvPath
            Invoke-SyncRandomizeKeys -Rows $rowsCopy -OnlySections $onlySections -KeyPoolPath $poolPath -CrimsonBindCsvPath $csvPath
            # -ForceGeneralIniKeys: user explicitly asked to randomize, so write keys even into previously-empty [General] lines.
            $n = Set-ConfigIniFromRows -ConfigPath $cfg -Rows $rowsCopy -ForceGeneralIniKeys
            Update-SyncRowsDecodedKeys -Rows $rowsCopy
            $Script:SyncRows = $rowsCopy
            $out = Join-Path $wtf "CrimsonBind.lua"
            Write-CrimsonBindVarsLuaFile -OutPath $out -Rows $Script:SyncRows
            Write-SyncPathsIni -ConfigIni $cfg -WoWSavedVariables $wtf -KeyPoolFile $poolPath -CrimsonBindCsv $tbCsv.Text.Trim()
            $keyWrites = Write-CrimsonBindCsvKeysFromSyncRows -Rows $Script:SyncRows -CsvPath $csvPath
            $csvKeyNote = if ($keyWrites -gt 0) { " CSV: $keyWrites Key cell(s) updated." } else { "" }
            $secLabel = if ($onlySections) { $onlySections -join ", " } else { "all sections" }
            $mfNote = if ($mf -gt 0) { " Merged $mf macro field(s) from CSV before randomize." } else { "" }
            $lblStatus.Text = "Randomized ($secLabel); config.ini lines touched: $n$mfNote$csvKeyNote`nWrote CrimsonBind.lua"
        } catch {
            $ex = $_.Exception.Message
            $pos = if ($_.InvocationInfo.PositionMessage) { "`n" + $_.InvocationInfo.PositionMessage.Trim() } else { "" }
            $lblStatus.Text = "Randomize failed: $ex$pos"
        }
    })

    $btnBackupNow.Add_Click({
        $cfg = $tbCfg.Text.Trim()
        $csv = $tbCsv.Text.Trim(); if (-not $csv -and $cfg) { $csv = Get-CrimsonBindCsvDefaultPath -ConfigPath $cfg }
        $pool = $tbPool.Text.Trim(); if (-not $pool) { $pool = $Script:DefaultKeyPoolPath }
        $wtf = $tbWtf.Text.Trim()
        if (-not $cfg -and -not $csv -and -not $pool) {
            $lblStatus.Text = "Set at least one file path before backing up."
            return
        }
        try {
            $bkDir = New-CrimsonBindBackup -ConfigPath $cfg -CsvPath $csv -KeyPoolPath $pool -SavedVarsDir $wtf
            if ($bkDir) {
                $count = (Get-ChildItem -LiteralPath $bkDir -File | Where-Object { $_.Name -ne 'manifest.json' }).Count
                $lblStatus.Text = "Backed up $count file(s) to BACKUPS\$(Split-Path -Leaf $bkDir)"
            } else {
                $lblStatus.Text = "Nothing to back up (no files found at the configured paths)."
            }
        } catch {
            $lblStatus.Text = "Backup failed: $($_.Exception.Message)"
        }
    })

    $btnRestore.Add_Click({
        $backupsRoot = Join-Path $Script:ToolDir "BACKUPS"
        if (-not (Test-Path -LiteralPath $backupsRoot)) {
            $lblStatus.Text = "No BACKUPS folder found. Run an operation or click 'Backup now' first."
            return
        }
        Show-RestoreBackupDialog -BackupsRoot $backupsRoot -StatusLabel $lblStatus
    })

    function Update-AddonStatus {
        $wtfPath = $tbWtf.Text.Trim()
        $addonDir = Get-CrimsonBindAddonPath -SavedVarsDir $wtfPath
        if ($addonDir -and (Test-Path -LiteralPath (Join-Path $addonDir "CrimsonBind.toc") -PathType Leaf)) {
            $lblAddonStatus.Text = "CrimsonBind addon: installed"
            $lblAddonStatus.ForeColor = [System.Drawing.Color]::Green
            $btnInstallAddon.Visible = $false
        } elseif ($addonDir) {
            $lblAddonStatus.Text = "CrimsonBind addon: not found"
            $lblAddonStatus.ForeColor = [System.Drawing.Color]::OrangeRed
            $btnInstallAddon.Visible = $true
        } else {
            $lblAddonStatus.Text = "CrimsonBind addon: set SavedVariables path first"
            $lblAddonStatus.ForeColor = [System.Drawing.Color]::Gray
            $btnInstallAddon.Visible = $false
        }
    }

    $btnInstallAddon.Add_Click({
        $wtfPath = $tbWtf.Text.Trim()
        $addonDir = Get-CrimsonBindAddonPath -SavedVarsDir $wtfPath
        if (-not $addonDir) {
            $lblStatus.Text = "Set the WoW SavedVariables path first, then click Install CrimsonBind."
            return
        }
        $zipPath = Join-Path $Script:ToolDir "CrimsonBind.zip"
        if (-not (Test-Path -LiteralPath $zipPath -PathType Leaf)) {
            $lblStatus.Text = "CrimsonBind.zip not found next to script at:`n$zipPath"
            return
        }
        try {
            $addonsParent = Split-Path $addonDir -Parent
            if (-not (Test-Path -LiteralPath $addonsParent)) {
                [void](New-Item -Path $addonsParent -ItemType Directory -Force)
            }
            Expand-Archive -LiteralPath $zipPath -DestinationPath $addonsParent -Force
            Update-AddonStatus
            $lblStatus.Text = "CrimsonBind addon installed to:`n$addonDir`nRestart WoW to activate it."
        } catch {
            $lblStatus.Text = "Install failed: $($_.Exception.Message)"
        }
    })

    $btnSettings.Add_Click({
        Show-SettingsDialog -TbPool $tbPool -TbCfg $tbCfg -TbCsv $tbCsv -TbWtf $tbWtf -StatusLabel $lblStatus
    })

    # Run addon detection on startup if WTF path is already set
    if ($tbWtf.Text.Trim()) { Update-AddonStatus }

    $form.Add_FormClosing({
        Write-SyncPathsIni -ConfigIni $tbCfg.Text.Trim() -WoWSavedVariables $tbWtf.Text.Trim() -KeyPoolFile $tbPool.Text.Trim() -CrimsonBindCsv $tbCsv.Text.Trim()
    })

    $btnImportCsv.Add_Click({
        $cfg = $tbCfg.Text.Trim()
        $wtf = $tbWtf.Text.Trim()
        $csv = $tbCsv.Text.Trim()
        if (-not $csv) {
            if ($cfg -and (Test-Path -LiteralPath $cfg)) {
                $csv = Get-CrimsonBindCsvDefaultPath -ConfigPath $cfg
            } else {
                $lblStatus.Text = "Set CSV path or config.ini (for default CrimsonBind_binds.csv)."
                return
            }
        }
        if (-not (Test-Path -LiteralPath $csv)) {
            $lblStatus.Text = "CSV not found: $csv"
            return
        }
        if (-not $wtf -or -not (Test-Path -LiteralPath $wtf)) {
            $lblStatus.Text = "Set WoW SavedVariables folder before Import CSV."
            return
        }
        $poolBk = $tbPool.Text.Trim(); if (-not $poolBk) { $poolBk = $Script:DefaultKeyPoolPath }
        New-CrimsonBindBackup -ConfigPath $cfg -CsvPath $csv -KeyPoolPath $poolBk -SavedVarsDir $wtf | Out-Null
        try {
            $tmFilled = Update-CrimsonBindCsvFileTargetMemberDefaultMacros -CsvPath $csv
            $rows = Import-CrimsonBindCsvToRows -CsvPath $csv
            if ($rows.Count -eq 0) {
                $lblStatus.Text = "CSV had no data rows (need Section,ActionName,MacroText,Key headers)."
                return
            }
            Update-SyncRowsDecodedKeys -Rows $rows
            Sync-SpecRowKeysToMatchGeneral -Rows $rows
            $Script:SyncRows = $rows
            Update-SectionCombo
            if ($cfg -and (Test-Path -LiteralPath $cfg)) {
                Set-ConfigIniFromRows -ConfigPath $cfg -Rows $rows -ForceGeneralIniKeys | Out-Null
            }
            $out = Join-Path $wtf "CrimsonBind.lua"
            Write-CrimsonBindVarsLuaFile -OutPath $out -Rows $Script:SyncRows
            Write-SyncPathsIni -ConfigIni $cfg -WoWSavedVariables $wtf -KeyPoolFile $tbPool.Text.Trim() -CrimsonBindCsv $csv
            $pre = if ($tmFilled -gt 0) { "Filled $tmFilled empty Target Member macro(s); CSV re-saved with quoted fields (Excel-safe).`n`n" } else { "" }
            $lblStatus.Text = $pre + "Imported $($rows.Count) row(s) from CSV.`nWrote CrimsonBind.lua (macros from CSV); config.ini keys only (macro tails after ';' preserved).`n$out`n(Reload UI in WoW.)"
        } catch {
            $lblStatus.Text = "Import CSV failed: $($_.Exception.Message)"
        }
    })

    [void]$form.ShowDialog()
}

if ($ImportCsvToWow) {
    $paths = Read-SyncPathsIni
    $cfg = $ConfigIni
    if (-not $cfg) { $cfg = $paths.ConfigIni }
    $wtf = $WoWSavedVariables
    if (-not $wtf) { $wtf = $paths.WoWSavedVariables }
    $csv = $CrimsonBindCsv
    if (-not $csv) { $csv = $paths.CrimsonBindCsv }
    if (-not $csv -and $cfg) { $csv = Get-CrimsonBindCsvDefaultPath -ConfigPath $cfg }
    if (-not $csv) { $csv = [System.IO.Path]::GetFullPath((Join-Path $ToolDir "CrimsonBind_binds.csv")) }
    $csv = $csv.Trim()
    if (-not (Test-Path -LiteralPath $csv -PathType Leaf)) {
        throw "ImportCsvToWow: CSV not found: $csv"
    }
    if (-not $wtf -or -not (Test-Path -LiteralPath $wtf -PathType Container)) {
        throw "ImportCsvToWow: WoW SavedVariables folder missing. Set WoWSavedVariables in sync_paths.ini or pass -WoWSavedVariables."
    }
    if (-not $NoBackup) {
        $pool = $paths.KeyPoolFile; if (-not $pool) { $pool = $Script:DefaultKeyPoolPath }
        $bkDir = New-CrimsonBindBackup -ConfigPath $cfg -CsvPath $csv -KeyPoolPath $pool -SavedVarsDir $wtf
        if ($bkDir) { Write-Host "CrimsonBind: backup -> $(Split-Path -Leaf $bkDir)" }
    }
    $tmFilled = Update-CrimsonBindCsvFileTargetMemberDefaultMacros -CsvPath $csv
    $rows = Import-CrimsonBindCsvToRows -CsvPath $csv
    if ($rows.Count -eq 0) {
        throw "ImportCsvToWow: CSV had no data rows."
    }
    if ($tmFilled -gt 0) {
        Write-Host "CrimsonBind: filled $tmFilled empty Target Member macro(s) in CSV; file re-saved with quoted fields (Excel-safe)."
    }
    Update-SyncRowsDecodedKeys -Rows $rows
    Sync-SpecRowKeysToMatchGeneral -Rows $rows
    if ($cfg -and (Test-Path -LiteralPath $cfg -PathType Leaf)) {
        Set-ConfigIniFromRows -ConfigPath $cfg -Rows $rows -ForceGeneralIniKeys | Out-Null
    }
    $out = Join-Path $wtf "CrimsonBind.lua"
    Write-CrimsonBindVarsLuaFile -OutPath $out -Rows $rows
    $pool = $paths.KeyPoolFile
    if (-not $pool) { $pool = $Script:DefaultKeyPoolPath }
    Write-SyncPathsIni -ConfigIni $cfg -WoWSavedVariables $wtf -KeyPoolFile $pool -CrimsonBindCsv $csv
    Write-Host "CrimsonBind ImportCsvToWow: $($rows.Count) row(s) -> $out"
    Write-Host "Reload UI in WoW (/reload)."
    return
}

Show-SyncCrimsonBindsMainForm
