# CrimsonBind

Keybind management for World of Warcraft (Retail). Manages hundreds of macro-text binds across classes, specs, and a shared General pool, with an external sync tool that randomizes keys and keeps your hotkey profile, CSV source of truth, and WoW SavedVariables in sync.

## What is CrimsonBind?

CrimsonBind is two things:

1. **A WoW addon** (`CrimsonBind/`) that reads bind data from SavedVariables, applies them via secure macros (`CLICK CrimsonBindMacro:CBn`), and provides an in-game panel to browse, search, edit, and validate every bind.
2. **A Windows sync tool** (`Sync-CrimsonBinds.ps1`) with a WinForms GUI that reads your hotkey profile (`config.ini`), optionally randomizes keys from a configurable pool, merges macro text from a CSV, writes `CrimsonBind.lua` into WoW's SavedVariables, and updates `config.ini` so every tool sees the same keys.

Unlike BindPad or Debounce, CrimsonBind does not require you to manually drag icons or define binds one by one inside WoW. Instead it bulk-imports every class/spec bind from an external profile and lets you manage macro text in a CSV spreadsheet — while still supporting in-game creation of custom macros.

## Quick start

### Install the addon

Copy the `CrimsonBind/` folder to `World of Warcraft/_retail_/Interface/AddOns/CrimsonBind`.

### Sync binds from your hotkey profile

1. Run `Sync-CrimsonBinds.ps1` (requires PowerShell 5.1+).
2. Set the paths for `config.ini`, WoW SavedVariables folder, and optionally `bindable_keys.ini` and `CrimsonBind_binds.csv`.
3. Click **Load config.ini**, then **Randomize + Sync** (or **Import CSV to WoW** if your CSV already has keys).
4. `/reload` in WoW. CrimsonBind auto-applies binds on login.

### Open the panel in WoW

Type `/cb` or `/crimsonbind`, or bind **Toggle CrimsonBind** in WoW's Key Bindings menu.

## How binds work

### Sections

Every bind belongs to a **Section**:

| Section | When it applies |
|---------|----------------|
| **General** | Always active for every class and spec. |
| **Class - Spec** (e.g. `Druid - Restoration`) | Active only when that spec is your current specialization. Binds with the same ActionName as a General row inherit General's key. |
| **CUSTOM** | Always active, regardless of spec. User-created macros that are never written to `config.ini` and never touched by the randomizer. |

On login and on spec change, CrimsonBind applies: **General** first, then **current spec** (overrides General on key collision), then **CUSTOM** (overrides everything on key collision).

### Macro text

Macro bodies are stored in `CrimsonBind_binds.csv` (the source of truth). The CSV uses `§` as a line separator instead of real newlines so it stays one-row-per-bind. The addon normalises `§` back to `\n` at runtime.

Each bind is applied via a hidden secure button with `SetAttribute("*macrotext-CBn", body)` and `SetBinding(KEY, "CLICK CrimsonBindMacro:CBn")` — no WoW macro slots are consumed.

### Key format

Keys follow WoW's `SetBinding` names: `CTRL-SHIFT-T`, `ALT-F5`, `BUTTON4`, etc. Numpad keys and `CTRL-MINUS` / bare `MINUS` are blocked (unreliable in WoW's binding system).

## In-game panel

The panel (`/cb`) shows a scrollable list with columns: **Status**, **Section**, **Action**, **Key**, and an **Edit** button per row.

### Status indicators

| Colour | Meaning |
|--------|---------|
| Green | OK — key is bound and macro is valid. |
| Yellow | Duplicate key within the same section, or key conflicts with a General bind. |
| Red | Error — empty macro with a key, macro exceeds 255 chars, plain text that would be /say, numpad key, or CTRL-MINUS. |
| Grey | Inactive — belongs to a different spec (not applied). |

### Filters

- **Search**: free-text filter across action name, macro text, section name, and key.
- **Section dropdown**: filter to a single section.
- **This spec + General** checkbox: only show rows that are currently active (General + current spec + CUSTOM).
- **Issues only** checkbox: only show rows with validation warnings.

### Inline editing

- Click a key cell to capture a new key (ESC cancels, DELETE clears).
- Click **...** (Edit) to open the bind editor with key capture, macro text editing, character count (255 limit), /say detection warning, and Save / Revert / Cancel buttons.
- For **CUSTOM** rows the editor also shows a Name field (rename on Save) and a Delete button with confirmation.

### Undo

`/cb undo` reverts the last key or macro edit (up to 20 levels). Undo entries are stored per session in SavedVariables.

### Test mode

`/cb test` or the panel checkbox toggles test mode. In test mode every bound key prints the stored key + action name to chat instead of executing the macro — useful for verifying which physical key maps to which bind.

## Custom macros (CUSTOM section)

Click **New macro** in the panel toolbar to create a new CUSTOM-section bind. CUSTOM binds:

- Are always active regardless of your current spec.
- Are never written to `config.ini`.
- Are never re-keyed by the randomizer.
- Reserve their keys so the randomizer cannot assign them to other binds.
- Can be exported to CSV via `/cb exportcustom` (prints quoted CSV rows to chat) or via the Sync tool's **Settings > Merge CUSTOM from WoW -> CSV** button.
- Round-trip: create in-game, export to CSV, Import CSV to WoW — keys and macros preserved.

## Sync tool (Sync-CrimsonBinds.ps1)

### GUI buttons

| Button | What it does |
|--------|-------------|
| **Load config.ini** | Parses the hotkey profile, decodes vk/sc key tokens, merges macro text from CSV, appends CUSTOM rows from CSV. |
| **Randomize + Sync** | Shuffles keys from the pool (per-section or all), writes `config.ini` + `CrimsonBind.lua` + updates CSV keys. Respects General-reserved and CUSTOM-reserved keys. |
| **Import CSV to WoW** | Writes `CrimsonBind.lua` from CSV (macros are source of truth), updates `config.ini` keys only. |
| **Settings** | Write key pool presets, export CSV from config, open key pool in Notepad, import pendingEdits from WoW, merge CUSTOM from WoW to CSV. |

### Key pool (`bindable_keys.ini`)

The randomizer draws from `[Modifiers]` x `[BaseKeys]`, minus `[ExcludedKeys]`, minus any key whose base is in `[ExcludedBaseKeys]`, minus numpad. Three presets: **Full**, **Right-hand / Resto**, and **Minimal**.

### CSV format

```
"Section","ActionName","MacroText","Key"
"General","TargetMouseOver","/target mouseover","CTRL-N"
"Druid - Restoration","Rejuvenation","/cast Rejuvenation","ALT-H"
"CUSTOM","My Arena Burst","/cast Incarnation§/cast Nature's Vigil","SHIFT-G"
```

`§` separates macro lines. All four columns are always quoted (safe for Excel when macros contain `;`).

### Headless mode

```powershell
.\Sync-CrimsonBinds.ps1 -ImportCsvToWow
```

Reads `sync_paths.ini` (auto-saved by the GUI), imports CSV to WoW SavedVariables, and updates `config.ini` — no GUI window. Useful for scripting or scheduled tasks.

### Backup

Every destructive operation (Randomize, Import, Merge CUSTOM) creates a timestamped backup in `BACKUPS/` with copies of `config.ini`, `CrimsonBind_binds.csv`, `bindable_keys.ini`, and `CrimsonBind.lua`.

## Slash commands

| Command | Description |
|---------|-------------|
| `/cb` or `/crimsonbind` | Toggle the panel. |
| `/cb apply` | Apply binds (confirms if overwriting WoW keys). |
| `/cb validate` | Print all validation issues to chat. |
| `/cb test` | Toggle test mode. |
| `/cb count` | Print bind count and current section. |
| `/cb list` | Print active binds with slot indices. |
| `/cb exportcustom` | Print CUSTOM binds as CSV rows for pasting. |
| `/cb clearbindpad` | Unbind keys still pointing at BindPad CLICK commands. |
| `/cb undo` | Undo last key or macro edit. |
| `/cb debug` | Toggle debug logging. |
| `/cb status` | Print pending edits, combat queue, test mode state. |
| `/cb help` | Print all commands. |

## How CrimsonBind compares

| Feature | BindPad | Debounce | CrimsonBind |
|---------|---------|----------|-------------|
| Bind source | Drag-and-drop icons in WoW | WoW UI + conditions | External profile (config.ini) + CSV + in-game CUSTOM |
| Macro storage | BindPad Macros (SavedVariables) | Macro Texts (SavedVariables) | CSV (source of truth) + SavedVariables |
| Key assignment | Click icon, press key | Left-click action, press key | Inline capture in list, or editor |
| Spec switching | Profile tabs (manual) | Tabs: shared / char / spec / class | Automatic: General + current spec + CUSTOM |
| Randomization | No | No | Yes — full key pool with exclusions |
| Bulk import | No | No | Yes — from config.ini or CSV |
| Validation | No | No | Yes — duplicates, conflicts, /say detection, numpad, length |
| External sync tool | No | No | Yes — WinForms GUI + headless mode |
| Conditional bindings | No | Yes (hover, combat, shapeshift, custom states) | No (use macro conditionals) |
| Click casting | No | Yes | No |

## Potential future features

These are features inspired by BindPad and Debounce that could extend CrimsonBind:

- **Icon picker for CUSTOM macros**: choose a spell/item icon for custom binds (cosmetic; display in the panel list and editor).
- **Drag-to-action-bar**: allow dragging a CUSTOM macro onto a WoW action bar slot (BindPad supports this for its macros).
- **Per-bind conditional states**: like Debounce's hover/combat/shapeshift conditions on individual binds, evaluated at key-press time via SecureHandler attributes rather than embedded in macro text.
- **Import from BindPad / Debounce**: one-click migration of existing BindPad Macros or Debounce Macro Texts into CUSTOM rows.
- **Profile snapshots**: save/restore named sets of CUSTOM binds (e.g. "PvP burst", "M+ utility") and switch between them.
- **Conflict map / key heatmap**: visual grid showing which modifier+base combinations are used, free, or conflicting across all sections.
- **Multi-character CSV**: support per-character CUSTOM rows (currently all CUSTOM binds are account-wide).

## File structure

```
CrimsonBind/
  CrimsonBind.lua        # WoW addon
  CrimsonBind.toc        # Addon manifest
  Bindings.xml           # Toggle CrimsonBind keybinding
Sync-CrimsonBinds.ps1   # Sync tool (WinForms GUI + headless)
CrimsonBind_binds.csv    # Macro source of truth
bindable_keys.ini        # Randomizer key pool
sync_paths.ini           # Saved GUI paths (auto-generated)
config.ini               # Hotkey profile (read/written by sync tool)
BACKUPS/                 # Timestamped backups
tasks/
  todo.md                # Development checklist
  lessons.md             # Post-correction notes
tools/                   # Utility scripts (parse checks, migrations)
```

## Requirements

- **WoW Retail** (Interface 120000+). Not tested on Classic / Mists.
- **PowerShell 5.1+** for the sync tool (ships with Windows 10/11).
- A hotkey profile `config.ini` from ggloader or compatible tool (optional if you only use CSV + CUSTOM macros).

## License

Personal use. Not distributed on CurseForge or other addon repositories.
