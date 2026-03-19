# CrimsonBind

> 🎯 A powerful tool for managing, editing, and randomising keybinds — with full CSV support and WoW addon integration.

---

## ✨ Overview

**CrimsonBind** is a flexible keybind management tool designed to:

- Export and import keybinds
- Edit them in a structured grid
- Randomise intelligently
- Push updates back into config and addon files

It supports **custom CSV formats**, making it adaptable beyond just Crimson configs.

---

## ⚠️ Important Notice

> **Only use the script from this repository.**

Other sources may distribute modified or unsafe versions.

---

## Changelog

### 1.02
- **Update Debounce:** Fixed "Cannot index into a null array" when updating Debounce with a partial CSV (e.g. General + one spec). Root cause: `$newBindSections` was never initialized in `Export-DebounceFromRows`.
- **Update Debounce:** Added defensive null-safety (initialized hashtables/arrays, safe iteration over bind sections and class tabs, null guards after reading existing Debounce file and in the success-path UI).
- **Update Debounce:** On error, the status bar now shows the first line of the script stack trace to simplify debugging.

### 1.01
- General keybinds button, version in window titles, randomize skips General-bound actions, Debounce keys/macros and feedback, README and CSV updates (see earlier conversation).

---

- 🔍 The code is **fully readable (not obfuscated)**
- 🛠 You are free to inspect and adapt it
- 🔐 Always verify the source before running scripts

---

## How General and spec bindings work

Many actions exist in both the **General** section and in each class/spec section (e.g. racials, Target Member1–40, Focus Party/Arena, Trinket Rotation, Universal, Potion). The script treats General as the single source of truth for those shared actions.

### Key and macro source

- When writing to **config.ini** or validating duplicates, the script builds a key map: for every section, if a row’s **ActionName** exists in General, it uses **General’s Key and MacroText** for that action in that section. Spec rows with the same action name do not override General; they effectively “inherit” General’s bind.
- So: configure shared binds (targeting, focus, racials, consumables, etc.) once in **General**; that key and macro apply to every spec that has that action.

### Grid behaviour when a spec is selected

- When you filter the grid by a class/spec (e.g. Paladin - Retribution), rows whose action exists in General show **General’s Key and MacroText** in the grid (read-only display for those rows).
- **Saving from the grid** (e.g. after editing) only writes back Key/MacroText for **spec-only** rows. Rows that inherit from General are not updated from the grid when you’re viewing a spec — so you don’t accidentally overwrite them. To change a shared bind, switch to the **General** section and edit it there.

### Apply general keybinds (grid toolbar)

- The **Apply general keybinds** button copies General’s **Key** and **MacroText** to every non-General row that has the same **ActionName**. Use it after you’ve edited General and want all specs to match without opening each section. Then use **Save to CSV** to persist.

### Duplicate detection

- A “duplicate” is one key bound to two or more different actions **in the same section**. The duplicate list and red highlighting are per-section.
- Rows that inherit from General (same ActionName as a General row) are **excluded** from duplicate counting in that section, so they don’t create false duplicates. The effective key for those rows is General’s key.

---

## 🚀 Features

### 📤 Import / Export
- Export keybinds from `config.ini` → CSV  
- Load and edit your own custom CSV files  

### 🧩 Data Grid Editor
- Edit:
  - Section  
  - Action  
  - Key  
  - MacroText  
- Filter by section and search  
- Key column has a **blank option at the top** of the dropdown to clear the key  

### 🎹 Key Selection
- Dropdown-based key selection  
- Prevent invalid/manual entry issues  
- Supports **exclusion lists**  
- F1–F12 and common modifier combos are always in the list (for General target/focus binds)  

### 🎲 Smart Randomisation
- **Randomize keys** (main form, with a section selected): randomises only the **spec-only** rows in that section. Rows that inherit from General (Target Member, racials, Universal, etc.) are **skipped** so their keys are not overwritten and duplicates are not introduced. Each randomised row gets a unique key from the pool (General’s keys are excluded when the section is a spec).
- **Randomize selected** (grid toolbar): only the selected rows that are **not** General-bound actions get new keys. Same skip rule and pool rules as above.
- **Randomize all** (main form, no section filter): General is randomised first from the full key pool; then each spec is randomised from a pool that excludes General’s keys (no duplicates within a section, and no overlap with General).
- Respects **exclusion lists** (excluded key combos are not assigned).  

### ⚙️ Config Integration
- Write changes back to:
  - `config.ini`
  - WoW addon files  

### 🔌 Addon Support
- Push bindings into:
  - Debounce addon (General + Class/Spec tabs)

### 💾 Backup & Restore
- One-click backups for:
  - `config.ini`
  - SavedVariables (Debounce / BindPad)
- Restore from previous snapshots  

### 🚫 Exclusions System
- Prevent specific key combos from being used when randomising or assigning keys.  
- Excluded keys are stored in a file next to the script and are respected by all randomise operations.
