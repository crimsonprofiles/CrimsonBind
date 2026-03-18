# \# CrimsonBind

# 

# A tool to manage keybinds. Export binds to CSV, edit in a grid, randomize keys, push changes to your config and addon files. Works with \*\*custom CSVs\*\* so you can adapt it for other providers’ keybinds.

# 

# \*\*Only use the script from this repository.\*\* Other sources could host modified versions. The code is not obfuscated—read and adapt it as you like.

# 

# \---

# 

# \## Features

# 

# \- \*\*Export / import\*\* – Export keybinds from `config.ini` to CSV, or load your own CSV

# \- \*\*Data grid\*\* – View and edit binds (Section, Action, Key, MacroText) with filter and search

# \- \*\*Key dropdown\*\* – Pick from a defined key pool instead of typing; exclude keys from randomize

# \- \*\*Randomize keys\*\* – Assign keys from a pool by section; respects exclusions list

# \- \*\*Update config.ini\*\* – Write key changes from the grid/CSV back to your Crimson config

# \- \*\*Update Debounce\*\* – Write binds into WoW’s Debounce addon (General + class/spec tabs); supports filtered section or all

# \- \*\*Backup / restore\*\* – One-click backup of config.ini and Debounce/BindPad SavedVariables; restore from a chosen backup

# \- \*\*Exclusions\*\* – Mark key combinations to exclude from Randomize (saved to `excluded\_keys.csv`)

# 

# \---

# 

# \## Requirements

# 

# \- \*\*Windows\*\*

# \- \*\*PowerShell 5.1+\*\* (built-in on modern Windows)

# 

# \---



# \## Addon support

# 

# | Addon   | Status | Notes                                      |

# |--------|--------|--------------------------------------------|

# | Debounce | ✅     | General + class/spec tabs; Midnight-compatible version included |

# 

# More addons may be added later. The tool is built so new targets can be wired in.

# 

# \## Custom CSVs and other providers

# 

# The tool expects a CSV with columns like: \*\*Section\*\*, \*\*ActionName\*\*, \*\*MacroText\*\*, \*\*BindPadKey\*\* (used as Key), \*\*TextureID\*\*. You can use your own CSV with the same structure and point the tool at it—no need to use Crimson’s config.ini as the source. That makes it easy to adapt for other keybind providers or your own spreadsheets.

# 

# \---

# 

# \## Security and distribution

# 

# \- \*\*Only use the script from this repository\*\* so you’re running the intended, supportable version.

# \- The script is \*\*not encrypted or obfuscated\*\*; you can inspect and modify it.

# \- If you share or fork it, point users back here so they get updates and the same security note.

# 

# \---

# 

# \## License

# 

# Copyright (c) 2026 Crimson Profiles

# 

# Permission is granted to use, copy, modify, and distribute this software for any purpose, provided that proper credit is given to Crimson Profiles.

# 

# Attribution must include:

# \- The original author name (Crimson Profiles)

# \- A link to the original source (if applicable)

# 

# This software is provided "as is", without warranty of any kind.



