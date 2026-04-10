--[[
  CrimsonBind (Retail only): applies macrotext binds from CrimsonBindVars (Synced via Sync-CrimsonBinds.ps1).
  Structured panel: columns, validation colors, inline key capture, section grouping.
]]

-- =========================
-- State
-- =========================

local CB = {
  prevBoundKeys = {},
  pendingApply = false,
  rowFrames = {},
  filtered = {},
  displayList = {},
  scrollOffset = 0,
  ROW_HEIGHT = 24,
  VISIBLE_ROWS = 14,
  searchText = "",
  sectionFilter = nil,
  conflictsOnly = false,
  filterToCurrentSpec = true,
  searchDebounceTimer = nil,
  capturingKeyForIndex = nil,
  keyCaptureToEditor = nil,
  rowStatus = {},
  dupCount = 0,
  issueCount = 0,
  externalActionByKey = {},
  testMode = false,
  testModeCheck = nil,
  testModeBanner = nil,
  lastSlotCount = 0,
  cachedApplyList = nil,
  emptyLabel = nil,
  debug = false,
  undoStack = {},
  editorFrame = nil,
  editorBindIndex = nil,
  editorSnapshotKey = nil,
  editorSnapshotMacro = nil,
  editorKeyEdit = nil,
  editorMacroEdit = nil,
  editorCharLabel = nil,
  captureOverlay = nil,
  captureKeyboard = nil,
  captureBanner = nil,
  iconCache = {},
  heatmapFrame = nil,
}

local COL_STATUS = 14
local COL_ICON = 22
local COL_SECTION = 104
local COL_ACTION = 164
local COL_KEY = 108
local COL_EDIT = 36

local COLOR_SECTION = "|cff888888"
local COLOR_ACTION = "|cffffffff"
local COLOR_KEY = "|cff66ccff"
local COLOR_KEY_HOT = "|cff99eeff"
local COLOR_OK = "|cff33cc33"
local COLOR_WARN = "|cffffcc00"
local COLOR_ERR = "|cffff3333"
local COLOR_MISS = "|cffaaaaaa"

-- =========================
-- Utilities: keys and text
-- =========================

local function trim(s)
  if not s then return "" end
  return (s:gsub("^%s+", ""):gsub("%s+$", ""))
end

--- Sync / config.ini / CSV use U+00A7 (§) for line breaks in stored macros; WoW needs real newlines. Strip §§§ trailing notes.
--- § may appear as UTF-8 (C2 A7) or as a single Latin-1/CP1252 byte (A7). Scan UTF-8 sequences so we do not break other 2-byte chars (e.g. Cyrillic).
local function normalizeStoredMacroText(s)
  s = trim(s or "")
  local tripleUtf8 = "\194\167\194\167\194\167"
  local pos = s:find(tripleUtf8, 1, true)
  if pos then
    s = trim(s:sub(1, pos - 1))
  end
  local tripleLatin = "\167\167\167"
  pos = s:find(tripleLatin, 1, true)
  if pos then
    s = trim(s:sub(1, pos - 1))
  end
  -- Stored macros may use "§§Note: ..." (double); only §§§ was stripped above. Leftover §§ becomes
  -- newlines and WoW treats non-/ lines as /say — strip note tails before § → newline.
  local doubleUtf8 = "\194\167\194\167"
  pos = s:find(doubleUtf8, 1, true)
  if pos then
    s = trim(s:sub(1, pos - 1))
  end
  local doubleLatin = "\167\167"
  pos = s:find(doubleLatin, 1, true)
  if pos then
    s = trim(s:sub(1, pos - 1))
  end
  local parts = {}
  local i = 1
  local n = #s
  while i <= n do
    local b = s:byte(i)
    if b >= 0xC2 and b <= 0xDF and i < n then
      local b2 = s:byte(i + 1)
      if b2 >= 0x80 and b2 <= 0xBF then
        if b == 0xC2 and b2 == 0xA7 then
          parts[#parts + 1] = "\n"
        else
          parts[#parts + 1] = s:sub(i, i + 1)
        end
        i = i + 2
      else
        parts[#parts + 1] = string.char(b)
        i = i + 1
      end
    elseif b == 0xA7 then
      parts[#parts + 1] = "\n"
      i = i + 1
    elseif b >= 0xE0 and b <= 0xEF and i + 2 <= n then
      parts[#parts + 1] = s:sub(i, i + 2)
      i = i + 3
    elseif b >= 0xF0 and b <= 0xF7 and i + 3 <= n then
      parts[#parts + 1] = s:sub(i, i + 3)
      i = i + 4
    else
      parts[#parts + 1] = string.char(b)
      i = i + 1
    end
  end
  s = table.concat(parts)
  s = s:gsub("\r\n", "\n"):gsub("\r", "\n")
  return trim(s)
end

--- Non-empty lines must start with / or # (showtooltip); otherwise WoW treats them as /say.
--- Also flags known profile placeholder strings mistaken for macros.
--- @return string|nil short reason, or nil if OK
local function macroPlainSayIssue(rawMacroText)
  local mac = normalizeStoredMacroText(rawMacroText or "")
  if mac == "" then
    return nil
  end
  if mac:find("Same hotkey must be assigned", 1, true) then
    return "placeholder (not a /slash macro)"
  end
  for line in (mac .. "\n"):gmatch("([^\n]*)\n") do
    local t = trim(line)
    if t ~= "" then
      local c = t:sub(1, 1)
      if c ~= "/" and c ~= "#" then
        if #t > 64 then
          t = t:sub(1, 61) .. "..."
        end
        return "plain line → /say: " .. t
      end
    end
  end
  return nil
end

-- =========================
-- Icon resolution + drag-to-bar
-- =========================

local QUESTION_MARK_ICON = 134400
local TEMP_MACRO_NAME = "CBIconTmp"

local function parseFirstSpellOrItemFromMacro(macroText)
  local mac = normalizeStoredMacroText(macroText or "")
  if mac == "" then return nil, nil end
  local candidate
  for line in mac:gmatch("[^\n]+") do
    local t = trim(line)
    if not candidate then
      local afterShow = t:match("^#showtooltip%s+(.+)$") or t:match("^#show%s+(.+)$")
      if afterShow and trim(afterShow) ~= "" then
        candidate = trim(afterShow)
        break
      end
    end
    if not candidate then
      local afterCmd = t:match("^/cast%s+(.+)$") or t:match("^/use%s+(.+)$")
      if afterCmd then
        candidate = trim(afterCmd)
        break
      end
    end
  end
  if not candidate then return nil, nil end
  local stripped = candidate
  if stripped:find("%[") then
    for segment in stripped:gmatch("[^;]+") do
      local seg = trim(segment)
      local name = seg:match("%]%s*(.+)$")
      if name and trim(name) ~= "" then
        stripped = trim(name)
        break
      elseif not seg:find("%[") and seg ~= "" then
        stripped = seg
        break
      end
    end
  else
    local first = stripped:match("^([^;]+)")
    if first then stripped = trim(first) end
  end
  local itemId = stripped:match("^item:(%d+)")
  if itemId then return "item", tonumber(itemId) end
  return "name", stripped
end

local function getIconViaTempMacro(macroText)
  if InCombatLockdown() then return nil end
  if MacroFrame and MacroFrame:IsShown() then return nil end
  local existing = GetMacroInfo(TEMP_MACRO_NAME)
  if not existing then
    local cnt1, cnt2 = GetNumMacros()
    if cnt1 >= MAX_ACCOUNT_MACROS then
      if cnt2 >= MAX_CHARACTER_MACROS then return nil end
      CreateMacro(TEMP_MACRO_NAME, QUESTION_MARK_ICON, macroText, true)
    else
      CreateMacro(TEMP_MACRO_NAME, QUESTION_MARK_ICON, macroText, false)
    end
  else
    EditMacro(TEMP_MACRO_NAME, nil, nil, macroText)
  end
  local _, tex = GetMacroInfo(TEMP_MACRO_NAME)
  DeleteMacro(TEMP_MACRO_NAME)
  return tex
end

local function resolveIconForBind(bind)
  if not bind then return QUESTION_MARK_ICON end
  if bind.textureId and bind.textureId ~= 0 then return bind.textureId end
  local mac = normalizeStoredMacroText(bind.macroText or "")
  if mac == "" then return QUESTION_MARK_ICON end
  if CB.iconCache[mac] then return CB.iconCache[mac] end
  local kind, value = parseFirstSpellOrItemFromMacro(bind.macroText)
  local icon
  if kind == "name" and value then
    local info = C_Spell and C_Spell.GetSpellInfo and C_Spell.GetSpellInfo(value)
    if info and info.iconID then icon = info.iconID end
    if not icon then
      local ok, r1, r2, r3, r4, r5, r6, r7, r8, r9, r10 = pcall(GetItemInfo, value)
      if ok and r10 then icon = r10 end
    end
  elseif kind == "item" and value then
    if C_Item and C_Item.GetItemIconByID then
      icon = C_Item.GetItemIconByID(value)
    end
  end
  if not icon then
    icon = getIconViaTempMacro(mac)
  end
  icon = icon or QUESTION_MARK_ICON
  CB.iconCache[mac] = icon
  return icon
end

local function parseBindForPickup(bind)
  if not bind then return nil end
  local mac = normalizeStoredMacroText(bind.macroText or "")
  if mac == "" then return nil end
  local kind, value = parseFirstSpellOrItemFromMacro(bind.macroText)
  if kind == "name" and value then
    local info = C_Spell and C_Spell.GetSpellInfo and C_Spell.GetSpellInfo(value)
    if info and info.spellID then
      return { type = "spell", spellID = info.spellID, name = info.name }
    end
    local ok, itemName, itemLink = pcall(GetItemInfo, value)
    if ok and itemLink then
      local itemID = tonumber(itemLink:match("item:(%d+)"))
      if itemID then return { type = "item", itemID = itemID, name = itemName } end
    end
  elseif kind == "item" and value then
    return { type = "item", itemID = value, name = tostring(value) }
  end
  return { type = "macro", macroText = mac, name = bind.actionName or "CrimsonBind" }
end

local function placeBindOnCursor(bind)
  if InCombatLockdown() then
    print("|cff00ccffCrimsonBind|r Exit combat to place a bind on the cursor.")
    return
  end
  local info = parseBindForPickup(bind)
  if not info then
    print("|cff00ccffCrimsonBind|r No actionable content in this bind.")
    return
  end
  if info.type == "spell" then
    if C_Spell and C_Spell.PickupSpell then
      C_Spell.PickupSpell(info.spellID)
    else
      PickupSpell(info.spellID)
    end
    print("|cff00ccffCrimsonBind|r Picked up spell: " .. tostring(info.name) .. " — drop on your action bar.")
  elseif info.type == "item" then
    if C_Item and C_Item.PickupItem then
      C_Item.PickupItem(info.itemID)
    else
      PickupItem(info.itemID)
    end
    print("|cff00ccffCrimsonBind|r Picked up item: " .. tostring(info.name) .. " — drop on your action bar.")
  elseif info.type == "macro" then
    local cnt1, cnt2 = GetNumMacros()
    local macName = "CB:" .. (bind.actionName or "Macro"):sub(1, 12)
    local existing = GetMacroInfo(macName)
    if existing then
      EditMacro(macName, nil, nil, info.macroText)
    else
      if cnt1 >= MAX_ACCOUNT_MACROS then
        if cnt2 >= MAX_CHARACTER_MACROS then
          print("|cff00ccffCrimsonBind|r Macro slots full (" .. cnt1 .. " account + " .. cnt2 .. " character). Free a slot first.")
          return
        end
        CreateMacro(macName, resolveIconForBind(bind), info.macroText, true)
      else
        CreateMacro(macName, resolveIconForBind(bind), info.macroText, false)
      end
    end
    PickupMacro(macName)
    print("|cff00ccffCrimsonBind|r Created WoW macro '" .. macName .. "' (uses 1 slot). Drop it on your action bar.")
  end
end

-- =========================
-- Heatmap data
-- =========================

local HEATMAP_REGIONS = {
  { name = "Number Row", keys = {"6","7","8","9","0","MINUS","EQUALS"}, labels = {"6","7","8","9","0","-","="} },
  { name = "Top Row", keys = {"T","Y","U","I","O","P","LBRACKET","RBRACKET","BACKSLASH"}, labels = {"T","Y","U","I","O","P","[","]","\\"} },
  { name = "Home Row", keys = {"G","H","J","K","L","SEMICOLON","APOSTROPHE"}, labels = {"G","H","J","K","L",";","'"} },
  { name = "Bottom Row", keys = {"Z","X","C","V","B","N","M","COMMA","PERIOD","SLASH"}, labels = {"Z","X","C","V","B","N","M",",",".","/"}},
  { name = "F-Keys", keys = {"F1","F2","F3","F4","F5","F6","F7","F8","F9","F10","F11","F12"}, labels = {"F1","F2","F3","F4","F5","F6","F7","F8","F9","F10","F11","F12"} },
  { name = "Navigation", keys = {"TAB","INSERT","HOME","PAGEUP","PAGEDOWN","END","UP","DOWN","LEFT","RIGHT"}, labels = {"Tab","Ins","Hm","PU","PD","End","Up","Dn","Lt","Rt"} },
}

local HEATMAP_MODS = {
  { prefix = "", label = "None" },
  { prefix = "CTRL-", label = "C-" },
  { prefix = "ALT-", label = "A-" },
  { prefix = "SHIFT-", label = "S-" },
  { prefix = "CTRL-ALT-", label = "CA-" },
  { prefix = "CTRL-SHIFT-", label = "CS-" },
  { prefix = "CTRL-ALT-SHIFT-", label = "CAS-" },
}

local HEATMAP_COLORS = {
  available = {0.15, 0.15, 0.18, 0.9},
  general = {0.2, 0.4, 0.8, 0.9},
  spec = {0.2, 0.7, 0.3, 0.9},
  custom = {0.85, 0.7, 0.2, 0.9},
  conflict = {0.8, 0.2, 0.2, 0.9},
  excluded = {0.3, 0.15, 0.15, 0.7},
  excluded_base = {0.08, 0.08, 0.1, 0.6},
}

local HEATMAP_EXCLUDED_KEYS = {
  ["CTRL-C"] = true, ["CTRL-V"] = true, ["CTRL-MINUS"] = true,
  ["ALT-TAB"] = true, ["CTRL-ALT-TAB"] = true, ["ALT-F4"] = true,
  ["ALT-ESCAPE"] = true, ["CTRL-ESCAPE"] = true, ["ALT-SPACE"] = true,
  ["CTRL-SHIFT-ESCAPE"] = true, ["ALT-Z"] = true, ["ALT-F1"] = true,
  ["ALT-F9"] = true, ["ALT-F10"] = true,
}

local HEATMAP_EXCLUDED_BASES = {
  Q = true, W = true, E = true, A = true, S = true, D = true,
  R = true, F = true,
  ["1"] = true, ["2"] = true, ["3"] = true, ["4"] = true, ["5"] = true,
}

local function splitKey(key)
  key = trim(key)
  if key == "" then return {}, "" end
  local mods = {}
  local rest = key
  while true do
    local a, b, mod, tail = rest:find("^(%a+)%-(.*)$")
    if not mod then break end
    local u = mod:upper()
    if u == "CTRL" or u == "ALT" or u == "SHIFT" then
      table.insert(mods, u)
      rest = tail
    else
      break
    end
  end
  table.sort(mods)
  -- WoW binding name is MINUS/EQUALS; a lone "-" after mod prefixes was parsed as base "-" → normalizeKey produced CTRL--.
  if rest == "-" then
    rest = "MINUS"
  elseif rest == "=" then
    rest = "EQUALS"
  end
  return mods, rest
end

local function normalizeKey(key)
  local mods, base = splitKey(key)
  if base == "" then return "" end
  if #mods == 0 then return base:upper() end
  return table.concat(mods, "-") .. "-" .. base:upper()
end

--- Human/list display: same canonical form as SetBinding (fixes CTRL-- vs CTRL-MINUS).
--- Punctuation: SEMICOLON shows as ";" or "shift-:" (WoW still stores SHIFT-SEMICOLON).
local function keyStringForDisplay(raw)
  local k = trim(raw or "")
  if k == "" then
    return ""
  end
  local norm = normalizeKey(k)
  local mods, base = splitKey(norm)
  if base == "" then
    return ""
  end
  local bu = strupper(base)
  if bu == "SEMICOLON" then
    local hasShift = false
    for _, m in ipairs(mods) do
      if m == "SHIFT" then
        hasShift = true
        break
      end
    end
    base = hasShift and ":" or ";"
  elseif bu == "LBRACKET" then
    base = "["
  elseif bu == "RBRACKET" then
    base = "]"
  elseif bu == "COMMA" then
    base = ","
  elseif bu == "BACKSLASH" then
    base = "\\"
  elseif bu == "PERIOD" then
    base = "."
  elseif bu == "SLASH" then
    base = "/"
  elseif bu == "APOSTROPHE" then
    base = "'"
  elseif bu == "GRAVE" then
    base = "`"
  else
    base = bu
  end
  if #mods == 0 then
    return base
  end
  return table.concat(mods, "-") .. "-" .. base
end

--- Numpad nav labels (e.g. from profiles) → WoW NUMPAD0–9 (same physical keys with Num Lock on).
local NUMPAD_NAV_LABEL_TO_NUM = {
  NUMPADHOME = "NUMPAD7",
  NUMPADEND = "NUMPAD1",
  NUMPADUP = "NUMPAD8",
  NUMPADDOWN = "NUMPAD2",
  NUMPADLEFT = "NUMPAD4",
  NUMPADRIGHT = "NUMPAD6",
  NUMPADPGUP = "NUMPAD9",
  NUMPADPGDN = "NUMPAD3",
  NUMPADPAGEUP = "NUMPAD9",
  NUMPADPAGEDOWN = "NUMPAD3",
  NUMPADINS = "NUMPAD0",
  NUMPADINSERT = "NUMPAD0",
}

--- WoW / driver tokens for numpad operators → SetBinding names.
local BINDING_NUMPAD_OPERATOR_CANON = {
  NUMPADMULTIPLY = "NUMPADMULTIPLY",
  NUMPADMINUS = "NUMPADMINUS",
  NUMPADPLUS = "NUMPADPLUS",
  NUMPADDIVIDE = "NUMPADDIVIDE",
  MULTIPLY = "NUMPADMULTIPLY",
  DIVIDE = "NUMPADDIVIDE",
  -- SUBTRACT / ADD: WoW may use these for the main - and + keys too; never force to numpad without IsKeyDown (see buildBindingStringFromKeyPress).
  NUMPADSUBTRACT = "NUMPADMINUS",
  NUMPADDIV = "NUMPADDIVIDE",
}

--- Profile / UI may use "NUM PAD 9" or "NUM-PAD-9"; SetBinding / GetBindingAction expect NUMPAD9.
local function coerceNumpadKeyName(key)
  key = trim(key or "")
  if key == "" then return "" end
  local mods, base = splitKey(key)
  if base == "" then return key end
  local u = base:upper()
  local coerced = u:gsub("^NUM[%-_ ]*PAD[%-_ ]+(%d)$", "NUMPAD%1")
  if coerced ~= u then
    base = coerced
  else
    base = u
  end
  local navNum = NUMPAD_NAV_LABEL_TO_NUM[base]
  if navNum then
    base = navNum
  end
  local opCanon = BINDING_NUMPAD_OPERATOR_CANON[base]
  if opCanon then
    base = opCanon
  end
  if #mods == 0 then return base end
  table.sort(mods)
  return table.concat(mods, "-") .. "-" .. base
end

--- WoW SetBinding expects literal , \ . / ; ' ` [ ] for these keys; spelled tokens (COMMA, BACKSLASH, …) from SavedVariables/sync often do not fire.
--- Omit MINUS/EQUALS: this addon keeps MINUS/EQUALS tokens (CTRL-MINUS policy, splitKey normalization).
local BINDING_NAME_TO_SETBINDING_LITERAL = {
  LBRACKET = "[",
  RBRACKET = "]",
  COMMA = ",",
  BACKSLASH = "\\",
  PERIOD = ".",
  SLASH = "/",
  SEMICOLON = ";",
  APOSTROPHE = "'",
  GRAVE = "`",
}

local function coercePunctuationToSetBindingLiteral(key)
  key = trim(key or "")
  if key == "" then
    return ""
  end
  local mods, base = splitKey(key)
  if base == "" then
    return key
  end
  local bu = strupper(base)
  local lit = BINDING_NAME_TO_SETBINDING_LITERAL[bu]
  if lit then
    base = lit
  end
  if #mods == 0 then
    return base
  end
  table.sort(mods)
  return table.concat(mods, "-") .. "-" .. base
end

--- UnitClass returns localized name first, English token second (e.g. "Druid", "DRUID"). Config rows use localized "Class - Spec".
local function localizedPlayerClass()
  local locName, token = UnitClass("player")
  if locName and locName ~= "" then
    return locName
  end
  return token
end

local function bindSectionsEquivalent(a, b)
  if not a or not b then return false end
  if a == b then return true end
  return a:lower() == b:lower()
end

--- Exact section string for config.ini / CSV (e.g. "Druid - Guardian"), or nil if APIs have not returned a spec yet.
--- Uses active dual-spec group. Never returns bare class name (that never matches stored sections).
local function resolveCurrentBindSection()
  local classLoc = localizedPlayerClass()
  if not classLoc or classLoc == "" then
    return nil
  end
  local group = 1
  if GetActiveSpecGroup then
    group = GetActiveSpecGroup(false, false) or 1
  end
  local specIdx
  if GetSpecialization then
    specIdx = GetSpecialization(false, false, group)
    if not specIdx then
      specIdx = GetSpecialization()
    end
  end
  local specName
  if specIdx and GetSpecializationInfo then
    local _, sn = GetSpecializationInfo(specIdx)
    if sn and trim(sn) ~= "" then
      specName = trim(sn)
    end
  end
  if not specName and GetPrimaryTalentTree and GetTalentTreeInfo then
    local tab = GetPrimaryTalentTree(false, false, group)
    if tab then
      local ok, a, b = pcall(GetTalentTreeInfo, tab)
      if ok then
        if type(b) == "string" and trim(b) ~= "" then
          specName = trim(b)
        elseif type(a) == "string" and trim(a) ~= "" then
          specName = trim(a)
        end
      end
    end
  end
  if specName then
    return classLoc .. " - " .. specName
  end
  return nil
end

--- For chat/UI; when spec is unknown, say so (never return bare "Druid" — that mismatches bind sections).
local function getCurrentSection()
  local r = resolveCurrentBindSection()
  if r then
    return r
  end
  return "General (spec not detected)"
end

local function classBindSectionPrefix()
  local classLoc = localizedPlayerClass()
  if not classLoc or classLoc == "" then
    return nil
  end
  return classLoc .. " - "
end

--- Binds that apply in-game: [General], current class/spec section, and [CUSTOM] (always-on user binds).
local function isBindRowActiveForCurrentSpec(b)
  if not b then return false end
  local sec = b.section or ""
  if sec == "General" then return true end
  if strupper(sec) == "CUSTOM" then return true end
  local cur = resolveCurrentBindSection()
  if cur then
    return bindSectionsEquivalent(sec, cur)
  end
  return false
end

local function initSavedVars()
  CrimsonBindVars = CrimsonBindVars or {}
  CrimsonBindVars.version = CrimsonBindVars.version or 1
  CrimsonBindVars.binds = CrimsonBindVars.binds or {}
  CrimsonBindVars.abilityManifest = CrimsonBindVars.abilityManifest or {}
  CrimsonBindVars.pendingEdits = CrimsonBindVars.pendingEdits or {}
  if CrimsonBindVars.testMode == nil then
    CrimsonBindVars.testMode = false
  end
end

local function loadTestModeFromSavedVars()
  CB.testMode = CrimsonBindVars.testMode and true or false
end

--- Escape for Lua string literals inside /run ... "..." in a macro line.
local function escapeRunArg(s)
  s = trim(s or "")
  s = s:gsub("\\", "\\\\")
  s = s:gsub('"', '\\"')
  s = s:gsub("\r\n", " "):gsub("\n", " "):gsub("\r", " ")
  return s
end

-- Macro body max length (Blizzard UI limit).
local MACRO_MAX = 255

--- config.ini-style compact modifiers: ^ ! + prefix without dashes (e.g. !NUMPADHOME → ALT-NUMPADHOME).
local function expandCompactModKeyPrefix(key)
  key = trim(key or "")
  if key == "" or key:find("%-") then
    return key
  end
  if key:match("^%^?[!%+]*vk%x+_") or key:match("^%^?[!%+]*sc%x+_") then
    return key
  end
  local i = 1
  local mods = {}
  local len = #key
  while i <= len do
    local c = key:sub(i, i)
    if c == "^" then
      table.insert(mods, "CTRL")
    elseif c == "!" then
      table.insert(mods, "ALT")
    elseif c == "+" then
      table.insert(mods, "SHIFT")
    else
      break
    end
    i = i + 1
  end
  if i == 1 then
    return key
  end
  local base = trim(key:sub(i))
  if base == "" then
    return key
  end
  if base == "=" then
    base = "EQUALS"
  elseif base == "-" then
    base = "MINUS"
  end
  table.sort(mods)
  return table.concat(mods, "-") .. "-" .. base
end

local function bindingTokenFromKeyString(raw)
  local k0 = expandCompactModKeyPrefix(raw)
  local k = normalizeKey(k0 or "")
  if k == "" then
    k = trim(k0 or "")
  end
  return coercePunctuationToSetBindingLiteral(coerceNumpadKeyName(k))
end

--- After canonicalization, true if binding uses any numpad physical key (0–9, operators, nav synonyms → NUMPADn).
local function bindKeyUsesNumpadBase(raw)
  local k = bindingTokenFromKeyString(raw or "")
  if k == "" then
    return false
  end
  local _, base = splitKey(k)
  if base == "" then
    return false
  end
  local u = strupper(base)
  return u:match("^NUMPAD") ~= nil
end

--- Bare MINUS or CTRL-MINUS: unreliable in WoW (stale dual binds, lookups vs fire mismatch). Other combos (e.g. CTRL-SHIFT-MINUS) are left enabled.
local function bindKeyUsesBlockedCtrlMinusOrBare(raw)
  local k = bindingTokenFromKeyString(raw or "")
  if k == "" then
    return false
  end
  local mods, base = splitKey(k)
  if strupper(base or "") ~= "MINUS" then
    return false
  end
  if #mods == 0 then
    return true
  end
  return #mods == 1 and mods[1] == "CTRL"
end

local function getBindsToApply()
  local cur = resolveCurrentBindSection()
  local byKey = {}
  local order = {}
  local function addBind(b)
    local macEff = normalizeStoredMacroText(b and b.macroText or "")
    if not b or not b.key or b.key == "" or macEff == "" then
      return
    end
    if bindKeyUsesNumpadBase(b.key) then
      return
    end
    if bindKeyUsesBlockedCtrlMinusOrBare(b.key) then
      return
    end
    local nk = bindingTokenFromKeyString(b.key)
    if nk == "" then return end
    if not byKey[nk] then
      byKey[nk] = b
      table.insert(order, b)
    else
      byKey[nk] = b
      for i, ob in ipairs(order) do
        if bindingTokenFromKeyString(ob.key) == nk then
          order[i] = b
          break
        end
      end
    end
  end
  for _, b in ipairs(CrimsonBindVars.binds) do
    if b.section == "General" then
      addBind(b)
    end
  end
  if cur then
    for _, b in ipairs(CrimsonBindVars.binds) do
      if bindSectionsEquivalent(b.section, cur) then
        addBind(b)
      end
    end
  end
  -- CUSTOM binds applied last: user-pinned keys always win on any collision with spec/General
  for _, b in ipairs(CrimsonBindVars.binds) do
    if strupper(b.section or "") == "CUSTOM" then
      addBind(b)
    end
  end
  return order
end

--- Canonical key string for SetBinding / GetBindingAction (same rules as applyAll).
local function bindKeyFromBind(b)
  if not b then return "" end
  return bindingTokenFromKeyString(b.key or "")
end

--- Single-line macro that prints action name and configured key (no spell/item).
local function buildTestModeMacroText(b, slotIndex)
  local act = escapeRunArg(b.actionName or "?")
  local ky = escapeRunArg(bindKeyFromBind(b))
  local line = '/run print("\\124cff00ccff[CrimsonBind Test]\\124r", "' .. act .. '", "' .. ky .. '")'
  if #line <= MACRO_MAX then
    return line
  end
  act = escapeRunArg((b.actionName or "?"):sub(1, 80))
  line = '/run print("\\124cff00ccff[CrimsonBind Test]\\124r", "' .. act .. '", "' .. ky .. '")'
  if #line <= MACRO_MAX then
    return line
  end
  line = '/run print("\\124cff00ccffCB\\124r",' .. tostring(slotIndex) .. ',"' .. ky .. '")'
  if #line <= MACRO_MAX then
    return line
  end
  ky = escapeRunArg((bindKeyFromBind(b)):sub(1, 40))
  line = '/run print("CB",' .. tostring(slotIndex) .. ',"' .. ky .. '")'
  if #line > MACRO_MAX then
    line = '/run print("CB",' .. tostring(slotIndex) .. ')'
  end
  return line
end

local function syncTestModeUi()
  if CB.testModeCheck then
    CB.testModeCheck:SetChecked(CB.testMode)
  end
  if CB.testModeBanner then
    if CB.testMode then
      CB.testModeBanner:Show()
    else
      CB.testModeBanner:Hide()
    end
  end
end

local function isOurCrimsonBindClickAction(action)
  if not action or action == "" then return false end
  return action:match("^CLICK CrimsonBindMacro:")
end

--- For GetBindingAction only: WoW may index punctuation keys under the symbol form (e.g. CTRL--)
--- while SetBinding accepts CTRL-MINUS — same physical key, different lookup string.
local BINDING_BASE_GET_ALIASES = {
  MINUS = { "-" },
  EQUALS = { "=" },
  SEMICOLON = { ";" },
  PERIOD = { "." },
  SLASH = { "/" },
  BACKSLASH = { "\\" },
  LBRACKET = { "[" },
  RBRACKET = { "]" },
  COMMA = { "," },
  APOSTROPHE = { "'" },
  GRAVE = { "`" },
  --- Reverse: after canonical uses literals, GetBindingAction may still index by token name.
  [","] = { "COMMA" },
  ["\\"] = { "BACKSLASH" },
  ["."] = { "PERIOD" },
  ["/"] = { "SLASH" },
  [";"] = { "SEMICOLON" },
  ["'"] = { "APOSTROPHE" },
  ["`"] = { "GRAVE" },
  ["["] = { "LBRACKET" },
  ["]"] = { "RBRACKET" },
}

local function composeModsAndBase(mods, baseToken)
  local s = tostring(baseToken or "")
  local bp
  if #s == 1 and not s:match("%a") then
    bp = s
  else
    bp = strupper(s)
  end
  if not mods or #mods == 0 then
    return bp
  end
  return table.concat(mods, "-") .. "-" .. bp
end

local function bindingKeyLookupVariants(canonicalKey)
  local variants = {}
  local seen = {}
  local function add(v)
    if v and v ~= "" and not seen[v] then
      seen[v] = true
      variants[#variants + 1] = v
    end
  end
  add(canonicalKey)
  local mods, base = splitKey(canonicalKey)
  if base ~= "" then
    local alts = BINDING_BASE_GET_ALIASES[strupper(base)]
    if alts then
      for _, ab in ipairs(alts) do
        add(composeModsAndBase(mods, ab))
      end
    end
  end
  return variants
end

--- Binding lookup for keys CrimsonBind sets via SetBinding+SaveBindings. Retail keeps those in the
--- "override" layer; GetBindingAction(key) with one argument often returns "" (false negatives).
local function getBindingActionForKey(key)
  if type(GetBindingAction) ~= "function" or not key or key == "" then
    return ""
  end
  for _, vk in ipairs(bindingKeyLookupVariants(key)) do
    local ok, act = pcall(GetBindingAction, vk, true)
    if ok and type(act) == "string" and act ~= "" then
      return act
    end
    ok, act = pcall(GetBindingAction, vk, false)
    if ok and type(act) == "string" and act ~= "" then
      return act
    end
    ok, act = pcall(GetBindingAction, vk)
    if ok and type(act) == "string" and act ~= "" then
      return act
    end
  end
  return ""
end

-- =========================
-- Apply / binding helpers
-- =========================

local function countBindingReplacements(list)
  local n = 0
  for _, b in ipairs(list or {}) do
    local bk = bindKeyFromBind(b)
    if bk ~= "" then
      local act = getBindingActionForKey(bk)
      if act and act ~= "" and not isOurCrimsonBindClickAction(act) then
        n = n + 1
      end
    end
  end
  return n
end

--- Keys in the current apply set that WoW still maps to something other than our CLICK (Blizzard UI, BindPad, etc.).
local function getWoWKeyCompetitorsForPlannedApply()
  local ext = {}
  local list = CB.cachedApplyList or getBindsToApply()
  local seen = {}
  for _, b in ipairs(list) do
    local bk = bindKeyFromBind(b)
    if bk ~= "" and not seen[bk] then
      seen[bk] = true
      local act = getBindingActionForKey(bk)
      if act and act ~= "" and not isOurCrimsonBindClickAction(act) then
        ext[bk] = act
      end
    end
  end
  return ext
end

local function competingBindingLines()
  local lines = {}
  local ext = getWoWKeyCompetitorsForPlannedApply()
  for bk, act in pairs(ext) do
    table.insert(lines, "WoW key conflict: " .. bk .. " -> " .. act .. " (Apply overwrites)")
  end
  table.sort(lines)
  return lines
end

--- Disabling BindPad does not remove its CLICK bindings from the client; they stay until SetBinding(key,nil).
--- Pass 1: GetBinding(i) rows whose command mentions BindPad (often omits CLICK binds on many clients).
--- Pass 2: Every key on any CrimsonBind row whose GetBindingAction still says BindPad (matches UI tooltip).
local function clearBindPadClickBindingsFromWoW()
  if type(SetBinding) ~= "function" then
    return 0
  end
  local toClear = {}
  local function mark(k)
    if type(k) == "string" and k ~= "" then
      toClear[k] = true
    end
  end
  if type(GetNumBindings) == "function" and type(GetBinding) == "function" then
    local n = GetNumBindings()
    if n and n >= 1 then
      for i = 1, n do
        local ok, t = pcall(function()
          return { GetBinding(i) }
        end)
        if ok and type(t) == "table" and type(t[1]) == "string" then
          local cmd = t[1]
          if string.lower(cmd):find("bindpad", 1, true) then
            for j = 3, #t do
              mark(t[j])
            end
          end
        end
      end
    end
  end
  if CrimsonBindVars and type(CrimsonBindVars.binds) == "table" then
    local seenBk = {}
    for _, b in ipairs(CrimsonBindVars.binds) do
      local bk = bindKeyFromBind(b)
      if bk ~= "" and not seenBk[bk] then
        seenBk[bk] = true
        local act = getBindingActionForKey(bk)
        if act ~= "" and string.lower(act):find("bindpad", 1, true) then
          for _, vk in ipairs(bindingKeyLookupVariants(bk)) do
            mark(vk)
          end
        end
      end
    end
  end
  local c = 0
  for kk in pairs(toClear) do
    if pcall(SetBinding, kk, nil) then
      c = c + 1
    end
  end
  if c > 0 and SaveBindings and GetCurrentBindingSet then
    pcall(SaveBindings, GetCurrentBindingSet())
  end
  return c
end

--- WoW allows multiple keys per command via SetBinding. Stale keys (e.g. CTRL-MINUS) can keep
--- pointing at CLICK CrimsonBindMacro:CBn after list/order changes while we only nil prevBoundKeys.
--- Remove every key WoW has on each CB slot before we SetBinding again.
local function clearWoWKeysBoundToCrimsonBindClicks(maxSlots)
  if not maxSlots or maxSlots < 1 then
    return
  end
  if type(GetBindingKey) ~= "function" or type(SetBinding) ~= "function" then
    return
  end
  local cleared = {}
  for i = 1, maxSlots do
    local cmd = "CLICK CrimsonBindMacro:CB" .. i
    local ok, keyList = pcall(function()
      return { GetBindingKey(cmd) }
    end)
    if ok and type(keyList) == "table" then
      for _, k in ipairs(keyList) do
        if type(k) == "string" and k ~= "" and not cleared[k] then
          cleared[k] = true
          pcall(SetBinding, k, nil)
        end
      end
    end
  end
end

--- @param maxMacroSlots number clear *macrotext-CB1..CBn (n = max of previous apply count and this pass; avoids wiping 600 every time).
local function clearOurBindings(maxMacroSlots)
  for k in pairs(CB.prevBoundKeys) do
    SetBinding(k, nil)
  end
  wipe(CB.prevBoundKeys)
  if CB.macroBtn and maxMacroSlots and maxMacroSlots > 0 then
    for i = 1, maxMacroSlots do
      CB.macroBtn:SetAttribute("*macrotext-CB" .. i, nil)
    end
  end
end

--- Match BindPad: CLICK bindings only fire if the secure button registers for the same press phase as ActionButtonUseKeyDown.
local function syncMacroBtnRegisterForClicks()
  if not CB.macroBtn then return end
  local keyDown
  if GetCVarBool then
    keyDown = GetCVarBool("ActionButtonUseKeyDown")
  else
    keyDown = GetCVar("ActionButtonUseKeyDown") == "1"
  end
  if keyDown then
    CB.macroBtn:RegisterForClicks("AnyDown")
  else
    CB.macroBtn:RegisterForClicks("AnyUp")
  end
end

function CB.applyAll()
  if InCombatLockdown() then
    CB.pendingApply = true
    return false
  end
  if not CB.macroBtn then return false end
  syncMacroBtnRegisterForClicks()
  local list = getBindsToApply()
  local clearSlots = math.max(CB.lastSlotCount or 0, #list)
  local replacedOther = 0
  for _, b in ipairs(list) do
    local bk = bindKeyFromBind(b)
    if bk ~= "" then
      local act = getBindingActionForKey(bk)
      if act and act ~= "" and not isOurCrimsonBindClickAction(act) then
        replacedOther = replacedOther + 1
      end
    end
  end
  clearOurBindings(clearSlots)
  clearWoWKeysBoundToCrimsonBindClicks(clearSlots)
  CB.macroBtn:SetAttribute("*type*", "macro")
  local bindErrCount = 0
  for i, b in ipairs(list) do
    local slot = "CB" .. i
    local body = (CB.testMode and buildTestModeMacroText(b, i)) or normalizeStoredMacroText(b.macroText or "")
    local okAttr, errAttr = pcall(function()
      CB.macroBtn:SetAttribute("*macrotext-" .. slot, body)
    end)
    if not okAttr then
      bindErrCount = bindErrCount + 1
      if CB.debug then
        print("|cff00ccffCrimsonBind|r SetAttribute error slot " .. i .. ": " .. tostring(errAttr))
      end
    end
    local bindKey = bindKeyFromBind(b)
    if bindKey ~= "" then
      local clickCmd = "CLICK CrimsonBindMacro:" .. slot
      --- WoW may keep an alternate token for the same physical key (e.g. CTRL-. vs CTRL-PERIOD) on another command.
      for _, vk in ipairs(bindingKeyLookupVariants(bindKey)) do
        pcall(SetBinding, vk, nil)
      end
      local okBind, errBind = pcall(SetBinding, bindKey, clickCmd)
      if okBind then
        CB.prevBoundKeys[bindKey] = true
      else
        bindErrCount = bindErrCount + 1
        if CB.debug then
          print("|cff00ccffCrimsonBind|r SetBinding error slot " .. i .. " key " .. tostring(bindKey) .. ": " .. tostring(errBind))
        end
      end
    end
  end
  SaveBindings(GetCurrentBindingSet())
  CB.lastSlotCount = #list
  CB.pendingApply = false
  if replacedOther > 0 then
    print("|cff00ccffCrimsonBind|r Apply replaced WoW / other-addon binding(s) on |cffffcc00" .. replacedOther .. "|r key(s). Run |cffaaaaaa/cb validate|r for details.")
  end
  if CB.debug then
    print("|cff00ccffCrimsonBind|r applyAll done: slots=" .. #list .. " vsWoW=" .. replacedOther .. " bindErrors=" .. bindErrCount)
  end
  return true
end

local function isTargetMember110(name)
  if not name then return false end
  local n = name:match("^Target Member(%d+)$")
  if not n then return false end
  local d = tonumber(n)
  return d and d >= 1 and d <= 10
end

local function validationSummary()
  local cur = resolveCurrentBindSection()
  local bySec = {}
  for _, b in ipairs(CrimsonBindVars.binds) do
    local sec = b.section or ""
    if sec == "General" or (cur and bindSectionsEquivalent(sec, cur)) or strupper(sec) == "CUSTOM" then
      if not bySec[sec] then bySec[sec] = {} end
      local k = bindingTokenFromKeyString(b.key)
      if k ~= "" then
        if not bySec[sec][k] then bySec[sec][k] = {} end
        table.insert(bySec[sec][k], b.actionName or "?")
      end
    end
  end
  local lines = {}
  for sec, keys in pairs(bySec) do
    for k, names in pairs(keys) do
      if #names > 1 then
        table.insert(lines, sec .. ": duplicate " .. k .. " -> " .. table.concat(names, ", "))
      end
    end
  end
  for _, b in ipairs(CrimsonBindVars.binds) do
    if isBindRowActiveForCurrentSpec(b) and isTargetMember110(b.actionName) and b.key and b.key ~= "" then
      local low = b.key:lower()
      if low:find("ctrl") or low:find("alt") then
        table.insert(lines, (b.section or "") .. ": Target Member 1-10 should not use CTRL/ALT: " .. tostring(b.actionName))
      end
    end
  end
  for _, b in ipairs(CrimsonBindVars.binds) do
    if isBindRowActiveForCurrentSpec(b) then
      local nk = bindingTokenFromKeyString(b.key)
      local mac = normalizeStoredMacroText(b.macroText or "")
      if nk ~= "" and mac == "" then
        table.insert(lines, (b.section or "") .. ": empty macro with key " .. nk .. ": " .. tostring(b.actionName))
      end
      if #mac > MACRO_MAX then
        table.insert(lines, (b.section or "") .. ": macro > " .. MACRO_MAX .. " chars: " .. tostring(b.actionName))
      end
      local sayIss = macroPlainSayIssue(b.macroText or "")
      if sayIss and mac ~= "" then
        table.insert(
          lines,
          (b.section or "") .. ": " .. sayIss .. " — " .. tostring(b.actionName) .. (nk ~= "" and "" or " (no key yet)")
        )
      end
    end
  end
  for _, b in ipairs(CrimsonBindVars.binds) do
    if isBindRowActiveForCurrentSpec(b) and b.key and b.key ~= "" and bindKeyUsesNumpadBase(b.key) then
      table.insert(lines, (b.section or "") .. ": numpad key not applied (change or clear): " .. tostring(b.actionName))
    end
  end
  for _, b in ipairs(CrimsonBindVars.binds) do
    if isBindRowActiveForCurrentSpec(b) and b.key and b.key ~= "" and bindKeyUsesBlockedCtrlMinusOrBare(b.key) then
      table.insert(
        lines,
        (b.section or "") .. ": CTRL-MINUS / bare MINUS not applied (unreliable in WoW — use another key): " .. tostring(b.actionName)
      )
    end
  end
  for _, line in ipairs(competingBindingLines()) do
    table.insert(lines, line)
  end
  return lines
end

-- =========================
-- Validation engine
-- =========================

--- Per-row status: ok | missing | duplicate | conflict | invalid | numpad | ctrlminus | taken | toolong | emptymacro | plainmacro
local function buildValidationMeta()
  wipe(CB.rowStatus)
  CB.dupCount = 0
  CB.issueCount = 0
  wipe(CB.externalActionByKey)
  local extByKey = getWoWKeyCompetitorsForPlannedApply()
  for k, a in pairs(extByKey) do
    CB.externalActionByKey[k] = a
  end
  local curSec = resolveCurrentBindSection()
  local bySecNk = {}
  for i, b in ipairs(CrimsonBindVars.binds) do
    local sec = b.section or ""
    local nk = bindingTokenFromKeyString(b.key)
    if nk ~= "" then
      bySecNk[sec] = bySecNk[sec] or {}
      bySecNk[sec][nk] = bySecNk[sec][nk] or {}
      table.insert(bySecNk[sec][nk], i)
    end
  end
  local generalNk = {}
  for i, b in ipairs(CrimsonBindVars.binds) do
    if b.section == "General" then
      local nk = bindingTokenFromKeyString(b.key)
      if nk ~= "" then generalNk[nk] = true end
    end
  end
  local classPref = classBindSectionPrefix()
  for i, b in ipairs(CrimsonBindVars.binds) do
    local sec = b.section or ""
    local nk = bindingTokenFromKeyString(b.key)
    local inactiveRow = false
    if sec ~= "General" and strupper(sec) ~= "CUSTOM" then
      if curSec then
        if not bindSectionsEquivalent(sec, curSec) then
          inactiveRow = true
        end
      else
        if not classPref or sec:sub(1, #classPref):lower() ~= classPref:lower() then
          inactiveRow = true
        end
      end
    end
    if inactiveRow then
      CB.rowStatus[i] = "inactive"
    else
      local missing = (nk == "")
      local dup = false
      if not missing then
        local t = bySecNk[sec] and bySecNk[sec][nk]
        dup = t and #t > 1
      end
      local conflict = not missing and sec ~= "General" and generalNk[nk]
      local invalid = false
      if isTargetMember110(b.actionName) and b.key and b.key ~= "" then
        local low = b.key:lower()
        if low:find("ctrl") or low:find("alt") then
          invalid = true
        end
      end
      local macroRaw = normalizeStoredMacroText(b.macroText or "")
      local toolong = #macroRaw > MACRO_MAX
      local emptymacro = (nk ~= "") and (macroRaw == "")
      local plainBad = (macroRaw ~= "") and macroPlainSayIssue(b.macroText) and (nk ~= "")
      local numpadDeny = (b.key and b.key ~= "") and bindKeyUsesNumpadBase(b.key)
      local ctrlMinusDeny = (b.key and b.key ~= "") and bindKeyUsesBlockedCtrlMinusOrBare(b.key)
      local st = "ok"
      if invalid then
        st = "invalid"
      elseif numpadDeny then
        st = "numpad"
      elseif ctrlMinusDeny then
        st = "ctrlminus"
      elseif dup then
        st = "duplicate"
      elseif conflict then
        st = "conflict"
      elseif toolong then
        st = "toolong"
      elseif emptymacro then
        st = "emptymacro"
      elseif plainBad then
        st = "plainmacro"
      elseif missing then
        st = "missing"
      end
      local active = (sec == "General" or (curSec and bindSectionsEquivalent(sec, curSec)))
      if st == "ok" and active and nk ~= "" and extByKey[nk] then
        st = "taken"
      end
      CB.rowStatus[i] = st
      if st ~= "ok" then
        CB.issueCount = CB.issueCount + 1
      end
      if dup then
        CB.dupCount = CB.dupCount + 1
      end
    end
  end
end

local function rowHasIssue(i)
  local s = CB.rowStatus[i]
  return s and s ~= "ok" and s ~= "inactive"
end

local function buildFiltered()
  wipe(CB.filtered)
  local want = trim(CB.searchText or "")
  local wantLower = want:lower()
  local secFilter = CB.sectionFilter
  local curResolved = resolveCurrentBindSection()
  for idx, b in ipairs(CrimsonBindVars.binds or {}) do
    local sec = b.section or ""
    local skipSec = false
    if CB.filterToCurrentSpec then
      -- CUSTOM is always visible alongside General + current spec
      if strupper(sec) ~= "CUSTOM" then
        if curResolved then
          if sec ~= "General" and not bindSectionsEquivalent(sec, curResolved) then
            skipSec = true
          end
        else
          local pref = classBindSectionPrefix()
          if pref and sec ~= "General" and sec:sub(1, #pref):lower() ~= pref:lower() then
            skipSec = true
          elseif not pref and sec ~= "General" then
            skipSec = true
          end
        end
      end
    elseif secFilter and sec ~= secFilter then
      skipSec = true
    end
    if skipSec then
    elseif CB.conflictsOnly and not rowHasIssue(idx) then
    else
      if want == "" then
        table.insert(CB.filtered, { index = idx, bind = b })
      else
        local blob = (sec .. " " .. (b.actionName or "") .. " " .. (b.key or "") .. " " .. normalizeStoredMacroText(b.macroText or "")):lower()
        if blob:find(wantLower, 1, true) then
          table.insert(CB.filtered, { index = idx, bind = b })
        end
      end
    end
  end
end

local function sectionCompare(sx, sy)
  if sx == "General" and sy ~= "General" then return true end
  if sy == "General" and sx ~= "General" then return false end
  return sx:lower() < sy:lower()
end

local function buildDisplayList()
  wipe(CB.displayList)
  if #CB.filtered == 0 then return end
  local sorted = {}
  for _, e in ipairs(CB.filtered) do
    table.insert(sorted, e)
  end
  table.sort(sorted, function(x, y)
    local sx, sy = x.bind.section or "", y.bind.section or ""
    if sx ~= sy then
      return sectionCompare(sx, sy)
    end
    local ax, ay = (x.bind.actionName or ""):lower(), (y.bind.actionName or ""):lower()
    return ax < ay
  end)
  local lastSec = nil
  for _, e in ipairs(sorted) do
    local sec = e.bind.section or ""
    if sec ~= lastSec then
      table.insert(CB.displayList, { kind = "header", section = sec })
      lastSec = sec
    end
    table.insert(CB.displayList, { kind = "row", index = e.index, bind = e.bind })
  end
end

local refreshList
local updateScrollRange
local renderVisibleRows
local updateSummary

local UNDO_MAX = 20

local function pushUndoEntry(entry)
  table.insert(CB.undoStack, entry)
  while #CB.undoStack > UNDO_MAX do
    table.remove(CB.undoStack, 1)
  end
end

local function applyKeyChange(bindIndex, newKey)
  local b = CrimsonBindVars.binds[bindIndex]
  if not b then return end
  local oldKey = b.key or ""
  pushUndoEntry({ bindIndex = bindIndex, oldKey = oldKey, oldMacro = b.macroText or "" })
  local nk = trim(newKey or "")
  if nk ~= "" then
    nk = bindingTokenFromKeyString(nk)
    if bindKeyUsesNumpadBase(nk) then
      print("|cff00ccffCrimsonBind|r Numpad keys are disabled. Use the main keyboard or clear the key (DELETE while capturing).")
      return
    end
    if bindKeyUsesBlockedCtrlMinusOrBare(nk) then
      print("|cff00ccffCrimsonBind|r CTRL-MINUS and bare MINUS are not supported (WoW binding issues). Choose another key or clear with DELETE.")
      return
    end
  end
  b.key = nk
  table.insert(CrimsonBindVars.pendingEdits, {
    section = b.section,
    actionName = b.actionName,
    oldKey = oldKey,
    newKey = b.key,
    editType = "key",
  })
  if not InCombatLockdown() then
    CB.applyAll()
  else
    CB.pendingApply = true
  end
  buildValidationMeta()
  refreshList()
end

--- If WoW reports a nav key shared with numpad (Num Lock off), prefer NUMPADn only when that numpad key is down.
--- Do not map LEFT/RIGHT/UP/DOWN: Shift+arrow must stay SHIFT-RIGHT etc., not SHIFT-NUMPAD6 (different WoW bindings).
local NAV_KEY_NAME_TO_NUMPAD_IF_DOWN = {
  INSERT = "NUMPAD0",
  END = "NUMPAD1",
  PAGEDOWN = "NUMPAD3",
  HOME = "NUMPAD7",
  PAGEUP = "NUMPAD9",
}

local function buildBindingStringFromKeyPress(key)
  if not key or key == "UNKNOWN" then return nil end
  if key == "LSHIFT" or key == "RSHIFT" or key == "LCTRL" or key == "RCTRL" or key == "LALT" or key == "RALT" then
    return nil
  end
  local prefix = ""
  if IsShiftKeyDown() then prefix = prefix .. "SHIFT-" end
  if IsControlKeyDown() then prefix = prefix .. "CTRL-" end
  if IsAltKeyDown() then prefix = prefix .. "ALT-" end
  local k = strupper(key)
  if k == "-" then
    k = "MINUS"
  end
  local opCanon = BINDING_NUMPAD_OPERATOR_CANON[k]
  if opCanon then
    k = opCanon
  elseif k == "SUBTRACT" then
    if type(IsKeyDown) == "function" and IsKeyDown("NUMPADMINUS") then
      k = "NUMPADMINUS"
    else
      k = "MINUS"
    end
  elseif k == "ADD" then
    if type(IsKeyDown) == "function" and IsKeyDown("NUMPADPLUS") then
      k = "NUMPADPLUS"
    else
      k = "EQUALS"
    end
  elseif k == "PLUS" then
    if type(IsKeyDown) == "function" and IsKeyDown("NUMPADPLUS") then
      k = "NUMPADPLUS"
    else
      k = "EQUALS"
    end
  elseif k == "MINUS" then
    if type(IsKeyDown) == "function" and IsKeyDown("NUMPADMINUS") then
      k = "NUMPADMINUS"
    end
  elseif k == "SLASH" then
    if type(IsKeyDown) == "function" and IsKeyDown("NUMPADDIVIDE") then
      k = "NUMPADDIVIDE"
    end
  end
  if k == "PRIOR" then k = "PAGEUP" end
  if k == "NEXT" then k = "PAGEDOWN" end
  if type(IsKeyDown) == "function" and NAV_KEY_NAME_TO_NUMPAD_IF_DOWN[k] then
    local preferred = NAV_KEY_NAME_TO_NUMPAD_IF_DOWN[k]
    if IsKeyDown(preferred) then
      k = preferred
    else
      local down = {}
      for i = 0, 9 do
        local p = (i == 0) and "NUMPAD0" or ("NUMPAD" .. i)
        if IsKeyDown(p) then
          down[#down + 1] = p
        end
      end
      if #down == 1 then
        k = down[1]
      end
    end
  end
  -- When WoW reports only "0".."9", promote to NUMPADn only if that numpad key is down (Num Lock on); else keep top-row digit.
  if type(IsKeyDown) == "function" and #k == 1 and k >= "0" and k <= "9" then
    local padName = (k == "0") and "NUMPAD0" or ("NUMPAD" .. k)
    if IsKeyDown(padName) then
      k = padName
    end
  end
  return prefix .. k
end

local function stopKeyCapture()
  CB.capturingKeyForIndex = nil
  CB.keyCaptureToEditor = nil
  if CB.captureOverlay then
    CB.captureOverlay:Hide()
  end
  if CB.captureKeyboard then
    CB.captureKeyboard:Hide()
    CB.captureKeyboard:EnableKeyboard(false)
    CB.captureKeyboard:ClearFocus()
  end
  if CB.captureBanner then
    CB.captureBanner:Hide()
  end
end

local function startKeyCapture(bindIndex)
  if InCombatLockdown() then
    print("|cff00ccffCrimsonBind|r Exit combat to capture a new key.")
    return
  end
  CB.capturingKeyForIndex = bindIndex
  CB.keyCaptureToEditor = nil
  if CB.searchBox then CB.searchBox:ClearFocus() end
  if CB.captureOverlay then
    CB.captureOverlay:Show()
  end
  if CB.captureKeyboard then
    CB.captureKeyboard:Show()
    CB.captureKeyboard:EnableKeyboard(true)
    CB.captureKeyboard:SetFocus()
  end
  if CB.captureBanner then
    CB.captureBanner:Show()
    CB.captureBanner.text:SetText("Press a key |cffaaaaaa(ESC cancel, DELETE clear)|r")
  end
end

local function startKeyCaptureForEditor()
  if InCombatLockdown() then
    print("|cff00ccffCrimsonBind|r Exit combat to capture a new key.")
    return
  end
  if not CB.editorKeyEdit then
    return
  end
  CB.keyCaptureToEditor = CB.editorKeyEdit
  CB.capturingKeyForIndex = nil
  if CB.searchBox then CB.searchBox:ClearFocus() end
  if CB.editorMacroEdit then CB.editorMacroEdit:ClearFocus() end
  if CB.captureOverlay then
    CB.captureOverlay:Show()
  end
  if CB.captureKeyboard then
    CB.captureKeyboard:Show()
    CB.captureKeyboard:EnableKeyboard(true)
    CB.captureKeyboard:SetFocus()
  end
  if CB.captureBanner then
    CB.captureBanner:Show()
    CB.captureBanner.text:SetText("Key for editor |cffaaaaaa(ESC cancel, DELETE clear)|r")
  end
end

local function openKeyDialogFallback(bindIndex)
  local b = CrimsonBindVars.binds[bindIndex]
  if not b then return end
  StaticPopup_Show("CRIMSONBIND_SETKEY", b.actionName, nil, {
    bindIndex = bindIndex,
    existing = b.key or "",
  })
end

StaticPopupDialogs["CRIMSONBIND_SETKEY"] = {
  text = "New key for %s (numpad disabled — use main keyboard). Prefer capture. Empty clears.",
  button1 = OKAY,
  button2 = CANCEL,
  hasEditBox = true,
  maxLetters = 64,
  OnShow = function(self)
    self.editBox:SetText(self.data.existing or "")
    self.editBox:HighlightText()
  end,
  OnAccept = function(self)
    if InCombatLockdown() then
      print("|cff00ccffCrimsonBind|r Exit combat to change keys.")
      return
    end
    local t = trim(self.editBox:GetText())
    local idx = self.data.bindIndex
    applyKeyChange(idx, t)
  end,
  timeout = 0,
  whileDead = true,
  hideOnEscape = true,
  preferredIndex = 3,
}

local function truncateStr(s, maxLen)
  if not s then return "" end
  if #s <= maxLen then return s end
  return s:sub(1, maxLen - 2) .. ".."
end

local function statusLabelForRow(i)
  local s = CB.rowStatus[i] or "ok"
  if s == "duplicate" then return COLOR_ERR .. "!" .. "|r", "Duplicate key in section"
  elseif s == "conflict" then return COLOR_WARN .. "!" .. "|r", "Key also used in General"
  elseif s == "invalid" then return COLOR_WARN .. "!" .. "|r", "Invalid (TM1-10: no CTRL/ALT)"
  elseif s == "toolong" then return COLOR_ERR .. "L" .. "|r", "Macro exceeds " .. MACRO_MAX .. " characters"
  elseif s == "emptymacro" then return COLOR_ERR .. "M" .. "|r", "Key set but macro text is empty"
  elseif s == "plainmacro" then return COLOR_ERR .. "S" .. "|r", "Macro has plain text (WoW uses /say) or import placeholder"
  elseif s == "missing" then return COLOR_MISS .. "-" .. "|r", "No key"
  elseif s == "taken" then return COLOR_WARN .. "W" .. "|r", "WoW binds this key to something else (Apply overwrites)"
  elseif s == "inactive" then return COLOR_SECTION .. "." .. "|r", "Other spec — not loaded (General + this spec only apply)"
  elseif s == "numpad" then return COLOR_WARN .. "N" .. "|r", "Numpad keys disabled — change key or clear"
  elseif s == "ctrlminus" then return COLOR_WARN .. "H" .. "|r", "CTRL-MINUS / bare MINUS not supported"
  end
  return COLOR_OK .. "+" .. "|r", "OK"
end

-- =========================
-- List refresh (virtual rows)
-- =========================

renderVisibleRows = function()
  for r = 1, CB.VISIBLE_ROWS do
    local row = CB.rowFrames[r]
    local entry = CB.displayList[CB.scrollOffset + r]
    if not row then
    elseif not entry then
      row:Hide()
    elseif entry.kind == "header" then
      row.isHeader = true
      row.bindIndex = nil
      row.headerText:SetText(COLOR_ACTION .. "== " .. truncateStr(entry.section, 48) .. " ==" .. "|r")
      row.headerText:Show()
      row.sec:Hide()
      row.action:Hide()
      row.keyBtn:Hide()
      row.status:Hide()
      if row.icon then row.icon:Hide() end
      if row.editBtn then
        row.editBtn:Hide()
      end
      row:SetScript("OnClick", nil)
      row:SetScript("OnEnter", nil)
      row:SetScript("OnLeave", nil)
      row:Show()
    else
      row.isHeader = false
      local b = entry.bind
      local i = entry.index
      row.bindIndex = i
      row.headerText:Hide()
      row.sec:Show()
      row.action:Show()
      row.keyBtn:Show()
      row.status:Show()
      if row.icon then
        row.icon:SetTexture(resolveIconForBind(b))
        row.icon:Show()
      end
      if row.editBtn then
        row.editBtn:Show()
      end
      local slabel, stip = statusLabelForRow(i)
      row.status:SetText(slabel)
      row.sec:SetText(COLOR_SECTION .. truncateStr(b.section or "", 18) .. "|r")
      local act = b.actionName or ""
      local actDisp = act
      local want = trim(CB.searchText or "")
      if want ~= "" and act:lower():find(want:lower(), 1, true) then
        actDisp = COLOR_ACTION .. act .. "|r"
      else
        actDisp = COLOR_ACTION .. truncateStr(act, 24) .. "|r"
      end
      row.action:SetText(actDisp)
      local ky = keyStringForDisplay(b.key)
      if ky == "" then
        row.keyText:SetText(COLOR_MISS .. "(none)|r")
      else
        row.keyText:SetText(COLOR_KEY .. truncateStr(ky, 14) .. "|r")
      end
      local st = CB.rowStatus[i] or "ok"
      if st == "duplicate" or st == "conflict" or st == "toolong" or st == "emptymacro" or st == "plainmacro" then
        row.keyText:SetText(COLOR_ERR .. truncateStr(ky == "" and "(none)" or ky, 14) .. "|r")
      elseif st == "invalid" or st == "numpad" or st == "ctrlminus" then
        row.keyText:SetText(COLOR_WARN .. truncateStr(ky, 14) .. "|r")
      elseif st == "missing" then
        row.keyText:SetText(COLOR_MISS .. "(none)|r")
      elseif st == "taken" then
        row.keyText:SetText(COLOR_WARN .. truncateStr(ky, 14) .. "|r")
      elseif st == "inactive" then
        row.keyText:SetText(COLOR_SECTION .. truncateStr(ky == "" and "—" or ky, 14) .. "|r")
      end
      row:SetScript("OnEnter", function(self)
        GameTooltip:SetOwner(self, "ANCHOR_RIGHT")
        GameTooltip:AddLine(b.actionName or "?", 1, 1, 1)
        GameTooltip:AddLine(COLOR_SECTION .. (b.section or "") .. "|r", 1, 1, 1, true)
        local bkTip = bindKeyFromBind(b)
        GameTooltip:AddLine("Key: " .. COLOR_KEY .. (bkTip ~= "" and bkTip or "(none)") .. "|r", 1, 1, 1, true)
        GameTooltip:AddLine("Status: " .. stip, 0.8, 0.8, 0.8, true)
        if st == "inactive" then
          GameTooltip:AddLine(
            "Not loaded: CrimsonBind only applies General + your current talent spec. Switch to this row's spec, enable \"This spec + General\", or use the section filter for viewing only, then click Apply.",
            1,
            0.72,
            0.45,
            true
          )
        end
        if bkTip ~= "" and CB.externalActionByKey[bkTip] then
          GameTooltip:AddLine("WoW currently: |cffffcc00" .. truncateStr(CB.externalActionByKey[bkTip], 90) .. "|r", 1, 0.85, 0.5, true)
        end
        local mac = normalizeStoredMacroText(b.macroText or "")
        if mac ~= "" then
          GameTooltip:AddLine("Macro:", 0.7, 0.7, 0.9)
          for line in mac:gmatch("[^\n]+") do
            GameTooltip:AddLine(truncateStr(line, 200), 0.9, 0.9, 0.9, true)
          end
        end
        GameTooltip:Show()
      end)
      row:SetScript("OnLeave", function()
        GameTooltip:Hide()
      end)
      row:SetScript("OnClick", function(self, button)
        if button == "LeftButton" and IsShiftKeyDown() then
          local bi = self.bindIndex
          local bnd = bi and CrimsonBindVars.binds[bi]
          if bnd then placeBindOnCursor(bnd) end
          return
        end
        if button == "RightButton" then
          openKeyDialogFallback(self.bindIndex)
        else
          startKeyCapture(self.bindIndex)
        end
      end)
      row.keyBtn:SetScript("OnClick", function(_, button)
        if button == "RightButton" then
          openKeyDialogFallback(row.bindIndex)
        else
          startKeyCapture(row.bindIndex)
        end
      end)
      if row.editBtn then
        row.editBtn.bindIndex = i
        row.editBtn:SetScript("OnClick", function()
          if CB.openBindEditor then
            CB.openBindEditor(i)
          end
        end)
      end
      row:Show()
    end
  end
end

updateSummary = function()
  if CB.summaryBar then
    local bindsTbl = CrimsonBindVars.binds or {}
    local total = #bindsTbl
    local shown = #CB.filtered
    local issues = CB.issueCount
    local dups = CB.dupCount
    local wowKeys = 0
    for _ in pairs(CB.externalActionByKey) do
      wowKeys = wowKeys + 1
    end
    local testTag = CB.testMode and (COLOR_WARN .. " | TEST MODE|r") or ""
    CB.summaryBar:SetText(string.format(
      "%s%d binds|r (%s%d shown|r) | %s%d issues|r | %s%d dup|r | %s%d vs WoW|r%s",
      COLOR_ACTION, total,
      COLOR_KEY, shown,
      issues > 0 and COLOR_ERR or COLOR_OK, issues,
      dups > 0 and COLOR_WARN or COLOR_OK, dups,
      wowKeys > 0 and COLOR_WARN or COLOR_OK, wowKeys,
      testTag
    ))
  end
  if CB.statusText then
    local rows = #CB.displayList
    local hi = rows > 0 and math.min(CB.scrollOffset + CB.VISIBLE_ROWS, rows) or 0
    CB.statusText:SetText(string.format(
      "%d row issue(s) | Rows %d-%d of %d | /cb help",
      CB.issueCount, rows > 0 and (CB.scrollOffset + 1) or 0, hi, rows
    ))
  end
  if CB.emptyLabel then
    local bindsTbl = CrimsonBindVars.binds or {}
    if #bindsTbl == 0 then
      CB.emptyLabel:SetText("|cffaaaaaaNo binds loaded.|r\nUse Sync-CrimsonBinds (CSV/config) or place CrimsonBind.lua in SavedVariables, then /reload.")
      CB.emptyLabel:Show()
    elseif #CB.displayList == 0 then
      CB.emptyLabel:SetText("|cffaaaaaaNo matching binds.|r\nClear search or widen filters (Section / This spec + General).")
      CB.emptyLabel:Show()
    else
      CB.emptyLabel:Hide()
    end
  end
end

refreshList = function()
  CB.cachedApplyList = getBindsToApply()
  buildValidationMeta()
  if CB.debug then
    print("|cff00ccffCrimsonBind|r [debug] refreshList: issueCount=" .. tostring(CB.issueCount) .. " dupCount=" .. tostring(CB.dupCount))
  end
  buildFiltered()
  buildDisplayList()
  local maxOffset = math.max(0, #CB.displayList - CB.VISIBLE_ROWS)
  if CB.scrollOffset > maxOffset then CB.scrollOffset = maxOffset end
  if CB.scrollOffset < 0 then CB.scrollOffset = 0 end
  renderVisibleRows()
  updateSummary()
  if updateScrollRange then updateScrollRange() end
  syncTestModeUi()
  CB.cachedApplyList = nil
end

local function undoLastChange()
  local e = CB.undoStack[#CB.undoStack]
  if not e then
    print("|cff00ccffCrimsonBind|r Nothing to undo.")
    return false
  end
  table.remove(CB.undoStack)
  local b = CrimsonBindVars.binds[e.bindIndex]
  if not b then
    return false
  end
  b.key = e.oldKey or ""
  b.macroText = e.oldMacro or ""
  if not InCombatLockdown() then
    CB.applyAll()
  else
    CB.pendingApply = true
  end
  buildValidationMeta()
  refreshList()
  print("|cff00ccffCrimsonBind|r Undo: restored " .. tostring(b.actionName or "?"))
  return true
end

StaticPopupDialogs["CRIMSONBIND_APPLY_CONFIRM"] = {
  text = "CrimsonBind will overwrite %d existing WoW or addon key binding(s). Continue?",
  button1 = OKAY,
  button2 = CANCEL,
  OnAccept = function()
    if InCombatLockdown() then
      print("|cff00ccffCrimsonBind|r Exit combat to apply binds.")
      return
    end
    if CB.applyAll() then
      print("|cff00ccffCrimsonBind|r Apply binds done.")
    end
    if CB.frame and CB.frame:IsShown() then
      refreshList()
    end
  end,
  timeout = 0,
  whileDead = true,
  hideOnEscape = true,
}

local function CrimsonBind_RequestApply()
  if InCombatLockdown() then
    print("|cff00ccffCrimsonBind|r Exit combat to apply binds.")
    return
  end
  local list = getBindsToApply()
  local n = countBindingReplacements(list)
  if n > 0 then
    StaticPopup_Show("CRIMSONBIND_APPLY_CONFIRM", n)
  else
    if CB.applyAll() then
      print("|cff00ccffCrimsonBind|r Apply binds done.")
    end
    if CB.frame and CB.frame:IsShown() then
      refreshList()
    end
  end
end

-- =========================
-- UI: Editor frame
-- =========================

local function createBindEditorFrame(parent)
  local ed = CreateFrame("Frame", "CrimsonBindEditorFrame", parent, "BackdropTemplate")
  ed:SetSize(318, 430)
  ed:SetPoint("LEFT", parent, "RIGHT", 10, 0)
  ed:SetFrameStrata("DIALOG")
  ed:SetFrameLevel(parent:GetFrameLevel() + 20)
  ed:SetBackdrop({
    bgFile = "Interface\\DialogFrame\\UI-DialogBox-Background",
    edgeFile = "Interface\\DialogFrame\\UI-DialogBox-Border",
    tile = true,
    tileSize = 32,
    edgeSize = 16,
    insets = { left = 8, right = 8, top = 8, bottom = 8 },
  })
  ed:Hide()

  CB.editorIcon = ed:CreateTexture(nil, "ARTWORK")
  CB.editorIcon:SetSize(32, 32)
  CB.editorIcon:SetPoint("TOPLEFT", 12, -10)
  CB.editorIcon:SetTexture(QUESTION_MARK_ICON)

  local title = ed:CreateFontString(nil, "OVERLAY", "GameFontNormalLarge")
  title:SetPoint("TOP", 16, -18)
  title:SetWidth(252)
  title:SetJustifyH("CENTER")
  ed.titleStr = title

  local keyLabel = ed:CreateFontString(nil, "OVERLAY", "GameFontNormalSmall")
  keyLabel:SetPoint("TOPLEFT", 16, -48)
  keyLabel:SetText("Key")

  CB.editorKeyEdit = CreateFrame("EditBox", nil, ed, "InputBoxTemplate")
  CB.editorKeyEdit:SetSize(120, 22)
  CB.editorKeyEdit:SetPoint("TOPLEFT", 16, -64)
  CB.editorKeyEdit:SetAutoFocus(false)
  CB.editorKeyEdit:SetMaxLetters(64)

  local btnCap = CreateFrame("Button", nil, ed, "UIPanelButtonTemplate")
  btnCap:SetSize(72, 22)
  btnCap:SetPoint("LEFT", CB.editorKeyEdit, "RIGHT", 6, 0)
  btnCap:SetText("Capture")
  btnCap:SetScript("OnClick", function()
    startKeyCaptureForEditor()
  end)

  local btnClrKey = CreateFrame("Button", nil, ed, "UIPanelButtonTemplate")
  btnClrKey:SetSize(60, 22)
  btnClrKey:SetPoint("LEFT", btnCap, "RIGHT", 4, 0)
  btnClrKey:SetText("Clear")
  btnClrKey:SetScript("OnClick", function()
    CB.editorKeyEdit:SetText("")
  end)

  local macLabel = ed:CreateFontString(nil, "OVERLAY", "GameFontNormalSmall")
  macLabel:SetPoint("TOPLEFT", 16, -94)
  macLabel:SetText("Macro text")

  local macroBg = CreateFrame("Frame", nil, ed, "BackdropTemplate")
  macroBg:SetPoint("TOPLEFT", 14, -110)
  macroBg:SetSize(290, 156)
  macroBg:SetBackdrop({
    bgFile = "Interface\\ChatFrame\\ChatFrameBackground",
    edgeFile = "Interface\\Tooltips\\UI-Tooltip-Border",
    tile = true,
    tileSize = 16,
    edgeSize = 8,
    insets = { left = 3, right = 3, top = 3, bottom = 3 },
  })
  -- Default ChatFrameBackground reads very light; tint to a dark panel (adjust RGBA to taste).
  macroBg:SetBackdropColor(0.07, 0.08, 0.1, 0.98)
  macroBg:SetBackdropBorderColor(0.35, 0.5, 0.55, 0.85)
  -- Clicking anywhere on the background (including inset/border padding) focuses the EditBox.
  macroBg:EnableMouse(true)
  macroBg:SetScript("OnMouseDown", function()
    CB.editorMacroEdit:SetFocus()
  end)

  CB.editorMacroEdit = CreateFrame("EditBox", nil, macroBg)
  CB.editorMacroEdit:SetFontObject("ChatFontNormal")
  CB.editorMacroEdit:SetTextColor(0.88, 0.92, 0.86)
  CB.editorMacroEdit:SetMultiLine(true)
  CB.editorMacroEdit:SetMaxLetters(4096)
  -- Fill the entire backdrop so every pixel of it is clickable on the EditBox itself.
  CB.editorMacroEdit:SetPoint("TOPLEFT", 6, -6)
  CB.editorMacroEdit:SetPoint("BOTTOMRIGHT", -6, 6)
  CB.editorMacroEdit:SetAutoFocus(false)
  CB.editorMacroEdit:EnableMouse(true)
  CB.editorMacroEdit:SetTextInsets(4, 4, 4, 4)
  CB.editorMacroEdit:SetScript("OnEscapePressed", function(self)
    self:ClearFocus()
  end)

  CB.editorCharLabel = ed:CreateFontString(nil, "OVERLAY", "GameFontNormalSmall")
  CB.editorCharLabel:SetPoint("TOPLEFT", macroBg, "BOTTOMLEFT", 0, -6)
  CB.editorCharLabel:SetText("0 / " .. MACRO_MAX .. " chars")

  local function updateEditorCharCount()
    local t = CB.editorMacroEdit:GetText() or ""
    local n = #t
    if n > MACRO_MAX then
      CB.editorCharLabel:SetTextColor(1, 0.2, 0.2)
      CB.editorCharLabel:SetText(n .. " / " .. MACRO_MAX .. " (too long)")
    else
      CB.editorCharLabel:SetTextColor(0.8, 0.8, 0.9)
      CB.editorCharLabel:SetText(n .. " / " .. MACRO_MAX .. " chars")
    end
  end
  CB.editorMacroEdit:SetScript("OnTextChanged", function()
    updateEditorCharCount()
  end)

  local btnRow = CreateFrame("Frame", nil, ed)
  btnRow:SetPoint("BOTTOM", 0, 14)
  btnRow:SetSize(300, 28)

  local btnSave = CreateFrame("Button", nil, btnRow, "UIPanelButtonTemplate")
  btnSave:SetSize(78, 24)
  btnSave:SetPoint("BOTTOMLEFT", 0, 0)
  btnSave:SetText("Save")
  btnSave:SetScript("OnClick", function()
    if InCombatLockdown() then
      print("|cff00ccffCrimsonBind|r Exit combat to save bind edits.")
      return
    end
    local idx = CB.editorBindIndex
    local b = idx and CrimsonBindVars.binds[idx]
    if not b then
      ed:Hide()
      return
    end
    local newKeyRaw = trim(CB.editorKeyEdit:GetText())
    local newKey = newKeyRaw == "" and "" or bindingTokenFromKeyString(newKeyRaw)
    if newKey ~= "" and bindKeyUsesNumpadBase(newKey) then
      print("|cff00ccffCrimsonBind|r Numpad keys are disabled. Edit the key to a non-numpad binding.")
      return
    end
    if newKey ~= "" and bindKeyUsesBlockedCtrlMinusOrBare(newKey) then
      print("|cff00ccffCrimsonBind|r CTRL-MINUS / bare MINUS are not supported. Use a different binding.")
      return
    end
    local newMacro = CB.editorMacroEdit:GetText() or ""
    if #newMacro > MACRO_MAX then
      print("|cff00ccffCrimsonBind|r Macro exceeds " .. MACRO_MAX .. " characters; shorten before Save.")
      return
    end
    local sayIss = macroPlainSayIssue(newMacro)
    if sayIss then
      print("|cffff8800CrimsonBind|r Warning: " .. sayIss .. " — fix macro or key will speak in chat / do nothing useful.")
    end
    local oldKey = CB.editorSnapshotKey or ""
    local oldMacro = CB.editorSnapshotMacro or ""
    -- For CUSTOM rows: apply name rename if changed
    local isCustom = strupper(b.section or "") == "CUSTOM"
    local newName = isCustom and trim(CB.editorNameEdit:GetText()) or nil
    if newName == "" then newName = b.actionName end
    local nameChanged = isCustom and newName ~= (b.actionName or "")
    if newKey == oldKey and newMacro == oldMacro and not nameChanged then
      ed:Hide()
      return
    end
    pushUndoEntry({ bindIndex = idx, oldKey = oldKey, oldMacro = oldMacro })
    local editType = "both"
    if newKey ~= oldKey and newMacro == oldMacro then
      editType = "key"
    elseif newKey == oldKey and newMacro ~= oldMacro then
      editType = "macro"
    end
    table.insert(CrimsonBindVars.pendingEdits, {
      section = b.section,
      actionName = b.actionName,
      oldKey = oldKey,
      newKey = newKey,
      editType = editType,
      oldMacro = oldMacro,
      newMacro = newMacro,
    })
    if isCustom and nameChanged then
      b.actionName = newName
    end
    b.key = newKey
    b.macroText = newMacro
    local newMacNorm = normalizeStoredMacroText(newMacro)
    if newMacNorm ~= "" then CB.iconCache[newMacNorm] = nil end
    if not CB.applyAll() then
      CB.pendingApply = true
    end
    buildValidationMeta()
    refreshList()
    ed:Hide()
    print("|cff00ccffCrimsonBind|r Saved bind: " .. tostring(b.actionName))
  end)

  local btnRevert = CreateFrame("Button", nil, btnRow, "UIPanelButtonTemplate")
  btnRevert:SetSize(78, 24)
  btnRevert:SetPoint("LEFT", btnSave, "RIGHT", 6, 0)
  btnRevert:SetText("Revert")
  btnRevert:SetScript("OnClick", function()
    CB.editorKeyEdit:SetText(CB.editorSnapshotKey or "")
    CB.editorMacroEdit:SetText(CB.editorSnapshotMacro or "")
    updateEditorCharCount()
  end)

  local btnCancel = CreateFrame("Button", nil, btnRow, "UIPanelButtonTemplate")
  btnCancel:SetSize(78, 24)
  btnCancel:SetPoint("LEFT", btnRevert, "RIGHT", 6, 0)
  btnCancel:SetText("Cancel")
  btnCancel:SetScript("OnClick", function()
    ed:Hide()
  end)

  -- Name editor row (visible only for CUSTOM rows)
  local nameLabel = ed:CreateFontString(nil, "OVERLAY", "GameFontNormalSmall")
  nameLabel:SetPoint("TOPLEFT", 16, -282)
  nameLabel:SetText("Name")
  CB.editorNameEdit = CreateFrame("EditBox", nil, ed, "InputBoxTemplate")
  CB.editorNameEdit:SetSize(200, 22)
  CB.editorNameEdit:SetPoint("TOPLEFT", 16, -298)
  CB.editorNameEdit:SetAutoFocus(false)
  CB.editorNameEdit:SetMaxLetters(64)
  CB.editorNameEdit:SetScript("OnEscapePressed", function(self) self:ClearFocus() end)
  CB.editorNameEdit:SetScript("OnEnterPressed", function(self) self:ClearFocus() end)
  CB.editorNameLabel = nameLabel

  -- Delete button (visible only for CUSTOM rows)
  local btnDelete = CreateFrame("Button", nil, ed, "UIPanelButtonTemplate")
  btnDelete:SetSize(78, 24)
  btnDelete:SetPoint("TOPRIGHT", -8, -282)
  btnDelete:SetText("|cffff4444Delete|r")
  btnDelete:SetScript("OnClick", function()
    local idx = CB.editorBindIndex
    local b = idx and CrimsonBindVars.binds[idx]
    if not b or strupper(b.section or "") ~= "CUSTOM" then return end
    if InCombatLockdown() then
      print("|cff00ccffCrimsonBind|r Exit combat to delete a custom bind.")
      return
    end
    StaticPopupDialogs["CRIMSONBIND_DELETE_CUSTOM"] = StaticPopupDialogs["CRIMSONBIND_DELETE_CUSTOM"] or {
      text = "Delete custom bind \"%s\"?",
      button1 = "Delete",
      button2 = "Cancel",
      OnAccept = function()
        local di = CB.editorBindIndex
        if di and CrimsonBindVars.binds[di] and strupper(CrimsonBindVars.binds[di].section or "") == "CUSTOM" then
          local name = CrimsonBindVars.binds[di].actionName or "?"
          local nk = bindingTokenFromKeyString(CrimsonBindVars.binds[di].key or "")
          if nk ~= "" then
            pcall(SetBinding, nk, nil)
            SaveBindings(GetCurrentBindingSet())
          end
          table.remove(CrimsonBindVars.binds, di)
          CB.editorBindIndex = nil
          buildValidationMeta()
          refreshList()
          print("|cff00ccffCrimsonBind|r Deleted custom bind: " .. name)
        end
      end,
      timeout = 0,
      whileDead = true,
      hideOnEscape = true,
    }
    StaticPopup_Show("CRIMSONBIND_DELETE_CUSTOM", b.actionName or "?")
    ed:Hide()
  end)
  CB.editorDeleteBtn = btnDelete

  local btnPlace = CreateFrame("Button", nil, ed, "UIPanelButtonTemplate")
  btnPlace:SetSize(100, 22)
  btnPlace:SetPoint("TOPLEFT", 16, -316)
  btnPlace:SetText("Place on bar")
  btnPlace:SetScript("OnEnter", function(self)
    GameTooltip:SetOwner(self, "ANCHOR_RIGHT")
    GameTooltip:AddLine("Place on action bar", 1, 1, 1)
    GameTooltip:AddLine("Spell/item binds: free pickup (no macro slot used).", 0.8, 0.8, 0.8, true)
    GameTooltip:AddLine("Complex macros: creates a real WoW macro (uses 1 slot).", 1, 0.82, 0, true)
    GameTooltip:AddLine("Also: Shift+click any row in the list.", 0.6, 0.8, 1, true)
    GameTooltip:Show()
  end)
  btnPlace:SetScript("OnLeave", GameTooltip_Hide)
  btnPlace:SetScript("OnClick", function()
    local idx = CB.editorBindIndex
    local b = idx and CrimsonBindVars.binds[idx]
    if not b then return end
    local editorMacro = CB.editorMacroEdit and CB.editorMacroEdit:GetText() or ""
    if editorMacro ~= "" then
      local snapshot = { actionName = b.actionName, macroText = editorMacro, section = b.section, textureId = b.textureId, key = b.key }
      placeBindOnCursor(snapshot)
    else
      placeBindOnCursor(b)
    end
  end)
  CB.editorPlaceBtn = btnPlace

  CB.editorFrame = ed

  CB.openBindEditor = function(bindIndex)
    local b = CrimsonBindVars.binds[bindIndex]
    if not b then
      return
    end
    CB.editorBindIndex = bindIndex
    CB.editorSnapshotKey = keyStringForDisplay(b.key or "")
    CB.editorSnapshotMacro = normalizeStoredMacroText(b.macroText or "")
    local isCustom = strupper(b.section or "") == "CUSTOM"
    if isCustom then
      ed.titleStr:SetText("|cffff9900[CUSTOM]|r " .. truncateStr(b.actionName or "?", 28))
      CB.editorNameEdit:SetText(b.actionName or "")
      CB.editorNameEdit:Show()
      CB.editorNameLabel:Show()
      CB.editorDeleteBtn:Show()
    else
      ed.titleStr:SetText(truncateStr(b.actionName or "?", 36))
      CB.editorNameEdit:Hide()
      CB.editorNameLabel:Hide()
      CB.editorDeleteBtn:Hide()
    end
    CB.editorKeyEdit:SetText(CB.editorSnapshotKey)
    CB.editorMacroEdit:SetText(CB.editorSnapshotMacro)
    updateEditorCharCount()
    if CB.editorIcon then
      CB.editorIcon:SetTexture(resolveIconForBind(b))
    end
    ed:Show()
  end
end

-- =========================
-- UI: Key Heatmap
-- =========================

local function refreshHeatmapData()
  if not CB.heatmapFrame or not CB.heatmapFrame:IsShown() then return end
  local curSec = getCurrentSection()
  local bindsByKey = {}
  for _, b in ipairs(CrimsonBindVars.binds or {}) do
    local nk = bindingTokenFromKeyString(b.key or "")
    if nk ~= "" then
      local sec = b.section or ""
      local secUp = strupper(sec)
      local active = (sec == "General") or (secUp == "CUSTOM") or (curSec and bindSectionsEquivalent(sec, curSec))
      if not bindsByKey[nk] then bindsByKey[nk] = {} end
      table.insert(bindsByKey[nk], { section = sec, actionName = b.actionName or "", active = active })
    end
  end
  local cells = CB.heatmapFrame.cells
  for fullKey, cell in pairs(cells) do
    local nk = normalizeKey(fullKey)
    local entries = bindsByKey[nk]
    local color
    if HEATMAP_EXCLUDED_BASES[cell.baseKey] then
      color = HEATMAP_COLORS.excluded_base
    elseif HEATMAP_EXCLUDED_KEYS[nk] then
      color = HEATMAP_COLORS.excluded
    elseif entries then
      local activeEntries = {}
      for _, e in ipairs(entries) do
        if e.active then table.insert(activeEntries, e) end
      end
      if #activeEntries > 1 then
        color = HEATMAP_COLORS.conflict
      elseif #activeEntries == 1 then
        local su = strupper(activeEntries[1].section)
        if su == "GENERAL" then
          color = HEATMAP_COLORS.general
        elseif su == "CUSTOM" then
          color = HEATMAP_COLORS.custom
        else
          color = HEATMAP_COLORS.spec
        end
      else
        color = HEATMAP_COLORS.available
      end
    else
      color = HEATMAP_COLORS.available
    end
    cell:SetBackdropColor(unpack(color))
    cell:SetScript("OnEnter", function(self)
      GameTooltip:SetOwner(self, "ANCHOR_RIGHT")
      GameTooltip:AddLine(fullKey ~= "" and fullKey or cell.baseKey, 1, 1, 1)
      if HEATMAP_EXCLUDED_BASES[cell.baseKey] then
        GameTooltip:AddLine("Excluded base (movement / action bar)", 0.6, 0.6, 0.6)
      elseif HEATMAP_EXCLUDED_KEYS[nk] then
        GameTooltip:AddLine("Excluded (OS shortcut / unreliable)", 0.8, 0.4, 0.4)
      elseif entries then
        for _, e in ipairs(entries) do
          local c = e.active and {1, 1, 1} or {0.5, 0.5, 0.5}
          GameTooltip:AddLine((e.active and "" or "[inactive] ") .. e.actionName .. "  (" .. e.section .. ")", c[1], c[2], c[3], true)
        end
        local ac = 0
        for _, e in ipairs(entries) do if e.active then ac = ac + 1 end end
        if ac > 1 then
          GameTooltip:AddLine("|cffff3333CONFLICT: " .. ac .. " active binds on this key|r", 1, 0.2, 0.2, true)
        end
      else
        GameTooltip:AddLine("Available", 0.5, 0.8, 0.5)
      end
      GameTooltip:Show()
    end)
    cell:SetScript("OnLeave", function() GameTooltip:Hide() end)
    cell:SetScript("OnMouseDown", function()
      if entries and #entries > 0 then
        local target = entries[1]
        for idx, b in ipairs(CrimsonBindVars.binds) do
          if (b.actionName or "") == target.actionName and (b.section or "") == target.section then
            if CB.frame then
              if not CB.frame:IsShown() then CB.frame:Show() end
              CB.searchBox:SetText(target.actionName)
              CB.searchText = target.actionName
              CB.scrollOffset = 0
              refreshList()
            end
            break
          end
        end
      end
    end)
  end
end

local function createHeatmapFrame()
  local hm = CreateFrame("Frame", "CrimsonBindHeatmapFrame", UIParent, "BackdropTemplate")
  hm:SetSize(460, 620)
  hm:SetPoint("CENTER")
  hm:SetFrameStrata("DIALOG")
  hm:SetMovable(true)
  hm:EnableMouse(true)
  hm:RegisterForDrag("LeftButton")
  hm:SetScript("OnDragStart", hm.StartMoving)
  hm:SetScript("OnDragStop", function(self) self:StopMovingOrSizing() end)
  hm:SetBackdrop({
    bgFile = "Interface\\DialogFrame\\UI-DialogBox-Background",
    edgeFile = "Interface\\DialogFrame\\UI-DialogBox-Border",
    tile = true, tileSize = 32, edgeSize = 32,
    insets = { left = 8, right = 8, top = 8, bottom = 8 },
  })
  tinsert(UISpecialFrames, hm:GetName())

  local title = hm:CreateFontString(nil, "OVERLAY", "GameFontNormalLarge")
  title:SetPoint("TOP", 0, -14)
  title:SetText("CrimsonBind Key Heatmap")

  local closeBtn = CreateFrame("Button", nil, hm, "UIPanelCloseButton")
  closeBtn:SetPoint("TOPRIGHT", -6, -6)

  local legend = hm:CreateFontString(nil, "OVERLAY", "GameFontDisableSmall")
  legend:SetPoint("TOPLEFT", 16, -34)
  legend:SetWidth(420)
  legend:SetJustifyH("LEFT")
  legend:SetText("|cff3366ccGeneral|r  |cff33b34aSpec|r  |cffd9b333Custom|r  |cffcc3333Conflict|r  |cff4d2626Excluded|r  |cff262628Available|r")

  local scroll = CreateFrame("ScrollFrame", "CrimsonBindHeatmapScroll", hm, "UIPanelScrollFrameTemplate")
  scroll:SetPoint("TOPLEFT", 12, -52)
  scroll:SetPoint("BOTTOMRIGHT", -32, 12)

  local content = CreateFrame("Frame", nil, scroll)
  content:SetWidth(410)
  scroll:SetScrollChild(content)

  hm.content = content
  hm.cells = {}

  local CELL_W, CELL_H, GAP = 28, 18, 1
  local MOD_COL_W = 42
  local HEADER_H = 16
  local REGION_GAP = 10
  local y = 0

  for _, region in ipairs(HEATMAP_REGIONS) do
    local rh = content:CreateFontString(nil, "OVERLAY", "GameFontNormalSmall")
    rh:SetPoint("TOPLEFT", 4, -y)
    rh:SetText("|cffffd100" .. region.name .. "|r")
    y = y + HEADER_H

    for ci, label in ipairs(region.labels) do
      local cl = content:CreateFontString(nil, "OVERLAY", "GameFontDisableSmall")
      cl:SetPoint("TOPLEFT", MOD_COL_W + (ci - 1) * (CELL_W + GAP), -y)
      cl:SetWidth(CELL_W)
      cl:SetJustifyH("CENTER")
      cl:SetText(label)
    end
    y = y + 14

    for _, mod in ipairs(HEATMAP_MODS) do
      local ml = content:CreateFontString(nil, "OVERLAY", "GameFontDisableSmall")
      ml:SetPoint("TOPLEFT", 0, -y)
      ml:SetWidth(MOD_COL_W - 2)
      ml:SetJustifyH("RIGHT")
      ml:SetText(mod.label)

      for ci, baseKey in ipairs(region.keys) do
        local fullKey = mod.prefix .. baseKey
        local cell = CreateFrame("Frame", nil, content, "BackdropTemplate")
        cell:SetSize(CELL_W, CELL_H)
        cell:SetPoint("TOPLEFT", MOD_COL_W + (ci - 1) * (CELL_W + GAP), -y)
        cell:SetBackdrop({ bgFile = "Interface\\Buttons\\WHITE8X8" })
        cell:SetBackdropColor(0.15, 0.15, 0.18, 0.9)
        cell:EnableMouse(true)
        cell.fullKey = fullKey
        cell.baseKey = baseKey
        cell.modPrefix = mod.prefix
        hm.cells[fullKey] = cell
      end
      y = y + CELL_H + GAP
    end
    y = y + REGION_GAP
  end

  content:SetHeight(y + 10)
  hm:SetScript("OnShow", function() refreshHeatmapData() end)
  hm:Hide()
  CB.heatmapFrame = hm
end

-- =========================
-- UI: Main frame
-- =========================

local function createUI()
  local f = CreateFrame("Frame", "CrimsonBindFrame", UIParent, "BackdropTemplate")
  -- Tall enough that list (from -120) + bottom bar (test mode + buttons ~100px) do not overlap.
  f:SetSize(640, 572)
  local fp = CrimsonBindVars.framePos
  if fp and fp.point and fp.x ~= nil and fp.y ~= nil then
    f:SetPoint(fp.point, UIParent, fp.relPoint or fp.point, fp.x, fp.y)
  else
    f:SetPoint("CENTER")
  end
  f:SetFrameStrata("DIALOG")
  f:SetMovable(true)
  f:EnableMouse(true)
  f:RegisterForDrag("LeftButton")
  f:SetScript("OnDragStart", f.StartMoving)
  f:SetScript("OnDragStop", function(self)
    self:StopMovingOrSizing()
    local point, _, relPoint, x, y = self:GetPoint(1)
    if point and x ~= nil and y ~= nil then
      CrimsonBindVars.framePos = {
        point = point,
        relPoint = relPoint or point,
        x = x,
        y = y,
      }
    end
  end)
  f:SetBackdrop({
    bgFile = "Interface\\DialogFrame\\UI-DialogBox-Background",
    edgeFile = "Interface\\DialogFrame\\UI-DialogBox-Border",
    tile = true,
    tileSize = 32,
    edgeSize = 32,
    insets = { left = 8, right = 8, top = 8, bottom = 8 },
  })
  tinsert(UISpecialFrames, f:GetName())

  local title = f:CreateFontString(nil, "OVERLAY", "GameFontNormalLarge")
  title:SetPoint("TOP", 0, -14)
  title:SetText("CrimsonBind")

  CB.testModeBanner = f:CreateFontString(nil, "OVERLAY", "GameFontNormal")
  CB.testModeBanner:SetPoint("TOP", 0, -32)
  CB.testModeBanner:SetTextColor(1, 0.82, 0)
  CB.testModeBanner:SetText("TEST MODE — chat shows each bind's stored key (CSV/panel), not the key you pressed")
  CB.testModeBanner:Hide()

  local closeBtn = CreateFrame("Button", nil, f, "UIPanelCloseButton")
  closeBtn:SetPoint("TOPRIGHT", -6, -6)

  CB.summaryBar = f:CreateFontString(nil, "OVERLAY", "GameFontNormalSmall")
  CB.summaryBar:SetPoint("TOPLEFT", 20, -46)
  CB.summaryBar:SetWidth(580)
  CB.summaryBar:SetJustifyH("LEFT")

  local searchInstr = f:CreateFontString(nil, "OVERLAY", "GameFontDisableSmall")
  searchInstr:SetText("Search binds...")
  searchInstr:SetAlpha(0.65)

  CB.searchBox = CreateFrame("EditBox", nil, f, "InputBoxTemplate")
  CB.searchBox:SetSize(220, 22)
  CB.searchBox:SetPoint("TOPLEFT", 20, -70)
  searchInstr:SetPoint("LEFT", CB.searchBox, "LEFT", 6, 0)
  CB.searchBox:SetAutoFocus(false)
  CB.searchBox:SetScript("OnTextChanged", function(self, userInput)
    if self:GetText() == "" then
      searchInstr:Show()
    else
      searchInstr:Hide()
    end
    if CB.searchDebounceTimer then
      CB.searchDebounceTimer:Cancel()
      CB.searchDebounceTimer = nil
    end
    CB.searchDebounceTimer = C_Timer.NewTimer(0.12, function()
      CB.searchDebounceTimer = nil
      CB.searchText = CB.searchBox:GetText()
      CB.scrollOffset = 0
      refreshList()
    end)
  end)
  CB.searchBox:SetScript("OnEditFocusGained", function()
    if CB.searchBox:GetText() == "" then searchInstr:Hide() end
  end)
  CB.searchBox:SetScript("OnEditFocusLost", function()
    if CB.searchBox:GetText() == "" then searchInstr:Show() end
  end)
  CB.searchBox:SetScript("OnEnterPressed", function(self) self:ClearFocus() end)

  local secLabel = f:CreateFontString(nil, "OVERLAY", "GameFontNormalSmall")
  secLabel:SetPoint("LEFT", CB.searchBox, "RIGHT", 16, 0)
  secLabel:SetText("Section")

  local secDropdown = CreateFrame("Frame", "CrimsonBindSectionDropdown", f, "UIDropDownMenuTemplate")
  secDropdown:SetPoint("LEFT", secLabel, "RIGHT", 4, -2)
  UIDropDownMenu_SetWidth(secDropdown, 180)
  UIDropDownMenu_Initialize(secDropdown, function(_, level)
    local info = UIDropDownMenu_CreateInfo()
    info.text = "(all)"
    info.func = function()
      CB.sectionFilter = nil
      UIDropDownMenu_SetText(secDropdown, "(all)")
      CB.scrollOffset = 0
      refreshList()
    end
    UIDropDownMenu_AddButton(info)
    local seen = {}
    for _, b in ipairs(CrimsonBindVars.binds) do
      local s = b.section
      if s and not seen[s] then
        seen[s] = true
        info = UIDropDownMenu_CreateInfo()
        info.text = s
        info.func = function()
          CB.sectionFilter = s
          UIDropDownMenu_SetText(secDropdown, s)
          CB.scrollOffset = 0
          refreshList()
        end
        UIDropDownMenu_AddButton(info)
      end
    end
  end)
  UIDropDownMenu_SetText(secDropdown, "(all)")

  local cbCurrentSpec = CreateFrame("CheckButton", "CrimsonBindCurrentSpecCB", f, "UICheckButtonTemplate")
  cbCurrentSpec:SetPoint("TOPRIGHT", -300, -66)
  local cbCurrentSpecText = _G[cbCurrentSpec:GetName() .. "Text"]
  if cbCurrentSpecText then
    cbCurrentSpecText:SetText("This spec + General")
  end
  cbCurrentSpec:SetChecked(CB.filterToCurrentSpec and true or false)
  cbCurrentSpec:SetScript("OnClick", function(self)
    CB.filterToCurrentSpec = self:GetChecked() and true or false
    if CB.filterToCurrentSpec then
      CB.sectionFilter = nil
      UIDropDownMenu_SetText(secDropdown, "(all)")
    end
    CB.scrollOffset = 0
    refreshList()
  end)
  cbCurrentSpec:SetScript("OnEnter", function(self)
    GameTooltip:SetOwner(self, "ANCHOR_RIGHT")
    GameTooltip:AddLine("This spec + General", 1, 1, 1)
    GameTooltip:AddLine("When checked, the list and Apply only include General and your current class/spec section. Uncheck to pick any section from the dropdown.", 1, 0.82, 0, true)
    GameTooltip:Show()
  end)
  cbCurrentSpec:SetScript("OnLeave", GameTooltip_Hide)

  local cbConflicts = CreateFrame("CheckButton", "CrimsonBindIssuesOnlyCB", f, "UICheckButtonTemplate")
  cbConflicts:SetPoint("TOPRIGHT", -100, -66)
  local cbConflictsText = _G[cbConflicts:GetName() .. "Text"]
  if cbConflictsText then
    cbConflictsText:SetText("Issues only")
  end
  cbConflicts:SetChecked(false)
  cbConflicts:SetScript("OnClick", function(self)
    CB.conflictsOnly = self:GetChecked() and true or false
    CB.scrollOffset = 0
    refreshList()
  end)
  cbConflicts:SetScript("OnEnter", function(self)
    GameTooltip:SetOwner(self, "ANCHOR_RIGHT")
    GameTooltip:AddLine("Issues only", 1, 1, 1)
    GameTooltip:AddLine("Show rows with duplicate keys, General/spec conflicts, missing keys, empty or overlong macros, invalid Target Member binds, CTRL-MINUS/bare MINUS, or keys WoW maps to something other than CrimsonBind.", 1, 0.82, 0, true)
    GameTooltip:Show()
  end)
  cbConflicts:SetScript("OnLeave", GameTooltip_Hide)

  CB.statusText = f:CreateFontString(nil, "OVERLAY", "GameFontDisableSmall")
  CB.statusText:SetPoint("TOPLEFT", 20, -94)
  CB.statusText:SetWidth(600)
  CB.statusText:SetJustifyH("LEFT")

  local colY = -106
  local hdr = f:CreateFontString(nil, "OVERLAY", "GameFontNormalSmall")
  hdr:SetPoint("TOPLEFT", 24, colY)
  hdr:SetText(COLOR_SECTION .. "Sts      Section       Action            Key      Ed|r")

  local scroll = CreateFrame("ScrollFrame", "CrimsonBindScroll", f)
  scroll:SetPoint("TOPLEFT", 16, -120)
  scroll:SetSize(600, CB.VISIBLE_ROWS * CB.ROW_HEIGHT + 6)
  local child = CreateFrame("Frame", nil, scroll)
  child:SetSize(590, 400)
  scroll:SetScrollChild(child)

  scroll:EnableMouseWheel(true)
  scroll:SetScript("OnMouseWheel", function(_, delta)
    local maxO = math.max(0, #CB.displayList - CB.VISIBLE_ROWS)
    local step = 3
    if delta > 0 then
      CB.scrollOffset = math.max(0, CB.scrollOffset - step)
    else
      CB.scrollOffset = math.min(maxO, CB.scrollOffset + step)
    end
    refreshList()
  end)

  local x0 = 4
  for r = 1, CB.VISIBLE_ROWS do
    local row = CreateFrame("Button", nil, child)
    row:SetSize(590, CB.ROW_HEIGHT - 1)
    row:SetPoint("TOPLEFT", x0, -(r - 1) * CB.ROW_HEIGHT)
    row:SetHighlightTexture("Interface\\QuestFrame\\UI-QuestTitleHighlight", "ADD")

    row.headerText = row:CreateFontString(nil, "OVERLAY", "GameFontNormal")
    row.headerText:SetPoint("LEFT", 4, 0)
    row.headerText:SetWidth(560)
    row.headerText:SetJustifyH("LEFT")

    row.status = row:CreateFontString(nil, "OVERLAY", "GameFontNormalSmall")
    row.status:SetPoint("LEFT", 2, 0)
    row.status:SetWidth(COL_STATUS)

    row.icon = row:CreateTexture(nil, "ARTWORK")
    row.icon:SetSize(20, 20)
    row.icon:SetPoint("LEFT", row.status, "RIGHT", 1, 0)
    row.icon:SetTexture(QUESTION_MARK_ICON)

    row.sec = row:CreateFontString(nil, "OVERLAY", "GameFontNormalSmall")
    row.sec:SetPoint("LEFT", row.icon, "RIGHT", 2, 0)
    row.sec:SetWidth(COL_SECTION)
    row.sec:SetJustifyH("LEFT")

    row.action = row:CreateFontString(nil, "OVERLAY", "GameFontNormalSmall")
    row.action:SetPoint("LEFT", row.sec, "RIGHT", 4, 0)
    row.action:SetWidth(COL_ACTION)
    row.action:SetJustifyH("LEFT")

    row.keyBtn = CreateFrame("Button", nil, row)
    row.keyBtn:SetSize(COL_KEY, CB.ROW_HEIGHT)
    row.keyBtn:SetPoint("LEFT", row.action, "RIGHT", 4, 0)
    row.keyBtn:SetHighlightTexture("Interface\\Buttons\\WHITE8X8", "ADD")
    row.keyBtn:GetHighlightTexture():SetVertexColor(0.3, 0.6, 0.9, 0.25)

    row.keyText = row.keyBtn:CreateFontString(nil, "OVERLAY", "GameFontNormalSmall")
    row.keyText:SetAllPoints()
    row.keyText:SetJustifyH("LEFT")

    row.editBtn = CreateFrame("Button", nil, row, "UIPanelButtonTemplate")
    row.editBtn:SetSize(32, 18)
    row.editBtn:SetPoint("LEFT", row.keyBtn, "RIGHT", 2, 0)
    row.editBtn:SetText("...")
    row.editBtn:Hide()

    CB.rowFrames[r] = row
  end

  CB.emptyLabel = child:CreateFontString(nil, "OVERLAY", "GameFontHighlightSmall")
  CB.emptyLabel:SetPoint("CENTER", child, "CENTER", 0, 40)
  CB.emptyLabel:SetWidth(520)
  CB.emptyLabel:SetJustifyH("CENTER")
  CB.emptyLabel:SetJustifyV("MIDDLE")
  CB.emptyLabel:SetText("")
  CB.emptyLabel:Hide()

  local sb = CreateFrame("Slider", nil, f, "OptionsSliderTemplate")
  sb:SetPoint("TOPRIGHT", scroll, "TOPRIGHT", 28, -4)
  sb:SetHeight(CB.VISIBLE_ROWS * CB.ROW_HEIGHT)
  sb:SetWidth(18)
  sb:SetOrientation("VERTICAL")
  sb:SetMinMaxValues(0, 1)
  sb:SetValueStep(1)
  sb:SetScript("OnValueChanged", function(_, v)
    CB.scrollOffset = math.floor(v + 0.5)
    refreshList()
  end)
  CB.scrollBar = sb

  updateScrollRange = function()
    local maxO = math.max(0, #CB.displayList - CB.VISIBLE_ROWS)
    CB.scrollBar:SetMinMaxValues(0, maxO)
    if maxO <= 0 then
      CB.scrollBar:Disable()
    else
      CB.scrollBar:Enable()
    end
    CB.scrollBar:SetValue(CB.scrollOffset)
  end

  CB.captureBanner = CreateFrame("Frame", nil, f, "BackdropTemplate")
  CB.captureBanner:SetPoint("BOTTOM", 0, 52)
  CB.captureBanner:SetSize(400, 28)
  CB.captureBanner:SetBackdrop({
    bgFile = "Interface\\DialogFrame\\UI-DialogBox-Background-Dark",
    edgeFile = "Interface\\Tooltips\\UI-Tooltip-Border",
    tile = true,
    tileSize = 16,
    edgeSize = 12,
    insets = { left = 4, right = 4, top = 4, bottom = 4 },
  })
  CB.captureBanner.text = CB.captureBanner:CreateFontString(nil, "OVERLAY", "GameFontNormal")
  CB.captureBanner.text:SetPoint("CENTER")
  CB.captureBanner:Hide()

  CB.captureOverlay = CreateFrame("Frame", nil, UIParent)
  CB.captureOverlay:SetFrameStrata("FULLSCREEN_DIALOG")
  CB.captureOverlay:SetAllPoints()
  CB.captureOverlay:EnableMouse(true)
  CB.captureOverlay:Hide()

  local function captureOnKeyDown(_, key)
    if CB.keyCaptureToEditor then
      local eb = CB.keyCaptureToEditor
      if key == "ESCAPE" then
        stopKeyCapture()
        return
      end
      if key == "DELETE" then
        eb:SetText("")
        stopKeyCapture()
        return
      end
      local bindStr = buildBindingStringFromKeyPress(key)
      if bindStr then
        if bindKeyUsesNumpadBase(bindStr) then
          print("|cff00ccffCrimsonBind|r Numpad is disabled — use a main-keyboard key.")
          return
        end
        if bindKeyUsesBlockedCtrlMinusOrBare(bindStr) then
          print("|cff00ccffCrimsonBind|r CTRL-MINUS / bare MINUS are not supported — pick another key.")
          return
        end
        eb:SetText(bindStr)
        stopKeyCapture()
      end
      return
    end
    local idx = CB.capturingKeyForIndex
    if not idx then
      stopKeyCapture()
      return
    end
    if key == "ESCAPE" then
      stopKeyCapture()
      refreshList()
      return
    end
    if key == "DELETE" then
      applyKeyChange(idx, "")
      stopKeyCapture()
      return
    end
    local bindStr = buildBindingStringFromKeyPress(key)
    if bindStr then
      if bindKeyUsesNumpadBase(bindStr) then
        print("|cff00ccffCrimsonBind|r Numpad is disabled — use a main-keyboard key.")
        return
      end
      if bindKeyUsesBlockedCtrlMinusOrBare(bindStr) then
        print("|cff00ccffCrimsonBind|r CTRL-MINUS / bare MINUS are not supported — pick another key.")
        return
      end
      applyKeyChange(idx, bindStr)
      stopKeyCapture()
    end
  end

  local function captureOnMouseDown(_, button)
    if CB.keyCaptureToEditor then
      local eb = CB.keyCaptureToEditor
      if button == "RightButton" then
        stopKeyCapture()
        return
      end
      if button == "MiddleButton" then
        eb:SetText("BUTTON3")
        stopKeyCapture()
        return
      end
      local map = { LeftButton = "BUTTON1", RightButton = "BUTTON2", Button4 = "BUTTON4", Button5 = "BUTTON5" }
      local bn = map[button]
      if bn then
        local prefix = ""
        if IsShiftKeyDown() then prefix = prefix .. "SHIFT-" end
        if IsControlKeyDown() then prefix = prefix .. "CTRL-" end
        if IsAltKeyDown() then prefix = prefix .. "ALT-" end
        eb:SetText(prefix .. bn)
        stopKeyCapture()
      end
      return
    end
    local idx = CB.capturingKeyForIndex
    if not idx then return end
    if button == "RightButton" then
      stopKeyCapture()
      refreshList()
      return
    end
    if button == "MiddleButton" then
      applyKeyChange(idx, "BUTTON3")
      stopKeyCapture()
      return
    end
    local map = { LeftButton = "BUTTON1", RightButton = "BUTTON2", Button4 = "BUTTON4", Button5 = "BUTTON5" }
    local bn = map[button]
    if bn then
      local prefix = ""
      if IsShiftKeyDown() then prefix = prefix .. "SHIFT-" end
      if IsControlKeyDown() then prefix = prefix .. "CTRL-" end
      if IsAltKeyDown() then prefix = prefix .. "ALT-" end
      applyKeyChange(idx, prefix .. bn)
      stopKeyCapture()
    end
  end

  -- Mists / Classic: Frame has no SetFocus; use a fullscreen EditBox for keyboard focus (same pattern as many key-capture UIs).
  CB.captureKeyboard = CreateFrame("EditBox", nil, CB.captureOverlay)
  CB.captureKeyboard:SetFrameStrata("FULLSCREEN_DIALOG")
  CB.captureKeyboard:SetAllPoints()
  CB.captureKeyboard:EnableMouse(true)
  CB.captureKeyboard:SetAutoFocus(false)
  CB.captureKeyboard:SetMultiLine(false)
  CB.captureKeyboard:SetFontObject("GameFontHighlight")
  CB.captureKeyboard:SetText("")
  CB.captureKeyboard:SetTextColor(0, 0, 0, 0)
  CB.captureKeyboard:EnableKeyboard(true)
  if CB.captureKeyboard.SetPropagateKeyboardInput then
    CB.captureKeyboard:SetPropagateKeyboardInput(false)
  end
  CB.captureKeyboard:SetScript("OnKeyDown", captureOnKeyDown)
  CB.captureKeyboard:SetScript("OnMouseDown", captureOnMouseDown)
  CB.captureKeyboard:SetScript("OnChar", function(self)
    if CB.capturingKeyForIndex or CB.keyCaptureToEditor then
      self:SetText("")
    end
  end)
  CB.captureKeyboard:Hide()

  local btnVal = CreateFrame("Button", nil, f, "UIPanelButtonTemplate")
  btnVal:SetSize(100, 24)
  btnVal:SetPoint("BOTTOMLEFT", 20, 14)
  btnVal:SetText("Validate")
  btnVal:SetScript("OnClick", function()
    local v = validationSummary()
    print("|cff00ccffCrimsonBind|r " .. #v .. " issue(s)")
    for _, line in ipairs(v) do
      print("  " .. line)
    end
    if CB.debug then
      print("|cff00ccffCrimsonBind|r [debug] validationSummary printed " .. #v .. " lines; row issueCount=" .. tostring(CB.issueCount))
    end
    refreshList()
  end)

  local btnSync = CreateFrame("Button", nil, f, "UIPanelButtonTemplate")
  btnSync:SetSize(100, 24)
  btnSync:SetPoint("LEFT", btnVal, "RIGHT", 6, 0)
  btnSync:SetText("Apply binds")
  btnSync:SetScript("OnClick", function()
    CrimsonBind_RequestApply()
  end)

  local btnResetSec = CreateFrame("Button", nil, f, "UIPanelButtonTemplate")
  btnResetSec:SetSize(120, 24)
  btnResetSec:SetPoint("LEFT", btnSync, "RIGHT", 6, 0)
  btnResetSec:SetText("Reset section")
  btnResetSec:SetScript("OnClick", function()
    CB.sectionFilter = nil
    UIDropDownMenu_SetText(secDropdown, "(all)")
    CB.scrollOffset = 0
    refreshList()
  end)

  local btnNewMacro = CreateFrame("Button", nil, f, "UIPanelButtonTemplate")
  btnNewMacro:SetSize(100, 24)
  btnNewMacro:SetPoint("LEFT", btnResetSec, "RIGHT", 6, 0)
  btnNewMacro:SetText("New macro")
  btnNewMacro:SetScript("OnEnter", function(self)
    GameTooltip:SetOwner(self, "ANCHOR_RIGHT")
    GameTooltip:AddLine("New custom macro", 1, 1, 1)
    GameTooltip:AddLine("Creates a new CUSTOM-section bind that is always active regardless of spec, never written to config.ini, and never re-keyed by the randomizer.", 1, 0.82, 0, true)
    GameTooltip:Show()
  end)
  btnNewMacro:SetScript("OnLeave", GameTooltip_Hide)
  btnNewMacro:SetScript("OnClick", function()
    if InCombatLockdown() then
      print("|cff00ccffCrimsonBind|r Exit combat to create a new macro.")
      return
    end
    CrimsonBindVars.binds = CrimsonBindVars.binds or {}
    local ts = math.floor(GetTime() * 1000) % 100000
    local newName = "Custom Macro " .. (#CrimsonBindVars.binds + 1) .. "-" .. ts
    table.insert(CrimsonBindVars.binds, {
      section = "CUSTOM",
      actionName = newName,
      key = "",
      macroText = "",
      textureId = 0,
      actionType = "macro",
      source = "custom",
    })
    buildValidationMeta()
    refreshList()
    local newIdx = #CrimsonBindVars.binds
    CB.openBindEditor(newIdx)
  end)

  local btnHeatmap = CreateFrame("Button", nil, f, "UIPanelButtonTemplate")
  btnHeatmap:SetSize(78, 24)
  btnHeatmap:SetPoint("BOTTOMRIGHT", -20, 14)
  btnHeatmap:SetText("Heatmap")
  btnHeatmap:SetScript("OnEnter", function(self)
    GameTooltip:SetOwner(self, "ANCHOR_RIGHT")
    GameTooltip:AddLine("Key Heatmap", 1, 1, 1)
    GameTooltip:AddLine("Visual grid of all modifier + base key combinations, color-coded by section, conflicts, and exclusions.", 0.8, 0.8, 0.8, true)
    GameTooltip:Show()
  end)
  btnHeatmap:SetScript("OnLeave", GameTooltip_Hide)
  btnHeatmap:SetScript("OnClick", function()
    if not CB.heatmapFrame then createHeatmapFrame() end
    if CB.heatmapFrame:IsShown() then
      CB.heatmapFrame:Hide()
    else
      CB.heatmapFrame:Show()
    end
  end)

  local cbTestMode = CreateFrame("CheckButton", "CrimsonBindTestModeCB", f, "UICheckButtonTemplate")
  cbTestMode:SetPoint("BOTTOMLEFT", btnVal, "TOPLEFT", 0, 10)
  local cbTestModeText = _G[cbTestMode:GetName() .. "Text"]
  if cbTestModeText then
    cbTestModeText:SetText("Test mode (print only; exit combat to apply)")
  end
  cbTestMode:SetChecked(CB.testMode)
  CB.testModeCheck = cbTestMode
  cbTestMode:SetScript("OnClick", function(self)
    CB.testMode = self:GetChecked() and true or false
    CrimsonBindVars.testMode = CB.testMode
    syncTestModeUi()
    if not CB.applyAll() then
      print("|cff00ccffCrimsonBind|r Test mode saved; will apply when you leave combat.")
    else
      print("|cff00ccffCrimsonBind|r Test mode: " .. (CB.testMode and "ON (print only)" or "OFF (real macros)"))
    end
    refreshList()
  end)

  local hint = f:CreateFontString(nil, "OVERLAY", "GameFontDisableSmall")
  hint:SetPoint("BOTTOMRIGHT", -16, 16)
  hint:SetWidth(340)
  hint:SetJustifyH("RIGHT")
  hint:SetText("Shift+click row = place on bar | /cb test toggles test mode | /cb heatmap")

  createBindEditorFrame(f)

  f:Hide()
  CB.frame = f
end

function CrimsonBind_Toggle()
  if not CB.frame then return end
  if CB.frame:IsShown() then
    stopKeyCapture()
    if CB.editorFrame then
      CB.editorFrame:Hide()
    end
    CB.frame:Hide()
  else
    refreshList()
    CB.frame:Show()
  end
end

local function onEvent(self, event, arg1)
  if event == "ADDON_LOADED" and arg1 == "CrimsonBind" then
    initSavedVars()
    loadTestModeFromSavedVars()
    CB.macroBtn = CreateFrame("Button", "CrimsonBindMacro", UIParent, "SecureActionButtonTemplate")
    CB.macroBtn:SetSize(1, 1)
    CB.macroBtn:SetPoint("TOPLEFT", UIParent, "BOTTOMRIGHT", 5000, 5000)
    CB.macroBtn:EnableMouse(true)
    CB.macroBtn:Show()
    syncMacroBtnRegisterForClicks()
    createUI()
    buildValidationMeta()
    refreshList()
    C_Timer.After(2, function()
      CB.applyAll()
    end)
  elseif event == "PLAYER_REGEN_ENABLED" then
    if CB.pendingApply then
      CB.applyAll()
    end
  elseif event == "ACTIVE_TALENT_GROUP_CHANGED" or event == "PLAYER_SPECIALIZATION_CHANGED" then
    if not InCombatLockdown() then
      C_Timer.After(0.5, function()
        CB.applyAll()
        refreshList()
      end)
    else
      CB.pendingApply = true
    end
  elseif event == "PLAYER_ENTERING_WORLD" or event == "PLAYER_ALIVE" then
    if CB.frame then
      buildValidationMeta()
      refreshList()
    end
  elseif event == "CVAR_UPDATE" and arg1 == "ActionButtonUseKeyDown" then
    if not InCombatLockdown() then
      syncMacroBtnRegisterForClicks()
    else
      CB.pendingApply = true
    end
  end
end

local frame = CreateFrame("Frame")
frame:RegisterEvent("ADDON_LOADED")
frame:RegisterEvent("PLAYER_REGEN_ENABLED")
frame:RegisterEvent("PLAYER_ENTERING_WORLD")
frame:RegisterEvent("PLAYER_ALIVE")
frame:RegisterEvent("ACTIVE_TALENT_GROUP_CHANGED")
frame:RegisterEvent("PLAYER_SPECIALIZATION_CHANGED")
frame:RegisterEvent("CVAR_UPDATE")
frame:SetScript("OnEvent", onEvent)

SLASH_CRIMSONBIND1 = "/cb"
SLASH_CRIMSONBIND2 = "/crimsonbind"
SlashCmdList["CRIMSONBIND"] = function(msg)
  msg = trim(msg):lower()
  if msg == "help" or msg == "?" then
    print("|cff00ccffCrimsonBind|r commands:")
    print("  |cffaaaaaa/cb|r or |cffaaaaaa/crimsonbind|r — open or close the panel")
    print("  |cffaaaaaa/cb toggle|r — same as bare |cffaaaaaa/cb|r")
    print("  |cffaaaaaa/cb apply|r — apply binds (confirms if WoW keys overwritten)")
    print("  |cffaaaaaa/cb validate|r — print issues to chat")
    print("  |cffaaaaaa/cb test|r — toggle test mode (print-only macros)")
    print("  |cffaaaaaa/cb count|r — bind count and current section")
    print("  |cffaaaaaa/cb list|r — print active binds with CB slot index (matches SetBinding CLICK)")
    print("  |cffaaaaaa/cb exportcustom|r — print CUSTOM binds as CSV rows for pasting into CrimsonBind_binds.csv")
    print("  |cffaaaaaa/cb heatmap|r — open/close the key heatmap (modifier x base grid)")
    print("  |cffaaaaaa/cb pickup <name>|r — place a bind on the cursor for action bar drag (spell/item = free, complex = macro slot)")
    print("  |cffaaaaaa/cb clearbindpad|r — unbind keys still pointing at BindPad CLICK (addon off does not clear WoW binds)")
    print("  |cffaaaaaa/cb undo|r — undo last key/macro edit (up to " .. UNDO_MAX .. ")")
    print("  |cffaaaaaa/cb debug|r — toggle debug logging")
    print("  |cffaaaaaa/cb status|r — pending edits, combat queue, test mode")
    print("  |cffaaaaaa/cb help|r — this list")
    return
  end
  if msg == "validate" or msg == "issues" then
    local v = validationSummary()
    print("|cff00ccffCrimsonBind|r " .. #v .. " issue(s)")
    for _, line in ipairs(v) do
      print("  " .. line)
    end
    return
  end
  if msg == "sync" or msg == "apply" then
    CrimsonBind_RequestApply()
    return
  end
  if msg == "list" then
    local list = getBindsToApply()
    print("|cff00ccffCrimsonBind|r active binds (" .. #list .. ") for " .. getCurrentSection() .. " + General (CBn = CLICK CrimsonBindMacro:CBn):")
    for i, b in ipairs(list) do
      print("  CB" .. i .. "  " .. tostring(b.actionName or "?") .. "  |  " .. tostring(b.key or ""))
    end
    return
  end
  if msg == "clearbindpad" then
    if InCombatLockdown() then
      print("|cff00ccffCrimsonBind|r Exit combat to clear BindPad bindings.")
      return
    end
    local n = clearBindPadClickBindingsFromWoW()
    print(
      "|cff00ccffCrimsonBind|r Cleared "
        .. tostring(n)
        .. " key(s) bound to BindPad. Run |cffaaaaaa/cb apply|r so CrimsonBind owns those keys again."
    )
    return
  end
  if msg == "undo" then
    undoLastChange()
    return
  end
  if msg == "debug" then
    CB.debug = not CB.debug
    print("|cff00ccffCrimsonBind|r Debug " .. (CB.debug and "ON" or "OFF"))
    return
  end
  if msg == "count" then
    print("|cff00ccffCrimsonBind|r binds: " .. #CrimsonBindVars.binds .. " | current section: " .. getCurrentSection())
    return
  end
  if msg == "status" then
    print("|cff00ccffCrimsonBind|r pendingEdits: " .. #(CrimsonBindVars.pendingEdits or {}) .. " | combat queue: " .. tostring(CB.pendingApply) .. " | test mode: " .. tostring(CB.testMode) .. " | debug: " .. tostring(CB.debug) .. " | undo: " .. tostring(#CB.undoStack))
    return
  end
  if msg == "toggle" then
    CrimsonBind_Toggle()
    return
  end
  if msg == "exportcustom" then
    local custom = {}
    for _, b in ipairs(CrimsonBindVars.binds or {}) do
      if strupper(b.section or "") == "CUSTOM" then
        table.insert(custom, b)
      end
    end
    if #custom == 0 then
      print("|cff00ccffCrimsonBind|r No CUSTOM binds found.")
      return
    end
    local function csvEscape(s)
      s = tostring(s or "")
      s = s:gsub('"', '""')
      return '"' .. s .. '"'
    end
    local function macroForCsv(macroText)
      local m = normalizeStoredMacroText(macroText or "")
      m = m:gsub("\n", "\194\167")
      return m
    end
    print("|cff00ccffCrimsonBind|r CUSTOM binds (" .. #custom .. ") — paste into CrimsonBind_binds.csv:")
    for _, b in ipairs(custom) do
      print(table.concat({
        csvEscape("CUSTOM"),
        csvEscape(b.actionName or ""),
        csvEscape(macroForCsv(b.macroText or "")),
        csvEscape(b.key or ""),
      }, ","))
    end
    return
  end
  if msg == "test" then
    CB.testMode = not CB.testMode
    CrimsonBindVars.testMode = CB.testMode
    syncTestModeUi()
    if not CB.applyAll() then
      print("|cff00ccffCrimsonBind|r Test mode " .. (CB.testMode and "ON" or "OFF") .. " (will apply when you leave combat).")
    else
      print(
        "|cff00ccffCrimsonBind|r Test mode: "
          .. (CB.testMode and "ON — chat shows each bind's stored key + action (not the key you pressed)" or "OFF — real macros")
      )
    end
    if CB.frame and CB.frame:IsShown() then
      refreshList()
    end
    return
  end
  if msg == "heatmap" then
    if not CB.heatmapFrame then createHeatmapFrame() end
    if CB.heatmapFrame:IsShown() then
      CB.heatmapFrame:Hide()
    else
      CB.heatmapFrame:Show()
    end
    return
  end
  if msg:sub(1, 6) == "pickup" then
    local want = trim(msg:sub(7))
    if want == "" then
      print("|cff00ccffCrimsonBind|r Usage: /cb pickup <action name>")
      return
    end
    local wantLower = want:lower()
    for _, b in ipairs(CrimsonBindVars.binds or {}) do
      if (b.actionName or ""):lower() == wantLower then
        placeBindOnCursor(b)
        return
      end
    end
    for _, b in ipairs(CrimsonBindVars.binds or {}) do
      if (b.actionName or ""):lower():find(wantLower, 1, true) then
        placeBindOnCursor(b)
        return
      end
    end
    print("|cff00ccffCrimsonBind|r No bind found matching '" .. want .. "'.")
    return
  end
  CrimsonBind_Toggle()
end

BINDING_HEADER_CRIMSONBIND = "CrimsonBind"
BINDING_NAME_TOGGLE_CRIMSONBIND = "Toggle CrimsonBind"
