-- LibreOffice as neoffice runs it (image previews, .doc conversions, PDF exports): always with a
-- profile of neoffice's own, never yours, made safe for documents from others before each start:
-- no macros (whatever the document asks), no active content (OLE objects, DDE links, linked
-- frames), no links followed to fetch what isn't in the document. Settings LibreOffice doesn't know
-- (an older one) are left alone by it.
local M = {}

local PATH = "/org.openoffice.Office.Common/Security/Scripting"
M.SETTINGS = {
  { "MacroSecurityLevel", "3" }, -- (very high: none but from trusted places, and neoffice trusts none)
  { "DisableMacrosExecution", "true" },
  { "DisableActiveContent", "true" },
  { "DisableOLEAutomation", "true" },
  { "BlockUntrustedRefererLinks", "true" },
}

local function item(name, value)
  return ('<item oor:path="%s"><prop oor:name="%s" oor:op="fuse"><value>%s</value></prop></item>'):format(PATH, name, value)
end

-- The profile directory `dir` with those settings (in its user/registrymodifications.xcu, made if
-- it isn't there, LibreOffice's other settings kept); returns its -env:UserInstallation argument.
function M.profile(dir)
  local file = dir .. "/user/registrymodifications.xcu"
  local f = io.open(file, "rb")
  local text = f and f:read("*a")
  if f then
    f:close()
  end
  if not text or not text:find("</oor:items>", 1, true) then
    text = '<?xml version="1.0" encoding="UTF-8"?>\n<oor:items xmlns:oor="http://openoffice.org/2001/registry"'
      .. ' xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">\n</oor:items>\n'
  end
  local changed = false
  for _, s in ipairs(M.SETTINGS) do
    local want = item(s[1], s[2])
    if not text:find(want, 1, true) then
      -- (one it had, of another value: taken out first)
      local pattern = '<item oor:path="' .. vim.pesc(PATH) .. '"><prop oor:name="' .. s[1] .. '"[^>]*>.-</item>%s*'
      text = text:gsub(pattern, "")
      text = text:gsub("</oor:items>", function() return want .. "\n</oor:items>" end, 1)
      changed = true
    end
  end
  if changed then
    vim.fn.mkdir(dir .. "/user", "p")
    local out = assert(io.open(file, "wb"))
    out:write(text)
    out:close()
  end
  return "-env:UserInstallation=file://" .. dir
end

return M
