-- Page settings: the document's, at its top ("page: {size = A4, margins = 2.5cm}"), and a section's
-- where it starts, saying what changes ("\newpage {orientation = landscape}"), as format lines do.
-- A section (Word's and LibreOffice's) starts on a new page with its own size, orientation and
-- margins; deleting its line joins it to the one before.
--
-- A section's settings, as the formats give them: props = { width, height, top, bottom, left,
-- right (cm), start = "page" | "continuous" | "odd" | "even" }; as shown: attrs = { size,
-- orientation, ["margin-top"]... , start } (values as written).
local markup = require("neoffice.markup")

local M = {}

-- (short side, long side: cm)
M.SIZES = { { "A3", 29.7, 42 }, { "A4", 21, 29.7 }, { "A5", 14.8, 21 }, { "B5", 17.6, 25 }, { "Letter", 21.59, 27.94 }, { "Legal", 21.59, 35.56 } }
M.SIDES = { "top", "bottom", "left", "right" }
M.KEYS = { "size", "orientation", "margin-top", "margin-bottom", "margin-left", "margin-right", "start" }
M.VALUES = {
  orientation = { "portrait", "landscape" },
  start = { "page", "continuous", "odd", "even" },
  size = vim.tbl_map(function(s) return s[1] end, M.SIZES),
}
M.DESCRIPTIONS = {
  size = "the paper: A4, Letter... or its sides, 21cm x 29.7cm",
  orientation = "portrait or landscape",
  margins = "the margins: one length for all four (2.5cm), or two: top and bottom, left and right (2.5cm 3cm)",
  ["margin-top"] = "the top margin",
  ["margin-bottom"] = "the bottom margin",
  ["margin-left"] = "the left margin",
  ["margin-right"] = "the right margin",
  start = "where the section starts: a new page (page), an odd or even page, or on the same one (continuous; Word only)",
}
local NAMES = { "size", "orientation", "margins", "margin-top", "margin-bottom", "margin-left", "margin-right", "start" }

-- A length as shown: cm, to a tenth of a millimetre ("2.54cm", "2cm").
function M.length(cm)
  local s = ("%.2f"):format(cm):gsub("0+$", ""):gsub("%.$", "")
  return (s == "-0" and "0" or s) .. "cm"
end

local function size_name(short, long)
  for _, s in ipairs(M.SIZES) do
    if math.abs(short - s[2]) < 0.05 and math.abs(long - s[3]) < 0.05 then
      return s[1]
    end
  end
end

-- The sides of a size (short, long: cm), as written: "A4" or "21cm x 29.7cm".
function M.sides(size)
  for _, s in ipairs(M.SIZES) do
    if s[1]:lower() == (size or ""):lower() then
      return s[2], s[3]
    end
  end
  local a, b = (size or ""):match("^%s*(%S+)%s*[x×]%s*(%S+)%s*$")
  a, b = a and markup.cm(a), b and markup.cm(b)
  if a and b and a > 0 and b > 0 then
    return math.min(a, b), math.max(a, b)
  end
end

-- What settings show: every attribute.
function M.view(props)
  local w, h = props.width or 21, props.height or 29.7
  local short, long = math.min(w, h), math.max(w, h)
  local out = {
    size = size_name(short, long) or (M.length(short) .. " x " .. M.length(long)),
    orientation = w > h + 0.01 and "landscape" or "portrait",
    start = props.start or "page",
  }
  for _, side in ipairs(M.SIDES) do
    out["margin-" .. side] = M.length(props[side] or 2.5)
  end
  return out
end

-- The attributes in braces ("orientation = landscape, margins = 2cm"): { name = value } as shown
-- (margins: each side), or nil and why.
function M.parse(text)
  local out = {}
  for _, a in ipairs((markup.attributes(text))) do
    local name, value = a.name, vim.trim(a.value)
    local lower = value:lower()
    if not name then -- (a value alone: whose it is)
      name = vim.tbl_contains(M.VALUES.orientation, lower) and "orientation" or vim.tbl_contains(M.VALUES.start, lower) and "start"
        or M.sides(value) and "size" or nil
    end
    if not name or not vim.tbl_contains(NAMES, name) then
      return nil, ("not a page's attribute: %s (they are %s)"):format(a.name or a.value, table.concat(NAMES, ", "))
    end
    if name == "size" then
      local short, long = M.sides(value)
      if not short then
        return nil, ("not a value of size: %s (%s, or its sides: 21cm x 29.7cm)"):format(value, table.concat(M.VALUES.size, ", "))
      end
      out.size = size_name(short, long) or (M.length(short) .. " x " .. M.length(long))
    elseif M.VALUES[name] then
      if not vim.tbl_contains(M.VALUES[name], lower) then
        return nil, ("not a value of %s: %s (%s)"):format(name, value, table.concat(M.VALUES[name], ", "))
      end
      out[name] = lower
    else -- (margins: one length for all four, or two: top and bottom, left and right)
      local a, b = value:match("^(%S+)%s+(%S+)$")
      local cm, cm2 = markup.cm(a or value), b and markup.cm(b)
      if not cm or cm < 0 or b and (name ~= "margins" or not cm2 or cm2 < 0) then
        return nil, ("not a value of %s: %s (a length: 2.5cm, 25mm, 1in%s)"):format(name, value, name == "margins" and "; or two: top and bottom, left and right" or "")
      end
      if name == "margins" then
        out["margin-top"], out["margin-bottom"] = M.length(cm), M.length(cm)
        out["margin-left"], out["margin-right"] = M.length(cm2 or cm), M.length(cm2 or cm)
      else
        out[name] = M.length(cm)
      end
    end
  end
  return out
end

-- How attributes are written, in M.KEYS's order (the four margins as one when the same). For the
-- document's (`whole`): its size and margins always, its orientation when landscape.
-- (the margins as one: "margins = 2.5cm" for all four, "margins = 2.5cm 3cm" for top and bottom,
-- left and right, when they are so)
function M.render(attrs, whole)
  local parts = {}
  local v_, h_ = attrs["margin-top"], attrs["margin-left"]
  local paired = v_ and h_ and attrs["margin-bottom"] == v_ and attrs["margin-right"] == h_
  for _, key in ipairs(M.KEYS) do
    local v = attrs[key]
    if v and not (whole and (key == "start" or key == "orientation" and v == "portrait")) then
      if key:match("^margin%-") and paired then
        if key == "margin-top" then
          parts[#parts + 1] = "margins = " .. v_ .. (h_ ~= v_ and " " .. h_ or "")
        end
      else
        parts[#parts + 1] = key .. " = " .. v
      end
    end
  end
  return table.concat(parts, ", ")
end

-- What changes from `before` to `attrs` (what a section's line says), and what that gives.
function M.delta(before, attrs)
  local out = {}
  for _, key in ipairs(M.KEYS) do
    if attrs[key] ~= before[key] then
      out[key] = attrs[key]
    end
  end
  return out
end
function M.apply(before, delta)
  return vim.tbl_extend("force", before, delta)
end

-- The page's sides (cm), as attributes say: width, height.
function M.dims(attrs)
  local short, long = M.sides(attrs.size)
  if attrs.orientation == "landscape" then
    return long, short
  end
  return short, long
end

-- What to write for settings that show `have` and are to show `want`: { size = true (width,
-- height: cm) when the size or orientation changed, margins = { side = cm } of those that did,
-- start }; nil if nothing.
function M.changes(have, want)
  local out, any = { margins = {} }, false
  if have.size ~= want.size or have.orientation ~= want.orientation then
    out.width, out.height = M.dims(want)
    out.landscape, out.size, out.name, any = want.orientation == "landscape", true, want.size, true
  end
  for _, side in ipairs(M.SIDES) do
    local key = "margin-" .. side
    if have[key] ~= want[key] then
      out.margins[side], any = markup.cm(want[key]), true
    end
  end
  if have.start ~= want.start then
    out.start, any = want.start, true
  end
  return any and out or nil
end

return M
