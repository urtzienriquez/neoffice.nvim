-- Paragraphs' formats: their alignment (a mark: "<=> ", see markup.lua) and these attributes, in
-- braces after it ("<=> {indent-left = 1cm, first = 0.5cm}"):
--   indent-left, indent-right   indentation (cm, mm, in, pt)
--   first                       the first line's indent (negative: a hanging indent)
--   before, after               the space before and after the paragraph (pt...)
--   line                        line spacing: single, 1.5, double (or another multiple), or a
--                               length (exactly that)
--   font, size, color           the font of its text, its size (pt) and its colour (a name, or
--                               #rrggbb): what most of its text has (a few words in another:
--                               "[words]{font = Arial}", markup.lua)
-- `none`: no value of its own (the style's). A format is { align, attrs = { name = value } }: the
-- direct formatting of a paragraph (what Word and LibreOffice show for it, rather than its style's).
--
-- A line with only a mark and its braces sets the format of the plain paragraphs after it (not
-- headings, lists, captions, quotes, table cells or boxes), until the next such line, as in LaTeX;
-- a paragraph that differs says how (its own mark, and the attributes that differ).
local markup = require("neoffice.markup")
local M = {}

M.KEYS = { "align", "font", "size", "color", "indent-left", "indent-right", "first", "before", "after", "line" } -- (in the order written)
M.ALIGNS = { "left", "center", "right", "justify" }
-- A style's (in the document's header, see init.lua): these, and bold and italic (said alone).
M.STYLE_KEYS = { "font", "size", "color", "align", "bold", "italic", "indent-left", "indent-right", "first", "before", "after", "line" }
M.LAYOUT = { "indent-left", "indent-right", "first", "before", "after", "line" } -- (the paragraph's own: not its text's)
M.DESCRIBE = {
  align = "its alignment: left, center, right, justify (none: its style's)",
  bold = "its text in bold", italic = "its text in italics",
  font = "the font of its text (Times New Roman, Arial...)",
  size = "the size of its text (11pt, 12pt...)",
  color = "the colour of its text (red, darkblue... or #rrggbb)",
  ["indent-left"] = "indentation on the left (cm, mm, in, pt)",
  ["indent-right"] = "indentation on the right",
  first = "the first line's indent (negative: a hanging indent)",
  before = "space before the paragraph (pt...)",
  after = "space after the paragraph",
  line = "line spacing: single, 1.5, double, a multiple, or a length (exactly)",
}

local function length(value) -- (cm, from "1cm", "-0.5cm", "6pt"...)
  local sign, rest = value:match("^(%-?)(.+)$")
  local cm = markup.cm(rest)
  return cm and (sign == "-" and -cm or cm)
end

-- A value's measure, for comparing: a number (cm for lengths) or, for line spacing, "x1.5" (a
-- multiple) or a number (exact, in cm); nil for none.
local measured = {} -- (each value's, once: they are few, and asked for often)
function M.measure(key, value)
  if not value or value == "none" then
    return nil
  end
  local known = measured[key]
  if not known then
    known = {}
    measured[key] = known
  end
  local m = known[value]
  if m == nil then
    m = M.measure_of(key, value) or false
    known[value] = m
  end
  return m or nil
end
function M.measure_of(key, value)
  if key == "font" then -- (its name: whatever its case)
    return vim.trim(value):lower()
  elseif key == "color" then -- (its hex value: black is Word's "automatic" too, and a grey that looks it,
    -- as text pasted from web pages has: #222222)
    local hex = markup.color_hex(value)
    if hex and tonumber(hex:sub(1, 2), 16) <= 0x30 and tonumber(hex:sub(3, 4), 16) <= 0x30 and tonumber(hex:sub(5, 6), 16) <= 0x30 then
      return "000000"
    end
    return hex
  elseif key == "align" then
    return vim.tbl_contains(M.ALIGNS, value) and value or nil
  elseif key == "size" then -- (in pt, to a tenth)
    local cm = length(value)
    return cm and cm > 0 and ("%.1f"):format(cm * 72 / 2.54)
  elseif key == "line" then
    local multiple = ({ single = 1, double = 2 })[value] or tonumber(value)
    if multiple then
      return ("x%.2f"):format(multiple)
    end
    local cm = length(value)
    return cm and ("%.3f"):format(cm)
  end
  local cm = length(value)
  return cm and ("%.3f"):format(cm)
end

-- The attributes in braces: { name = value }, or nil and why.
function M.parse(text)
  local out = {}
  for _, a in ipairs((markup.attributes(text))) do
    local name = a.name == "colour" and "color" or a.name
    local value = name == "font" and vim.trim(a.value) or a.value:lower()
    if not name or not vim.tbl_contains(M.KEYS, name) then
      return nil, ("not a paragraph's attribute: %s (they are %s)"):format(a.name or a.value, table.concat(M.KEYS, ", "))
    end
    if value ~= "none" and not M.measure(name, value) then
      return nil, ("not a value of %s: %s (%s, or none)"):format(name, a.value,
        name == "line" and "single, 1.5, double, a multiple or a length" or name == "size" and "a size: 11pt, 12pt"
          or name == "color" and "a colour: red, darkblue... or #rrggbb"
          or name == "font" and "a font's name" or name == "align" and table.concat(M.ALIGNS, ", ")
          or "a length: 1cm, 5mm, 6pt" .. (name == "first" and ", negative for a hanging indent" or ""))
    end
    out[name] = name == "size" and value ~= "none" and M.size(value) or name == "color" and value ~= "none" and markup.text_color(value) or value
  end
  return out
end

function M.render(attrs)
  local parts = {}
  for _, key in ipairs(M.KEYS) do
    if attrs and attrs[key] then
      parts[#parts + 1] = key .. " = " .. markup.attr_value(attrs[key])
    end
  end
  return table.concat(parts, ", ")
end

-- A style's attributes, in the document's header ("font = Arial, size = 14pt, bold"): { name =
-- value, bold = true... }, or nil and why. (No none: a style says what it gives.)
function M.parse_style(text)
  local out = {}
  for _, a in ipairs((markup.attributes(text))) do
    local flag = not a.name and (a.value:lower() == "bold" or a.value:lower() == "italic") and a.value:lower()
    if flag then
      out[flag] = true
    elseif a.name == "bold" or a.name == "italic" then
      local v = a.value:lower()
      local on = ({ yes = true, ["true"] = true, on = true, no = false, ["false"] = false, off = false })[v]
      if on == nil then
        return nil, ("not a value of %s: %s (yes or no; or %s alone)"):format(a.name, a.value, a.name)
      end
      out[a.name] = on
    else
      a.name = a.name == "colour" and "color" or a.name
      if not a.name or not vim.tbl_contains(M.STYLE_KEYS, a.name) then
        return nil, ("not a style's attribute: %s (they are %s)"):format(a.name or a.value, table.concat(M.STYLE_KEYS, ", "))
      end
      local one, bad = M.parse(a.name .. " = " .. markup.attr_value(vim.trim(a.value)))
      if not one or one[a.name] == "none" then
        return nil, bad or ("a style says what it gives: %s = none isn't one (leave it out for its default)"):format(a.name)
      end
      out[a.name] = one[a.name]
    end
  end
  return out
end
function M.render_style(attrs)
  local parts = {}
  for _, key in ipairs(M.STYLE_KEYS) do
    local v = attrs[key]
    if v == true then
      parts[#parts + 1] = key
    elseif v and v ~= false then
      parts[#parts + 1] = key .. " = " .. markup.attr_value(v)
    end
  end
  return table.concat(parts, ", ")
end
-- Whether two styles' values are the same (as shown: a default, or false, as none).
function M.same_style_value(key, a, b)
  if key == "bold" or key == "italic" then
    return (a or false) == (b or false)
  end
  local default = key == "line" and "single" or (key == "align" and "left") or (key == "color" and "black") or (vim.tbl_contains(M.LAYOUT, key) and "0cm") or nil
  return M.measure(key, a or default) == M.measure(key, b or default)
end

-- A size as shown: in pt ("11pt", "10.5pt"), from any length ("11pt", "0.4cm").
function M.size(value)
  local cm = length(value)
  return cm and M.pt(cm * 72 / 2.54)
end

-- As shown: lengths in cm (to a tenth of a millimetre), spaces in pt, line spacing as a multiple
-- (single, double) or pt.
function M.cm(n)
  return (("%.2f"):format(n):gsub("%.?0+$", "")) .. "cm"
end
function M.pt(n)
  return (("%.1f"):format(n):gsub("%.?0+$", "")) .. "pt"
end
function M.multiple(n)
  if math.abs(n - 1) < 0.01 then
    return "single"
  elseif math.abs(n - 2) < 0.01 then
    return "double"
  end
  return (("%.2f"):format(n):gsub("%.?0+$", ""))
end

-- Whether two formats are the same (alignment and attributes, as measured).
local function same_attrs(a, b)
  for _, key in ipairs(M.KEYS) do
    if M.measure(key, a[key]) ~= M.measure(key, b[key]) then
      return false
    end
  end
  return true
end
function M.same(a, b)
  return (a.align or false) == (b.align or false) and same_attrs(a.attrs or {}, b.attrs or {})
end
function M.key(format)
  local parts = { format.align or "" }
  for _, key in ipairs(M.KEYS) do
    parts[#parts + 1] = M.measure(key, (format.attrs or {})[key]) or ""
  end
  return table.concat(parts, "|")
end

-- A paragraph's format in a run's: the run's, with what the paragraph says (its own mark: align,
-- "left" for none; its own attributes: none for none).
function M.apply(run, own_align, own_attrs)
  local out = { align = run.align, attrs = {} }
  for key, value in pairs(run.attrs or {}) do
    out.attrs[key] = value
  end
  if own_align then -- (none: its style's; left: left, when its style isn't)
    out.align = own_align ~= "none" and own_align or nil
  end
  for key, value in pairs(own_attrs or {}) do
    out.attrs[key] = value ~= "none" and value or nil
  end
  return out
end

-- What a paragraph must say to have `format` in `run`: its mark (or nil: the run's) and
-- attributes (those that differ; none for one the run has and it hasn't). nil, nil: none.
function M.diff(run, format)
  local attrs, any = {}, false
  for _, key in ipairs(M.KEYS) do
    local want, have = (format.attrs or {})[key], (run.attrs or {})[key]
    if M.measure(key, want) ~= M.measure(key, have) then
      attrs[key], any = want or "none", true
    end
  end
  local align = (format.align or false) ~= (run.align or false) and (format.align or "none") or nil
  return align, any and attrs or nil
end

return M
