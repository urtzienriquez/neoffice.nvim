-- Paragraphs' formats: their alignment (a mark: "<=> ", see markup.lua) and these attributes, in
-- braces after it ("<=> {indent-left = 1cm, first = 0.5cm}"):
--   indent-left, indent-right   indentation (cm, mm, in, pt)
--   first                       the first line's indent (negative: a hanging indent)
--   before, after               the space before and after the paragraph (pt...)
--   line                        line spacing: single, 1.5, double (or another multiple), or a
--                               length (exactly that)
-- `none`: no value of its own (the style's). A format is { align, attrs = { name = value } }: the
-- direct formatting of a paragraph (what Word and LibreOffice show for it, rather than its style's).
--
-- A line with only a mark and its braces sets the format of the plain paragraphs after it (not
-- headings, lists, captions, quotes, table cells or boxes), until the next such line, as in LaTeX;
-- a paragraph that differs says how (its own mark, and the attributes that differ).
local markup = require("neoffice.markup")
local M = {}

M.KEYS = { "indent-left", "indent-right", "first", "before", "after", "line" } -- (in the order written)
M.DESCRIBE = {
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
function M.measure(key, value)
  if not value or value == "none" then
    return nil
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
    local name, value = a.name, a.value:lower()
    if not name or not vim.tbl_contains(M.KEYS, name) then
      return nil, ("not a paragraph's attribute: %s (they are %s)"):format(a.name or a.value, table.concat(M.KEYS, ", "))
    end
    if value ~= "none" and not M.measure(name, value) then
      return nil, ("not a value of %s: %s (%s, or none)"):format(name, a.value,
        name == "line" and "single, 1.5, double, a multiple or a length" or "a length: 1cm, 5mm, 6pt" .. (name == "first" and ", negative for a hanging indent" or ""))
    end
    out[name] = value
  end
  return out
end

function M.render(attrs)
  local parts = {}
  for _, key in ipairs(M.KEYS) do
    if attrs and attrs[key] then
      parts[#parts + 1] = key .. " = " .. attrs[key]
    end
  end
  return table.concat(parts, ", ")
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
  if own_align then
    out.align = own_align ~= "left" and own_align or nil
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
  local align = (format.align or false) ~= (run.align or false) and (format.align or "left") or nil
  if any and not align then
    align = format.align or "left" -- (attributes come after a mark)
  end
  return align, any and attrs or nil
end

return M
