-- Floating images and boxes: where they go, written as attributes ({width = 6cm, float = right}),
-- not as coordinates. A format reads its own XML into a geometry, which M.classify turns into
-- attributes (what it means: on the right, in the middle...), and writes the geometry M.geometry
-- makes of them.
--
-- Attributes (those not given: the defaults):
--   float  left | right | center   (none: in the line, as a character)
--   wrap   square | tight | none | front | behind   (none: text above and below it; front and
--          behind: over or under the text. Default: none in the middle, else square)
--   place  here | top | bottom | page   (here: at its paragraph; top, bottom: of the page; page:
--          at x, y from the page's corner)
--   gap    the space between it and the text (default 0.3cm)
--   x, y   an exact place: from the column's left edge and the paragraph's top (or, with place =
--          page, the page's corner), when it isn't one of those
-- (and width, as for any image; and crop-top, crop-right, crop-bottom, crop-left: how much of the
-- image is cut off at each side, as a percentage of it: "crop-top = 39%")
--
-- A geometry: { h = { rel = "column" | "page", align = "left" | "center" | "right" | nil,
-- offset = cm | nil }, v = { rel = "paragraph" | "margin" | "page", align = "top" | "bottom" | nil,
-- offset = cm | nil }, wrap, gap (cm), width (cm) }; nil for an image in the line.
-- `page`: { text = the text's width, left = the left margin, top = the top margin } (cm).
local markup = require("neoffice.markup")
local M = {}

M.KEYS = { "width", "crop-top", "crop-right", "crop-bottom", "crop-left", "float", "wrap", "place", "gap", "x", "y" } -- (in the order they are written)
M.CROP = { "crop-top", "crop-right", "crop-bottom", "crop-left" }
M.VALUES = {
  float = { "left", "right", "center" },
  wrap = { "square", "tight", "none", "front", "behind" },
  place = { "here", "top", "bottom", "page" },
}
M.GAP = 0.3
M.DESCRIBE = {
  ["float.left"] = "at the left of the column, the text beside it",
  ["float.right"] = "at the right of the column, the text beside it",
  ["float.center"] = "in the middle of the column, the text above and below it",
  ["wrap.square"] = "the text around its box (the default at the left or right)",
  ["wrap.tight"] = "the text close around it",
  ["wrap.none"] = "the text above and below it, none beside (the default in the middle)",
  ["wrap.front"] = "over the text",
  ["wrap.behind"] = "under the text",
  ["place.here"] = "at its paragraph (the default)",
  ["place.top"] = "at the top of the page",
  ["place.bottom"] = "at the bottom of the page",
  ["place.page"] = "at x, y from the page's corner",
  ["crop"] = "how much of the image is cut off at that side (a percentage of it)",
}

-- An image's crop: { top, right, bottom, left } as fractions (0 when not cut), or nil if it isn't.
function M.crop(img)
  local out, any = {}, false
  for k, key in ipairs(M.CROP) do
    out[k] = tonumber(((img or {})[key] or ""):match("^([%d.]+)%%$") or "") or 0
    out[k] = out[k] / 100
    any = any or out[k] > 0
  end
  return any and out or nil
end
-- (a fraction as written: "39.3%", nil for none)
function M.percent(fraction)
  if not fraction or fraction < 0.005 then
    return nil
  end
  return (("%.1f"):format(fraction * 100):gsub("%.0$", "")) .. "%"
end

local function is_length(v)
  return markup.cm(v or "") ~= nil
end
-- (a place may be left of or above where it is measured from: "x = -0.1cm")
local function signed_cm(v)
  local sign, rest = (v or ""):match("^(%-?)(.*)$")
  local cm = markup.cm(rest)
  return cm and (sign == "-" and -cm or cm)
end

-- The attribute a value alone stands for ("right": float, "12cm": width), or nil.
function M.name_of(value)
  for _, key in ipairs({ "float", "wrap", "place" }) do
    if vim.tbl_contains(M.VALUES[key], value) then
      return key
    end
  end
  if is_length(value) or value:match("^[%d.]+%%$") then
    return "width"
  end
end

-- The attributes in braces after an image ("width = 6cm, float = right"): { name = value }, or nil
-- (and why) if they aren't all an image's.
function M.parse(text)
  local out = {}
  for _, a in ipairs((markup.attributes(text))) do
    local name, value = a.name, a.value:lower()
    name = name or M.name_of(value)
    if not name or not vim.tbl_contains(M.KEYS, name) then
      return nil, ("not an image's attribute: %s (they are %s)"):format(a.name or a.value, table.concat(M.KEYS, ", "))
    end
    local percent = tonumber(value:match("^([%d.]+)%%$") or "")
    local ok = M.VALUES[name] and vim.tbl_contains(M.VALUES[name], value)
      or (name == "width") and (is_length(value) or value:match("^[%d.]+%%$"))
      or name:match("^crop%-") and percent and percent < 100
      or name == "gap" and is_length(value)
      or (name == "x" or name == "y") and signed_cm(value) ~= nil
    if not ok then
      return nil, ("not a value of %s: %s (%s)"):format(name, a.value,
        M.VALUES[name] and table.concat(M.VALUES[name], ", ") or name:match("^crop%-") and "a percentage of the image, under 100%"
          or "a length: 12cm, 80mm, 3in" .. (name == "width" and ", or a percentage" or ""))
    end
    out[name] = value
  end
  return out
end

-- How the attributes are written: those that aren't the defaults, in M.KEYS's order.
function M.render(attrs)
  local parts = {}
  for _, key in ipairs(M.KEYS) do
    local v = attrs[key]
    local default = key == "wrap" and (attrs.float == "center" and "none" or "square") or key == "place" and "here"
    local none = key:match("^crop%-") and (tonumber((v or ""):match("^([%d.]+)%%$") or "") or 0) < 0.05
    if v and v ~= default and not none and not (key == "gap" and markup.cm(v) and math.abs(markup.cm(v) - M.GAP) < 0.005) then
      parts[#parts + 1] = key .. " = " .. v
    end
  end
  return table.concat(parts, ", ")
end

-- What can change of an image, as a string (to tell whether it did): its attributes as they are
-- written (so what isn't shown, as a default, or a gap that shows as the default, is no change).
function M.key(img)
  return M.render(img or {})
end

-- What can change of where it goes (its width and place: not its crop), the same way.
function M.place_key(img)
  local copy = {}
  for key, value in pairs(img or {}) do
    copy[key] = not key:match("^crop%-") and value or nil
  end
  return M.render(copy)
end

-- Whether attributes place it out of the line.
function M.floating(attrs)
  return attrs and (attrs.float or attrs.place and attrs.place ~= "here" or attrs.x or attrs.y) and true or false
end

local function near(a, b, tolerance)
  return a and b and math.abs(a - b) < (tolerance or 0.3)
end
local function cm(n) -- (as written: to a millimetre)
  return markup.width(n)
end

-- What a geometry means, as attributes.
function M.classify(geo, page)
  if not geo then
    return {}
  end
  local out, w = { wrap = geo.wrap }, geo.width or 0
  local h, v = geo.h or {}, geo.v or {}
  if v.rel == "margin" and v.offset then -- (from the top margin: from the page)
    v = { rel = "page", offset = v.offset + page.top }
  end
  if h.rel == "page" or v.rel == "page" then -- (anywhere on the page: where, exactly)
    local x = h.offset or (h.align == "center" and (page.text + 2 * page.left - w) / 2) or (h.align == "right" and page.text + page.left - w) or page.left
    local y = v.offset or page.top
    if v.rel ~= "page" then
      y = nil
    end
    if v.align == "top" or v.align == "bottom" then
      out.place = v.align
      h = { rel = "column", offset = x - page.left }
    else
      out.place, out.x, out.y = "page", cm(x), y and cm(y) or nil
    end
  end
  if not out.x then
    if h.align then
      out.float = h.align
    elseif h.offset then
      local off = h.offset
      out.float = near(off, 0) and "left" or near(off + w, page.text) and "right" or near(off + w / 2, page.text / 2) and "center" or "left"
      if out.float == "left" and not near(off, 0) then
        out.x = cm(off)
      end
    else
      out.float = "left"
    end
  else
    out.float = "left"
  end
  if not out.place then
    if v.align == "top" or v.align == "bottom" then
      out.place = v.align
    elseif v.offset and not near(v.offset, 0, 0.2) then
      out.y = cm(v.offset)
    end
  end
  out.gap = geo.gap and cm(geo.gap) or nil
  if out.gap and near(geo.gap, M.GAP, 0.005) then
    out.gap = nil
  end
  if out.wrap == (out.float == "center" and "none" or "square") then
    out.wrap = nil
  end
  return out
end

-- The geometry for attributes (nil: in the line).
function M.geometry(attrs, page, width)
  if not M.floating(attrs) then
    return nil
  end
  local float = attrs.float or "left"
  local geo = {
    wrap = attrs.wrap or (float == "center" and "none" or "square"),
    gap = attrs.gap and markup.cm(attrs.gap) or M.GAP,
    width = width,
  }
  local x, y = attrs.x and signed_cm(attrs.x), attrs.y and signed_cm(attrs.y)
  if attrs.place == "page" then
    geo.h = { rel = "page", offset = x or (float == "right" and page.left + page.text - (width or 0)) or (float == "center" and page.left + (page.text - (width or 0)) / 2) or page.left }
    geo.v = { rel = "page", offset = y or page.top }
    return geo
  end
  geo.h = x and { rel = "column", offset = x } or { rel = "column", align = float }
  if attrs.place == "top" or attrs.place == "bottom" then
    geo.v = { rel = "margin", align = attrs.place }
  else
    geo.v = { rel = "paragraph", offset = y or 0 }
  end
  return geo
end

-- A box's attributes (as an image's, but a box can't be cropped): { name = value }, or nil and why.
function M.parse_box(text)
  local attrs, bad = M.parse(text)
  for _, key in ipairs(attrs and M.CROP or {}) do
    if attrs[key] then
      return nil, "a box can't be cropped (" .. key .. "): only its images"
    end
  end
  return attrs, bad
end

-- An error if an image's attributes aren't right (as written: see M.parse).
function M.check(text)
  local ok, err = M.parse(text)
  assert(ok, err)
end

return M
