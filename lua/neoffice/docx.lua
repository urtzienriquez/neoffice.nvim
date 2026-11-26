-- .docx: word/document.xml (+ comments, hyperlinks, images, heading styles) <-> paragraph spans.
-- Every part is changed in place, and only when needed: ctx.dirty lists the parts to write.
local xml = require("neoffice.xml")
local markup = require("neoffice.markup")
local zip = require("neoffice.zip")
local images = require("neoffice.images")
local M = {}

M.MAIN = "word/document.xml"
local RELS, TYPES, STYLES = "word/_rels/document.xml.rels", "[Content_Types].xml", "word/styles.xml"
local COMMENTS, EXTENDED = "word/comments.xml", "word/commentsExtended.xml"
local IDS, EXTENSIBLE = "word/commentsIds.xml", "word/commentsExtensible.xml"
local NUMBERING, FOOTNOTES = "word/numbering.xml", "word/footnotes.xml"
local JC = { center = "center", right = "right", ["end"] = "right", both = "justify", distribute = "justify" }

local NS = {
  w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
  w14 = "http://schemas.microsoft.com/office/word/2010/wordml",
  w15 = "http://schemas.microsoft.com/office/word/2012/wordml",
  r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
  wp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
  m = "http://schemas.openxmlformats.org/officeDocument/2006/math",
  mc = "http://schemas.openxmlformats.org/markup-compatibility/2006",
  wps = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
  a = "http://schemas.openxmlformats.org/drawingml/2006/main",
}
local REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
local REL_EX = "http://schemas.microsoft.com/office/2011/relationships/commentsExtended"
local MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml."
local MAX_WIDTH = 5486400 -- new images are at most 6 inches wide (in EMU)

-- Elements shown as a "￼" placeholder (others without text stay invisible but are kept).
local VISIBLE = {
  ["w:drawing"] = true, ["w:pict"] = true, ["w:object"] = true, ["w:br"] = true, ["w:cr"] = true,
  ["w:sym"] = true, ["w:noBreakHyphen"] = true, ["w:footnoteReference"] = true,
  ["w:endnoteReference"] = true, ["w:instrText"] = false,
}
-- Elements dropped from paragraphs you edit, because Word recreates them (or we do).
local DROP = { ["w:proofErr"] = true, ["w:lastRenderedPageBreak"] = true, ["w:commentReference"] = true }
-- Schema order of run properties: Word rejects files where they are out of order.
local RPR_ORDER = {}
for i, tag in ipairs(vim.split("rStyle rFonts b bCs i iCs caps smallCaps strike dstrike outline shadow emboss imprint noProof snapToGrid vanish webHidden color spacing w kern position sz szCs highlight u effect bdr shd fitText vertAlign rtl cs em lang eastAsianLayout specVanish oMath", " ")) do
  RPR_ORDER["w:" .. tag] = i
end

local function is_on(el) -- <w:b/> is on; <w:b w:val="0"/> and <w:u w:val="none"/> are off
  local v = el and xml.attr(el, "w:val")
  return el ~= nil and v ~= "0" and v ~= "false" and v ~= "off" and v ~= "none"
end

-- The formatting of run properties: { b, i, u, s, h, sup, sub } (see markup.FLAGS).
-- (Word's highlight colours: "darkBlue" <-> "darkblue")
local function word_color(name)
  return (name:gsub("^(%l+)(blue)$", "%1Blue"):gsub("^(%l+)(cyan)$", "%1Cyan"):gsub("^(%l+)(green)$", "%1Green")
    :gsub("^(%l+)(magenta)$", "%1Magenta"):gsub("^(%l+)(red)$", "%1Red"):gsub("^(%l+)(yellow)$", "%1Yellow")
    :gsub("^(%l+)(gray)$", "%1Gray"))
end
local function flags(rpr)
  local function get(tag)
    return is_on(rpr and xml.child(rpr, tag))
  end
  local align = rpr and xml.child(rpr, "w:vertAlign")
  align = align and xml.attr(align, "w:val")
  local highlight = rpr and xml.child(rpr, "w:highlight")
  return {
    b = get("w:b"), i = get("w:i"), u = get("w:u"), s = get("w:strike"),
    h = is_on(highlight) and (markup.color(xml.attr(highlight, "w:val")) or "yellow"),
    sup = align == "superscript", sub = align == "subscript",
  }
end

local function visible(el)
  local v = VISIBLE[el.tag]
  if v == nil then
    return xml.text(el) ~= ""
  end
  return v
end

-- A part of the package as XML (parsed once); nil if it doesn't exist, unless `empty` is given to
-- create it with.
local function part(ctx, name, empty)
  if ctx.parts[name] == nil then
    local text = zip.read(ctx.path, name)
    ctx.parts[name] = text and xml.parse(text) or false
  end
  if ctx.parts[name] == false and empty then
    ctx.parts[name] = xml.parse(empty)
  end
  return ctx.parts[name] or nil
end

local function ensure_ns(el, prefix)
  if not xml.attr(el, "xmlns:" .. prefix) then
    xml.set_attr(el, "xmlns:" .. prefix, NS[prefix])
  end
end

local function has(root, tag, key, value)
  return vim.iter(xml.collect(root, { [tag] = true })):find(function(el)
    return xml.attr(el, key) == value
  end)
end

-- ── Reading ─────────────────────────────────────────────────────────────────

local latex_of -- (below)

local function read_comments(ctx)
  local by_id, by_para = {}, {}
  for _, el in ipairs(xml.collect(part(ctx, COMMENTS) or {}, { ["w:comment"] = true })) do
    local paras, lines = xml.collect(el, { ["w:p"] = true }), {}
    for _, p in ipairs(paras) do
      lines[#lines + 1] = xml.text(p)
    end
    local c = {
      id = xml.attr(el, "w:id"),
      node = el,
      para = paras[1] and xml.attr(paras[#paras], "w14:paraId"),
      author = xml.attr(el, "w:author") or "?",
      date = xml.attr(el, "w:date") or "",
      text = table.concat(lines, "\n"),
      replies = {},
    }
    c.was_text = c.text -- (its text as read: written again only if edited)
    by_id[c.id] = c
    ctx.comment_of[el] = c
    ctx.next_comment = math.max(ctx.next_comment, (tonumber(c.id) or 0) + 1)
    if c.para then
      by_para[c.para] = c
    end
  end
  -- Threads and "resolved" live in commentsExtended.xml, keyed by the comment's last paragraph id.
  for _, el in ipairs(xml.collect(part(ctx, EXTENDED) or {}, { ["w15:commentEx"] = true })) do
    local c, parent = by_para[xml.attr(el, "w15:paraId")], by_para[xml.attr(el, "w15:paraIdParent") or ""]
    if c then
      c.ex = el
      if parent then
        c.reply_of = parent
        table.insert(parent.replies, c)
      end
      if xml.attr(el, "w15:done") == "1" then
        (parent or c).resolved = true
      end
    end
  end
  for _, c in pairs(by_id) do
    c.was_resolved = c.resolved
  end
  return by_id
end

-- Whether level `ilvl` of list `num_id` (numbering.xml) has bullets or numbers.
local function abstract_of(ctx, num_id) -- (a list's definition: its w:abstractNum, and its w:num)
  local root = part(ctx, NUMBERING)
  local num = root and has(root, "w:num", "w:numId", num_id)
  local abstract_id = num and xml.attr(xml.child(num, "w:abstractNumId") or {}, "w:val")
  return abstract_id and has(root, "w:abstractNum", "w:abstractNumId", abstract_id), num
end
local function lvl_of(abstract, ilvl)
  return abstract and has(abstract, "w:lvl", "w:ilvl", tostring(ilvl))
end
-- (a level as a list's instance has it: its own override, else its definition's)
local function override_of(num, ilvl)
  for _, o in ipairs(num and xml.collect(num, { ["w:lvlOverride"] = true }) or {}) do
    if xml.attr(o, "w:ilvl") == tostring(ilvl) then
      return o
    end
  end
end
local function eff_lvl(ctx, num_id, ilvl)
  local abstract, num = abstract_of(ctx, num_id)
  local o = override_of(num, ilvl)
  return o and xml.child(o, "w:lvl") or lvl_of(abstract, ilvl)
end
local function list_kind(ctx, num_id, ilvl)
  local lvl = eff_lvl(ctx, num_id, ilvl)
  return lvl and xml.attr(xml.child(lvl, "w:numFmt") or {}, "w:val") == "bullet" and "bullet" or "number"
end
-- A numbered level's pattern (see markup.list_marker): its text ("%1.%2."), each number in its
-- level's format; "1." if it isn't one that can be written so ("(1)", "Step 1:").
local SAMPLE = { decimal = "1", decimalZero = "1", lowerLetter = "a", upperLetter = "A", lowerRoman = "i", upperRoman = "I" }
local FORMAT = { ["1"] = "decimal", a = "lowerLetter", A = "upperLetter", i = "lowerRoman", I = "upperRoman" }
local function list_pattern(ctx, num_id, ilvl)
  local lvl = eff_lvl(ctx, num_id, ilvl)
  local text = lvl and xml.attr(xml.child(lvl, "w:lvlText") or {}, "w:val")
  if not text then
    return "1."
  end
  local ok = true
  local pattern = text:gsub("%%(%d)", function(d)
    local of = eff_lvl(ctx, num_id, tonumber(d) - 1)
    local sample = of and SAMPLE[xml.attr(xml.child(of, "w:numFmt") or {}, "w:val") or ""]
    ok = ok and sample ~= nil
    return sample or "?"
  end)
  local back = ok and markup.list_marker(pattern .. " x")
  return back == pattern and pattern or "1."
end

-- The list a paragraph (its pPr) is an item of, from its own numbering or its style's:
-- { kind = "bullet" | "number", level = 1.., id = numId }, or nil.
local function list_of(ctx, ppr)
  local style = ppr and xml.child(ppr, "w:pStyle")
  local num_pr = ppr and xml.child(ppr, "w:numPr") or style and ctx.style_numbering[xml.attr(style, "w:val")]
  local num_id = num_pr and xml.attr(xml.child(num_pr, "w:numId") or {}, "w:val")
  if not num_id or num_id == "0" then
    return nil
  end
  local ilvl = tonumber(xml.attr(xml.child(num_pr, "w:ilvl") or {}, "w:val") or "") or 0
  local kind = list_kind(ctx, num_id, ilvl)
  local abstract, num = abstract_of(ctx, num_id)
  local restarts = false -- (an instance numbering again from 1: a list of its own, where it starts)
  for _, o in ipairs(num and xml.collect(num, { ["w:lvlOverride"] = true }) or {}) do
    restarts = restarts or (xml.attr(o, "w:ilvl") == "0" and xml.child(o, "w:startOverride") ~= nil)
  end
  return { kind = kind, level = ilvl + 1, id = num_id, pattern = kind == "number" and list_pattern(ctx, num_id, ilvl) or nil,
    abstract = abstract and xml.attr(abstract, "w:abstractNumId"), restarts = restarts }
end

-- Lists' instances (w:num): one list (its w:abstractNum) can have several, which count on from
-- each other (Word and LibreOffice do so), each with its own levels where it overrides them. A
-- level numbered as a pattern says ("1.a.") other than the list's is an instance of it with that
-- level overridden: the sub-items under one item "a), b)", under another "1., 2.", the same list.
local function new_num(ctx, abstract_id, overrides) -- (overrides: w:lvlOverride elements)
  local root = xml.find(part(ctx, NUMBERING), "w:numbering")
  local id = tostring(vim.iter(xml.collect(root, { ["w:num"] = true })):fold(0, function(n, el)
    return math.max(n, tonumber(xml.attr(el, "w:numId")) or 0)
  end) + 1)
  local num = xml.elem("w:num", { { "w:numId", id } }, { xml.elem("w:abstractNumId", { { "w:val", abstract_id } }) })
  for _, o in ipairs(overrides or {}) do
    xml.append(num, o)
  end
  local nums = xml.collect(root, { ["w:num"] = true })
  local last = nums[#nums]
  xml.append(root, num, last and last.parent == root and xml.index(last) + 1 or nil)
  ctx.dirty[NUMBERING] = true
  return id
end
local function lvl_for(ctx, num_id, ilvl, pattern) -- (the level as num_id has it, numbered as pattern says)
  local lvl = xml.copy(eff_lvl(ctx, num_id, ilvl) or xml.elem("w:lvl", { { "w:ilvl", tostring(ilvl) } }))
  local parts = {}
  for part in pattern:gmatch("[%w]+[.)]") do
    parts[#parts + 1] = part
  end
  local text = {}
  for k, part in ipairs(parts) do -- (each part: the number of its level, the last this one's)
    text[k] = "%" .. (ilvl + 1 - #parts + k) .. part:sub(-1)
  end
  local function child(tag)
    return xml.child(lvl, tag) or xml.append(lvl, xml.elem(tag), 2)
  end
  xml.set_attr(child("w:numFmt"), "w:val", FORMAT[parts[#parts]:sub(1, -2)] or "decimal")
  xml.set_attr(child("w:lvlText"), "w:val", table.concat(text))
  return lvl
end
-- The instances the list items just before p have (back to the first item of their list).
local function nums_before(ctx, p)
  local out = {}
  for k = xml.index(p) - 1, 1, -1 do
    local sibling = p.parent.children[k]
    if sibling.tag == "w:p" then
      local list = list_of(ctx, xml.child(sibling, "w:pPr"))
      if not list then
        break
      end
      out[#out + 1] = list.id
    elseif sibling.tag then
      break
    end
  end
  return out
end
-- The instance of base's list for an item of level ilvl numbered as pattern: base itself, one the
-- items before it have, or a new one overriding that level (and keeping base's overrides of the
-- levels under it: a restart there).
local function num_for(ctx, p, base, ilvl, pattern)
  if not pattern or list_pattern(ctx, base, ilvl) == pattern then
    return base
  end
  local abstract, num = abstract_of(ctx, base)
  local id = abstract and xml.attr(abstract, "w:abstractNumId")
  for _, other in ipairs(nums_before(ctx, p)) do
    if xml.attr(select(2, abstract_of(ctx, other)) and xml.child(select(2, abstract_of(ctx, other)), "w:abstractNumId") or {}, "w:val") == id
      and list_pattern(ctx, other, ilvl) == pattern then
      return other
    end
  end
  local overrides = {}
  for _, o in ipairs(num and xml.collect(num, { ["w:lvlOverride"] = true }) or {}) do
    if (tonumber(xml.attr(o, "w:ilvl")) or 0) > ilvl then -- (deeper levels' own: as they were)
      overrides[#overrides + 1] = xml.copy(o)
    elseif tonumber(xml.attr(o, "w:ilvl")) == ilvl and xml.child(o, "w:startOverride") then
      overrides[#overrides + 1] = xml.elem("w:lvlOverride", { { "w:ilvl", tostring(ilvl) } }, { xml.copy(xml.child(o, "w:startOverride")) })
    end
  end
  local mine = vim.iter(overrides):find(function(o) return xml.attr(o, "w:ilvl") == tostring(ilvl) end)
  if not mine then
    mine = xml.elem("w:lvlOverride", { { "w:ilvl", tostring(ilvl) } })
    table.insert(overrides, 1, mine)
  end
  xml.append(mine, lvl_for(ctx, base, ilvl, pattern))
  return new_num(ctx, id, overrides)
end
-- A list's new instance numbering from 1 again (an item after a blank line, from the same list).
local function restart_num(ctx, base)
  local abstract, num = abstract_of(ctx, base)
  local overrides = {}
  for _, o in ipairs(num and xml.collect(num, { ["w:lvlOverride"] = true }) or {}) do
    local lvl = xml.child(o, "w:lvl")
    if lvl then
      overrides[#overrides + 1] = xml.elem("w:lvlOverride", { { "w:ilvl", xml.attr(o, "w:ilvl") } }, { xml.copy(lvl) })
    end
  end
  local first = vim.iter(overrides):find(function(o) return xml.attr(o, "w:ilvl") == "0" end)
  if not first then
    first = xml.elem("w:lvlOverride", { { "w:ilvl", "0" } })
    table.insert(overrides, 1, first)
  end
  xml.append(first, xml.elem("w:startOverride", { { "w:val", "1" } }), 1)
  return new_num(ctx, xml.attr(abstract, "w:abstractNumId"), overrides)
end

-- The block of a paragraph (see markup.prefix), from its properties.
-- Fonts (see fonts_of in init.lua): a run's own font and size (its properties', else its character
-- style's), and what a paragraph's style gives (following the styles it is based on, and the
-- document's defaults). Sizes as shown ("11pt"); a theme's font by its name.
local function theme_font(ctx, which) -- (the theme's latin font: "minor" for the text, "major" for headings)
  if not ctx.theme_fonts then
    ctx.theme_fonts = {}
    local ok, theme = pcall(part, ctx, "word/theme/theme1.xml")
    for _, kind in ipairs(ok and theme and { "minor", "major" } or {}) do
      local f = xml.find(theme, "a:" .. kind .. "Font")
      local latin = f and xml.child(f, "a:latin")
      ctx.theme_fonts[kind] = latin and xml.attr(latin, "typeface") ~= "" and xml.attr(latin, "typeface") or nil
    end
  end
  return ctx.theme_fonts[which]
end
local function rpr_font(ctx, rpr) -- (what run properties say: { font, size, color }, nil for what they don't)
  local out = {}
  local fonts = rpr and xml.child(rpr, "w:rFonts")
  if fonts then
    out.font = xml.attr(fonts, "w:ascii") or xml.attr(fonts, "w:hAnsi")
    local theme = not out.font and (xml.attr(fonts, "w:asciiTheme") or xml.attr(fonts, "w:hAnsiTheme"))
    if theme then
      out.font = theme_font(ctx, theme:match("^major") and "major" or "minor")
    end
  end
  local sz = rpr and xml.child(rpr, "w:sz")
  local half = sz and tonumber(xml.attr(sz, "w:val") or "")
  out.size = half and require("neoffice.paras").pt(half / 2)
  local color = rpr and xml.child(rpr, "w:color") -- ("auto", automatic: black, said; a theme's colour: its value)
  local val = color and xml.attr(color, "w:val") or nil
  out.color = val and (val == "auto" and "black" or markup.text_color(val)) or nil
  return out
end
local function style_font(ctx, id, paragraph) -- (a style's, following its chain; a paragraph's, from the defaults)
  ctx.style_fonts = ctx.style_fonts or {}
  local key = (paragraph and "p|" or "c|") .. (id or "")
  if not ctx.style_fonts[key] then
    local styles = part(ctx, STYLES) or {}
    local chain, cur = {}, id
    if not cur and paragraph then -- (the default paragraph style)
      local default = vim.iter(xml.collect(styles, { ["w:style"] = true })):find(function(s)
        return xml.attr(s, "w:type") == "paragraph" and xml.attr(s, "w:default") == "1"
      end)
      cur = default and xml.attr(default, "w:styleId")
    end
    for _ = 1, 10 do
      if not cur then
        break
      end
      chain[#chain + 1] = cur
      cur = ctx.based_on[cur]
    end
    local out = {}
    if paragraph then
      local defaults = xml.find(styles, "w:rPrDefault")
      out = rpr_font(ctx, defaults and xml.child(defaults, "w:rPr"))
    end
    for k = #chain, 1, -1 do
      local style = has(styles, "w:style", "w:styleId", chain[k])
      for name, value in pairs(rpr_font(ctx, style and xml.child(style, "w:rPr"))) do
        out[name] = value
      end
    end
    ctx.style_fonts[key] = out
  end
  return ctx.style_fonts[key]
end
local function own_font(ctx, rpr)
  local own = rpr_font(ctx, rpr)
  local rstyle = rpr and xml.child(rpr, "w:rStyle")
  if rstyle then
    local s = style_font(ctx, xml.attr(rstyle, "w:val"))
    own.font, own.size, own.color = own.font or s.font, own.size or s.size, own.color or s.color
  end
  return own
end
function M.run_font(ctx, r)
  if type(r) ~= "table" or r.tag ~= "w:r" then
    return {}
  end
  ctx.run_fonts = ctx.run_fonts or setmetatable({}, { __mode = "k" }) -- (a run read stays as it was: those written are new)
  if not ctx.run_fonts[r] then
    ctx.run_fonts[r] = own_font(ctx, xml.child(r, "w:rPr"))
  end
  return ctx.run_fonts[r] -- (as it is: not to be changed)
end
function M.font_default(ctx) -- (the document's: its default paragraph style's, else Normal's, as LibreOffice writes it)
  local styles = part(ctx, STYLES) or {}
  local flagged = vim.iter(xml.collect(styles, { ["w:style"] = true })):find(function(st)
    return xml.attr(st, "w:type") == "paragraph" and xml.attr(st, "w:default") == "1"
  end)
  if not flagged and has(styles, "w:style", "w:styleId", "Normal") then
    local base = style_font(ctx, "Normal", true)
    return { font = base.font, size = base.size or "10pt", color = base.color or "black" }
  end
  return M.font_base(ctx, nil)
end
function M.font_base(ctx, ppr)
  local style = ppr and xml.child(ppr, "w:pStyle")
  local base = style_font(ctx, style and xml.attr(style, "w:val"), true) -- (cached, by style)
  return { font = base.font, size = base.size or "10pt", color = base.color or "black" } -- (Word's size when nothing says one; its colour: automatic, black)
end

local function block_of(ctx, ppr)
  local style = ppr and xml.child(ppr, "w:pStyle")
  local id = style and xml.attr(style, "w:val")
  local jc = ppr and xml.child(ppr, "w:jc")
  local block = {}
  if id and ctx.level_of[id] then
    block = { heading = ctx.level_of[id] }
  elseif id and ctx.title_styles[id] then
    block = { title = true }
  elseif id and ctx.quote_styles[id] then
    block = { quote = true }
  else
    local list = list_of(ctx, ppr)
    block = list and { list = list } or {}
  end
  block.align = M.own_align(ctx, ppr)
  block.caption_style = id and ctx.caption_styles[id] or nil -- (see captions in init.lua)
  block.font_base, block.font_default = M.font_base(ctx, ppr), M.font_default(ctx)
  if not (block.heading or block.title or block.quote or block.list) then -- (a plain one: its own format, see paras.lua)
    block.para = M.para_attrs(ppr, ctx)
  end
  return block
end

-- A paragraph's own format's attributes (see paras.lua), from its properties: those its style
-- doesn't give it already (a value the same as its style's is none of its own).
local function style_attrs(ctx, id) -- (its style's, following the styles it is based on, and the defaults)
  ctx.style_attrs = ctx.style_attrs or {}
  local key = id or false
  if not ctx.style_attrs[key] then
    local styles = part(ctx, STYLES) or {}
    local chain, cur = {}, id
    if not cur then -- (the default paragraph style)
      local default = vim.iter(xml.collect(styles, { ["w:style"] = true })):find(function(s)
        return xml.attr(s, "w:type") == "paragraph" and xml.attr(s, "w:default") == "1"
      end)
      cur = default and xml.attr(default, "w:styleId")
    end
    for _ = 1, 10 do
      if not cur then
        break
      end
      chain[#chain + 1] = cur
      cur = ctx.based_on[cur]
    end
    local defaults = xml.find(styles, "w:pPrDefault")
    local out = M.para_attrs(defaults and xml.child(defaults, "w:pPr"))
    for k = #chain, 1, -1 do
      local style = has(styles, "w:style", "w:styleId", chain[k])
      for name, value in pairs(M.para_attrs(style and xml.child(style, "w:pPr"))) do
        out[name] = value
      end
    end
    ctx.style_attrs[key] = out
  end
  return ctx.style_attrs[key]
end
function M.para_attrs(ppr, ctx)
  local out = M.raw_para_attrs(ppr)
  if ctx then
    local paras = require("neoffice.paras")
    local id = ppr and xml.child(ppr, "w:pStyle") and xml.attr(xml.child(ppr, "w:pStyle"), "w:val")
    local style = style_attrs(ctx, id)
    for _, key in ipairs(paras.LAYOUT) do
      local base = style[key] or (key == "line" and "single" or "0cm")
      if out[key] and paras.measure(key, out[key]) == paras.measure(key, base) then
        out[key] = nil
      end
    end
  end
  return out
end
function M.raw_para_attrs(ppr)
  local paras, out = require("neoffice.paras"), {}
  local ind = ppr and xml.child(ppr, "w:ind")
  local function twips(el, ...)
    for _, name in ipairs({ ... }) do
      local v = el and tonumber(xml.attr(el, "w:" .. name) or "")
      if v then
        return v
      end
    end
  end
  local left, right = twips(ind, "left", "start"), twips(ind, "right", "end")
  local first, hanging = twips(ind, "firstLine"), twips(ind, "hanging")
  out["indent-left"] = left and left ~= 0 and paras.cm(left * 2.54 / 1440) or nil
  out["indent-right"] = right and right ~= 0 and paras.cm(right * 2.54 / 1440) or nil
  out.first = hanging and hanging ~= 0 and paras.cm(-hanging * 2.54 / 1440) or first and first ~= 0 and paras.cm(first * 2.54 / 1440) or nil
  local spacing = ppr and xml.child(ppr, "w:spacing")
  local before, after, line = twips(spacing, "before"), twips(spacing, "after"), twips(spacing, "line")
  out.before = before and paras.pt(before / 20) or nil
  out.after = after and paras.pt(after / 20) or nil
  if line then
    local rule = xml.attr(spacing, "w:lineRule") or "auto"
    out.line = rule == "auto" and paras.multiple(line / 240) or paras.pt(line / 20)
  end
  return out
end


-- The sequence a numbering field counts ("SEQ Table \* ARABIC": Table), if it is one that just
-- counts (not one that repeats, resets or hides the number).
function M.seq_name(instr)
  local name, rest = (instr or ""):match("^%s*SEQ%s+([^%s\\]+)(.*)$")
  return name and not rest:find("\\[rcsh]") and name or nil
end

-- A footnote's text (one line: its paragraphs joined by spaces).
local function footnote_text(ctx, id)
  local note = has(part(ctx, FOOTNOTES) or {}, "w:footnote", "w:id", id)
  local paras = note and xml.collect(note, { ["w:p"] = true }) or {}
  return (table.concat(vim.tbl_map(xml.text, paras), " "):gsub("^%s+", ""))
end

-- The run that ends the field (fldChar begin ... end) starting at children[k], if in this paragraph.
local function field_end(children, k)
  local depth = 0
  for j = k, #children do
    for _, char in ipairs(xml.collect(children[j], { ["w:fldChar"] = true })) do
      local kind = xml.attr(char, "w:fldCharType")
      depth = depth + (kind == "begin" and 1 or kind == "end" and -1 or 0)
      if depth == 0 then
        return j
      end
    end
    if children[j].tag ~= "w:r" then
      return nil
    end
  end
end

local function read_para(ctx, p)
  local spans, block = {}, block_of(ctx, nil) -- (one without properties: the default style's, as Word reads it)
  local function add(span)
    spans[#spans + 1] = span
  end
  local function comment_marker(el, key)
    local c = ctx.comments[xml.attr(el, "w:id")]
    if c and not c.reply_of then
      c.seen = true
      add({ [key] = c })
    end
  end

  -- a tracked formatting change: rejecting it puts back the run properties from before
  local function format_change(r, rpr, el)
    local old = xml.child(el, "w:rPr") or xml.elem("w:rPr")
    local old_flags = flags(old)
    return {
      type = "format", author = xml.attr(el, "w:author") or "?", date = xml.attr(el, "w:date") or "",
      old_flags = old_flags, -- (the formatting from before: see read.lua)
      node = el,
      accept = function()
        xml.remove(el)
      end,
      reject = function()
        rpr.children = {}
        for _, c in ipairs(xml.copy(old).children) do
          xml.append(rpr, c)
        end
        return old_flags, r
      end,
    }
  end

  local function image_of(el) -- an embedded picture: its file in the package, alt text, place
    local blip, doc_pr = xml.find(el, "a:blip"), xml.find(el, "wp:docPr")
    local target = blip and ctx.images[xml.attr(blip, "r:embed") or ""]
    local cx = tonumber(xml.attr(xml.find(el, "wp:extent") or {}, "cx") or "")
    local image = target and {
      path = "word/" .. target, alt = doc_pr and vim.trim((xml.attr(doc_pr, "descr") or ""):gsub("%s+", " ")) or "",
      width = cx and markup.width(cx / 360000) or nil,
    }
    if image then
      local floats = require("neoffice.floats")
      local place = floats.classify(M.drawing_geometry(el), M.page(ctx))
      for key, value in pairs(place) do
        image[key] = value
      end
      -- its crop (thousandths of a percent cut off; below 0, space added: not shown, kept)
      local rect = xml.find(el, "a:srcRect")
      for k, side in ipairs({ "t", "r", "b", "l" }) do
        image[floats.CROP[k]] = floats.percent((tonumber(rect and xml.attr(rect, side) or "") or 0) / 100000)
      end
    end
    return image
  end

  -- the run itself stands for its properties (rPr) and attributes
  local function run(r, link)
    local rpr = xml.child(r, "w:rPr")
    local f = flags(rpr)
    local own = own_font(ctx, rpr)
    f.run_font, f.run_size, f.run_color = own.font, own.size, own.color
    local change = rpr and xml.child(rpr, "w:rPrChange")
    change = change and format_change(r, rpr, change)
    if change then
      add({ insert_start = change })
    end
    for _, el in ipairs(r.children) do
      local br = (el.tag == "w:br" or el.tag == "w:cr") and (xml.attr(el, "w:type") or "textWrapping")
      if el.tag == "w:t" or el.tag == "w:delText" or el.tag == "w:tab" then -- (deleted text too: see w:del)
        local text = el.tag ~= "w:tab" and xml.text(el) or "\t"
        if el.tag ~= "w:tab" and xml.attr(el, "xml:space") ~= "preserve" then -- (as Word reads it)
          text = vim.trim(text)
        end
        add(vim.tbl_extend("force", { text = text, props = r, link = link }, f))
      elseif br == "textWrapping" or br == "page" then
        add({ [br == "page" and "page_break" or "line_break"] = true })
      elseif el.tag == "w:commentReference" then -- a comment on a single point
        local c = ctx.comments[xml.attr(el, "w:id")]
        if c and not c.seen and not c.reply_of then
          comment_marker(el, "comment_start")
          add({ comment_end = c })
        end
      elseif (el.tag == "mc:AlternateContent" or el.tag == "w:drawing") and M.box_content(el) then
        -- a box (a text box: its paragraphs show before this one's: see M.load)
        add({ object = r, visible = true, box = { node = r, content = M.box_content(el), attrs = M.box_attrs(ctx, r) } })
      elseif el.tag and el.tag ~= "w:rPr" and not DROP[el.tag] then
        local object = xml.elem("w:r", vim.deepcopy(r.attrs), { rpr and xml.copy(rpr), xml.copy(el) })
        local note = el.tag == "w:footnoteReference" and xml.attr(el, "w:id")
        if note then
          ctx.footnote_ids[note] = true
        end
        local image = el.tag == "w:drawing" and image_of(el) or nil
        local kind = not image and not note and M.object_kind(el) -- (shown by name: ⟦chart⟧, kept)
        add({
          object = object, visible = visible(el) or kind and true, image = image,
          footnote = note and { id = note, text = footnote_text(ctx, note) } or nil,
          field = kind and { text = kind } or nil,
        })
      end
    end
    if change then
      add({ insert_end = change })
    end
  end

  local deleting = false -- (inside a deletion: a deletion in it, as in a move, is part of it)
  local function walk(parent, link)
    local k = 0
    while k < #parent.children do
      k = k + 1
      local el = parent.children[k]
      local tag = el.tag
      local field = tag == "w:r" and xml.find(el, "w:fldChar") and xml.attr(xml.find(el, "w:fldChar"), "w:fldCharType") == "begin"
      local stop = field and field_end(parent.children, k)
      if stop then -- a field (citation, page number, reference...): one thing, showing its result
        local group = { children = {} }
        for j = k, stop do
          xml.append(group, xml.copy(parent.children[j]))
        end
        local result = table.concat(vim.tbl_map(xml.text, xml.collect(group, { ["w:t"] = true })))
        local instr = table.concat(vim.tbl_map(xml.text, xml.collect(group, { ["w:instrText"] = true })))
        add({ object = group, visible = true, field = { text = result, seq = M.seq_name(instr) } })
        k = stop
      elseif tag == "w:pPr" then
        block = block_of(ctx, el)
        local mark = xml.child(xml.child(el, "w:rPr") or {}, "w:ins") -- an inserted paragraph break
        if mark then
          ctx.inserted_break = { type = "insert", node = mark, author = xml.attr(mark, "w:author") or "?", date = xml.attr(mark, "w:date") or "" }
        end
        local gone = xml.child(xml.child(el, "w:rPr") or {}, "w:del") -- a deleted paragraph break
        if gone then
          ctx.inserted_break = { type = "delete", node = gone, author = xml.attr(gone, "w:author") or "?", date = xml.attr(gone, "w:date") or "", text = "\n" }
        end
      elseif tag == "w:r" then
        run(el, link)
      elseif tag == "w:hyperlink" then
        local id = xml.attr(el, "r:id")
        local url = id and ctx.links[id] or "#" .. (xml.attr(el, "w:anchor") or "")
        ctx.link_attrs[url] = ctx.link_attrs[url] or el.attrs
        walk(el, url)
      elseif tag == "w:ins" or tag == "w:moveTo" then -- (inserted, or moved here: a tracked move)
        local change = { type = "insert", node = el, author = xml.attr(el, "w:author") or "?", date = xml.attr(el, "w:date") or "" }
        change.moved = tag == "w:moveTo" or nil
        add({ insert_start = change })
        walk(el, link)
        add({ insert_end = change })
      elseif (tag == "w:del" or tag == "w:moveFrom") and deleting then
        walk(el, link)
      elseif tag == "w:del" or tag == "w:moveFrom" then -- (deleted, or moved away)
        local change = {
          type = "delete", node = el, author = xml.attr(el, "w:author") or "?", date = xml.attr(el, "w:date") or "",
          text = xml.text(el), moved = tag == "w:moveFrom" or nil,
        }
        if change.text ~= "" then -- (its runs as text, formatted, in a range of the deletion)
          add({ insert_start = change })
          deleting = true
          walk(el, link)
          deleting = false
          add({ insert_end = change })
        else -- (a deleted image, tab...: a point)
          add({ deletion = change })
        end
      elseif tag == "m:oMath" or tag == "m:oMathPara" then -- an equation (its LaTeX comes later)
        add({ object = el, visible = true, math = { display = tag == "m:oMathPara" } })
        table.insert(ctx.equations, spans[#spans])
      elseif tag == "w:commentRangeStart" then
        comment_marker(el, "comment_start")
      elseif tag == "w:commentRangeEnd" then
        comment_marker(el, "comment_end")
      elseif tag and not DROP[tag] then
        local text = visible(el) and xml.text(el) or ""
        local seq = tag == "w:fldSimple" and M.seq_name(xml.attr(el, "w:instr"))
        add({ object = el, visible = visible(el), field = text ~= "" and { text = text, seq = seq or nil } or nil })
      end
    end
  end

  if ctx.inserted_break_before then -- (the previous paragraph's break: up to here)
    add({ insert_end = ctx.inserted_break_before })
    ctx.inserted_break_before = nil
  end
  walk(p)
  if ctx.inserted_break then
    add({ insert_start = ctx.inserted_break })
    ctx.inserted_break_before, ctx.inserted_break = ctx.inserted_break, nil
  end
  return spans, block
end

-- ── Equations (as LaTeX, converted by pandoc) ─────────────────────────────

local function first_math(node) -- the LaTeX of the first equation in a pandoc JSON tree
  if type(node) ~= "table" then
    return nil
  elseif node.t == "Math" then
    return node.c[2]
  end
  for _, child in pairs(node) do
    local tex = first_math(child)
    if tex then
      return tex
    end
  end
end

-- Give each equation span its LaTeX. Equations pandoc can't read (or without pandoc) stay
-- objects, shown as "￼" and kept as they are.
latex_of = function(ctx, spans)
  local texs = {}
  if #spans > 0 and vim.fn.executable("pandoc") == 1 then
    local root = xml.find(ctx.root, "w:document")
    local attrs, body = {}, {}
    for _, kv in ipairs(root.attrs) do
      attrs[#attrs + 1] = (' %s="%s"'):format(kv[1], xml.escape(kv[2], true))
    end
    for _, span in ipairs(spans) do -- (a document of just the equations, one per paragraph)
      body[#body + 1] = "<w:p>" .. xml.serialize(span.object) .. "</w:p>"
    end
    local tmp = vim.fn.tempname() .. ".docx"
    vim.uv.fs_copyfile(ctx.path, tmp)
    zip.write(tmp, { [M.MAIN] = ("<w:document%s><w:body>%s</w:body></w:document>"):format(table.concat(attrs), table.concat(body)) })
    local r = vim.system({ "pandoc", "-f", "docx", "-t", "json", tmp }):wait()
    os.remove(tmp)
    local ok, ast = pcall(vim.json.decode, r.code == 0 and r.stdout or "")
    if ok and #ast.blocks == #spans then
      texs = vim.tbl_map(first_math, ast.blocks)
    end
  end
  for k, span in ipairs(spans) do
    if texs[k] then
      span.math.tex = texs[k]
    else
      span.math = nil
    end
  end
end

-- The XML of equations written as LaTeX ({ tex, display }), converted by pandoc.
function M.equations(ctx, list)
  assert(vim.fn.executable("pandoc") == 1, "editing equations needs pandoc")
  local source = vim.tbl_map(function(eq)
    return eq.display and "\\[" .. eq.tex .. "\\]" or "\\(" .. eq.tex .. "\\)"
  end, list)
  local out = vim.fn.tempname() .. ".docx"
  local r = vim.system({ "pandoc", "-f", "markdown+tex_math_single_backslash", "-o", out }, { stdin = table.concat(source, "\n\n") }):wait()
  assert(r.code == 0, "pandoc could not convert the equations: " .. (r.stderr or ""))
  local paras = xml.collect(xml.parse(zip.read(out, M.MAIN)), { ["w:p"] = true })
  os.remove(out)
  local nodes = {}
  for k, eq in ipairs(list) do
    nodes[k] = paras[k] and xml.find(paras[k], eq.display and "m:oMathPara" or "m:oMath")
    assert(nodes[k], "not an equation pandoc understands: " .. eq.tex)
  end
  ensure_ns(xml.find(ctx.root, "w:document"), "m")
  return nodes
end

function M.load(path)
  local ctx = {
    path = path,
    parts = {}, -- name -> parsed XML (false: the part doesn't exist)
    entries = zip.list(path), -- names of all files in the package
    links = {}, -- relationship id -> url
    link_attrs = {}, -- url -> attributes of its first <w:hyperlink>
    images = {}, -- relationship id -> image file (relative to word/)
    rel_ids = {}, -- all relationship ids in use
    level_of = {}, -- paragraph style id -> heading level
    style_of = {}, -- heading level -> paragraph style id
    comment_of = {}, -- <w:comment> -> comment
    title_styles = {}, -- paragraph style ids of titles
    quote_styles = {}, -- ... and of quotes
    caption_styles = {}, -- ... and of captions (and styles based on them)
    based_on = {}, -- style id -> the id of the style it is based on
    caption_style = {}, -- the id of the styles "caption", "table caption" and "image caption"
    footnote_ids = {}, -- the footnotes referenced when opened
    style_numbering = {}, -- paragraph style id -> its <w:numPr> (list styles, numbered headings)
    equations = {}, -- the spans of equations, in order
    next_comment = 0,
  }
  ctx.root = assert(part(ctx, M.MAIN), "no word/document.xml")
  ctx.rels = part(ctx, RELS, '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>')
  ctx.comments = read_comments(ctx)
  local settings = part(ctx, "word/settings.xml")
  ctx.tracking = settings and is_on(xml.find(settings, "w:trackRevisions")) or false
  -- Headings are styles named "heading N" (the internal name, in English in every language) or
  -- with an outline level; their ids vary ("Heading1", "berschrift1", ...).
  for _, style in ipairs(xml.collect(part(ctx, STYLES) or {}, { ["w:style"] = true })) do
    local name = (xml.attr(xml.child(style, "w:name") or {}, "w:val") or ""):lower()
    local outline = tonumber(xml.attr(xml.find(style, "w:outlineLvl") or {}, "w:val") or "")
    local level = tonumber(name:match("^heading (%d)$")) or (outline and outline < 9 and outline + 1)
    local id = xml.attr(style, "w:styleId")
    if name == "title" and id then -- (Word's "Title" style)
      ctx.title_styles[id] = true
      ctx.title_style = ctx.title_style or id
    elseif (name == "caption" or name == "table caption" or name == "image caption") and id then
      ctx.caption_styles[id] = true
      ctx.caption_style[name] = id
    elseif (name == "quote" or name == "block text" or name == "intense quote") and id then
      ctx.quote_styles[id] = true
      ctx.quote_style = name == "quote" and id or ctx.quote_style or id
    end
    local based = xml.child(style, "w:basedOn")
    if id and based then
      ctx.based_on[id] = xml.attr(based, "w:val")
    end
    local numbering = xml.child(xml.child(style, "w:pPr") or {}, "w:numPr")
    if id and numbering then
      ctx.style_numbering[id] = numbering
    end
    if level and id and xml.attr(style, "w:type") == "paragraph" then
      ctx.level_of[id] = level
      ctx.style_of[level] = name:match("^heading") and id or ctx.style_of[level] or id
    end
  end
  for id in pairs(ctx.based_on) do -- (styles based on a caption style are caption styles: LibreOffice's "Figure")
    local up = id
    for _ = 1, 10 do
      up = ctx.based_on[up]
      if not up then
        break
      elseif ctx.caption_styles[up] then
        ctx.caption_styles[id] = true
        break
      end
    end
  end
  -- the headings are numbered when their styles are (in a list), as Word does it
  ctx.numbered = vim.iter(pairs(ctx.level_of)):any(function(id)
    local numbering = ctx.style_numbering[id]
    return numbering and xml.attr(xml.child(numbering, "w:numId") or {}, "w:val") ~= "0"
  end)
  for _, rel in ipairs(xml.collect(ctx.rels, { Relationship = true })) do
    local id, type, external = xml.attr(rel, "Id"), xml.attr(rel, "Type"), xml.attr(rel, "TargetMode") == "External"
    ctx.rel_ids[id] = true
    if type == REL .. "hyperlink" then
      ctx.links[id] = xml.attr(rel, "Target")
    elseif type == REL .. "image" and not external then
      ctx.images[id] = xml.attr(rel, "Target")
    end
  end
  local paras = {}
  for _, p in ipairs(xml.collect(xml.find(ctx.root, "w:body"), { ["w:p"] = true })) do
    local spans, block = read_para(ctx, p)
    -- its boxes: their paragraphs first (not in a table's cell: there, a box stays as it is)
    local in_cell = false
    local up = p.parent
    while up do
      in_cell = in_cell or up.tag == "w:tc"
      up = up.parent
    end
    local own = {}
    for _, s in ipairs(spans) do
      if s.box and not in_cell then
        for _, bp in ipairs(xml.collect(s.box.content, { ["w:p"] = true })) do
          local bspans, bblock = read_para(ctx, bp)
          paras[#paras + 1] = { node = bp, block = bblock, spans = bspans, box = s.box }
        end
      else
        s.box = nil
        own[#own + 1] = s
      end
    end
    paras[#paras + 1] = { node = p, block = block, spans = own }
  end
  latex_of(ctx, ctx.equations)
  -- (or the headings themselves are numbered, as LibreOffice writes it)
  ctx.numbered = ctx.numbered or vim.iter(paras):any(function(para)
    local direct = para.block.heading and xml.child(xml.child(para.node, "w:pPr") or {}, "w:numPr")
    return direct and xml.attr(xml.child(direct, "w:numId") or {}, "w:val") ~= "0"
  end)
  -- (which items one after another are one list: a sub-item always its item's; an item of the first
  -- level another list where the definition changes, or where an instance numbering again from 1 is
  -- first used, as Word counts them)
  local seen, group, group_abstract, groups = {}, nil, nil, 0
  for _, para in ipairs(paras) do
    local l = para.block.list
    if l and l.abstract then
      local first = not seen[l.id]
      seen[l.id] = true
      if not group or l.level == 1 and (l.abstract ~= group_abstract or first and l.restarts) then
        groups = groups + 1
        group, group_abstract = groups, l.abstract
      end
      l.group = group
    elseif not (l and l.kind) then
      group = nil -- (a paragraph between: what follows is another)
    end
  end
  return ctx, paras
end

-- ── Writing paragraphs ──────────────────────────────────────────────────────

function M.new_para(template)
  local ppr = xml.child(template, "w:pPr")
  ppr = ppr and xml.copy(ppr)
  -- (not what belongs to the paragraph it comes after: a section break, a page break before it,
  -- a tracked change of its properties or of its paragraph mark)
  for _, tag in ipairs({ "w:sectPr", "w:pageBreakBefore", "w:pPrChange" }) do
    local el = ppr and xml.child(ppr, tag)
    if el then
      xml.remove(el)
    end
  end
  local mark = ppr and xml.child(ppr, "w:rPr")
  -- (nor its mark's bold, italics...: a list item's number is shown with them; its font and size stay)
  for _, tag in ipairs({ "w:ins", "w:del", "w:moveFrom", "w:moveTo", "w:rPrChange", "w:b", "w:bCs", "w:i", "w:iCs", "w:u", "w:strike",
    "w:dstrike", "w:caps", "w:smallCaps", "w:highlight", "w:vertAlign", "w:shd" }) do
    local el = mark and xml.child(mark, tag)
    if el then
      xml.remove(el)
    end
  end
  return xml.elem("w:p", {}, { ppr })
end

function M.first_para(ctx) -- at the end of the body, before its section properties
  local body = xml.find(ctx.root, "w:body")
  local section = xml.child(body, "w:sectPr")
  return xml.append(body, xml.elem("w:p"), section and xml.index(section) or nil)
end

local function set_flag(rpr, tag, value) -- value: true (on), false (off), or a w:val
  for k = #rpr.children, 1, -1 do
    local t = rpr.children[k].tag
    if t == tag or t == tag .. "Cs" then
      table.remove(rpr.children, k)
    end
  end
  if value then
    local at = #rpr.children + 1
    for k, c in ipairs(rpr.children) do
      if (RPR_ORDER[c.tag] or math.huge) > RPR_ORDER[tag] then
        at = k
        break
      end
    end
    xml.append(rpr, xml.elem(tag, type(value) == "string" and { { "w:val", value } } or nil), at)
  end
end

-- A run's font set to `name` (false: none of its own), or its size to `size` ("11pt"; false: none).
local function set_font(rpr, name)
  local el = xml.child(rpr, "w:rFonts")
  if not el and name then
    set_flag(rpr, "w:rFonts", true)
    el = xml.child(rpr, "w:rFonts")
  end
  if el then
    for _, attr in ipairs({ "w:ascii", "w:hAnsi", "w:cs" }) do
      xml.set_attr(el, attr, name or nil)
    end
    for _, attr in ipairs({ "w:asciiTheme", "w:hAnsiTheme", "w:cstheme" }) do
      xml.set_attr(el, attr, nil)
    end
    if #(el.attrs or {}) == 0 then
      xml.remove(el)
    end
  end
end
local function set_size(rpr, size)
  local half = size and math.floor(require("neoffice.markup").cm(size) * 72 / 2.54 * 2 + 0.5)
  set_flag(rpr, "w:sz", half and tostring(half))
  if half then
    set_flag(rpr, "w:szCs", tostring(half))
  end
end

-- A run's colour set (a name or "#rrggbb"; false: none of its own, its style's).
local function set_color(rpr, color)
  set_flag(rpr, "w:color", color and markup.color_hex(color):upper())
end

local function text_run(piece, ctx, base)
  local original = piece.props -- the <w:r> this text came from, if any
  local rpr = original and xml.child(original, "w:rPr")
  rpr = rpr and xml.copy(rpr) or xml.elem("w:rPr")
  local have = flags(rpr)
  for flag, tag in pairs({ b = "w:b", i = "w:i", u = "w:u", s = "w:strike", h = "w:highlight" }) do
    if have[flag] ~= (piece[flag] or false) then
      set_flag(rpr, tag, piece[flag] and (flag == "u" and "single" or flag == "h" and word_color(markup.named_color(piece.h)) or true))
    end
  end
  if have.sup ~= piece.sup or have.sub ~= piece.sub then
    set_flag(rpr, "w:vertAlign", piece.sup and "superscript" or piece.sub and "subscript")
  end
  if ctx and (piece.font ~= nil or piece.size ~= nil or piece.color ~= nil) then -- its font, size, colour: only if what it shows changes
    local paras = require("neoffice.paras")
    local own = own_font(ctx, rpr)
    base = vim.tbl_extend("keep", base or {}, { color = "black" })
    for key, set in pairs({ font = set_font, size = set_size, color = set_color }) do
      local want = piece[key]
      if want ~= nil and paras.measure(key, want or base[key]) ~= paras.measure(key, own[key] or base[key]) then
        set(rpr, want)
      end
    end
  end
  local run = xml.elem("w:r", original and vim.deepcopy(original.attrs), { #rpr.children > 0 and rpr or nil })
  for k, chunk in ipairs(vim.split(piece.text, "\t", { plain = true })) do
    if k > 1 then
      xml.append(run, xml.elem("w:tab"))
    end
    if chunk ~= "" then
      xml.append(run, xml.elem("w:t", { { "xml:space", "preserve" } }, { { text = chunk } }))
    end
  end
  return run
end

local function add_rel(ctx, type, target, external)
  local found = has(ctx.rels, "Relationship", "Target", target)
  if found and xml.attr(found, "Type") == type then
    return xml.attr(found, "Id")
  end
  local id, n = nil, 1
  repeat
    id, n = "rIdNeo" .. n, n + 1
  until not ctx.rel_ids[id]
  ctx.rel_ids[id] = true
  local attrs = { { "Id", id }, { "Type", type }, { "Target", target } }
  if external then
    table.insert(attrs, { "TargetMode", "External" })
  end
  xml.append(xml.find(ctx.rels, "Relationships"), xml.elem("Relationship", attrs))
  ctx.dirty[RELS] = true
  return id
end

local function content_type(ctx, kind, key, value, mime) -- kind: "Default" (Extension) or "Override" (PartName)
  local types = part(ctx, TYPES)
  if not has(types, kind, key, value) then
    xml.append(xml.find(types, "Types"), xml.elem(kind, { { key, value }, { "ContentType", mime } }))
    ctx.dirty[TYPES] = true
  end
end

local function hyperlink(ctx, url)
  if ctx.link_attrs[url] then -- as it was
    return xml.elem("w:hyperlink", vim.deepcopy(ctx.link_attrs[url]))
  elseif url:sub(1, 1) == "#" then
    return xml.elem("w:hyperlink", { { "w:anchor", url:sub(2) } })
  end
  local id = add_rel(ctx, REL .. "hyperlink", url, true)
  ctx.links[id] = url
  return xml.elem("w:hyperlink", { { "r:id", id } })
end

-- ── Lists and numbering ─────────────────────────────────────────────────────

-- Put a <w:numPr> (or nil: none) in paragraph or style properties, where the schema wants it.
local BEFORE_NUMPR = { ["w:pStyle"] = true, ["w:keepNext"] = true, ["w:keepLines"] = true, ["w:pageBreakBefore"] = true, ["w:framePr"] = true, ["w:widowControl"] = true }
local function set_numbering(ppr, num_id, ilvl)
  local old = xml.child(ppr, "w:numPr")
  if old then
    xml.remove(old)
  end
  if num_id then
    local at = 1
    for k, c in ipairs(ppr.children) do
      at = BEFORE_NUMPR[c.tag] and k + 1 or at
    end
    local children = { ilvl and xml.elem("w:ilvl", { { "w:val", tostring(ilvl) } }), xml.elem("w:numId", { { "w:val", num_id } }) }
    xml.append(ppr, xml.elem("w:numPr", {}, children), at)
  end
end

-- A new list in numbering.xml (created if needed): kind "bullet", "number" or "headings" (1, 1.1, ...).
local function new_list(ctx, kind)
  local root = xml.find(part(ctx, NUMBERING, '<w:numbering xmlns:w="' .. NS.w .. '"/>'), "w:numbering")
  if not ctx.entries[NUMBERING] then
    content_type(ctx, "Override", "PartName", "/" .. NUMBERING, MIME .. "numbering+xml")
    add_rel(ctx, REL .. "numbering", "numbering.xml")
    ctx.entries[NUMBERING] = true
  end
  local function next_id(tag, attr)
    return tostring(vim.iter(xml.collect(root, { [tag] = true })):fold(0, function(n, el)
      return math.max(n, tonumber(xml.attr(el, attr)) or 0)
    end) + 1)
  end
  local abstract_id, num_id = next_id("w:abstractNum", "w:abstractNumId"), next_id("w:num", "w:numId")
  local abstract = xml.elem("w:abstractNum", { { "w:abstractNumId", abstract_id } }, { xml.elem("w:multiLevelType", { { "w:val", "multilevel" } }) })
  for l = 0, 8 do
    local text, indent = "%" .. (l + 1) .. ".", 720 * (l + 1)
    if kind == "bullet" then
      text = ({ "•", "◦", "▪" })[l % 3 + 1]
    elseif kind == "headings" then
      local parts = {}
      for k = 1, l + 1 do
        parts[k] = "%" .. k
      end
      text, indent = table.concat(parts, "."), 432 + 144 * l
    end
    xml.append(abstract, xml.elem("w:lvl", { { "w:ilvl", tostring(l) } }, {
      xml.elem("w:start", { { "w:val", "1" } }),
      xml.elem("w:numFmt", { { "w:val", kind == "bullet" and "bullet" or "decimal" } }),
      xml.elem("w:lvlText", { { "w:val", text } }),
      xml.elem("w:lvlJc", { { "w:val", "left" } }),
      xml.elem("w:pPr", {}, { xml.elem("w:ind", { { "w:left", tostring(indent) }, { "w:hanging", kind == "headings" and tostring(indent) or "360" } }) }),
    }))
  end
  local first_num = xml.child(root, "w:num") -- (all abstractNum come before the num)
  xml.append(root, abstract, first_num and xml.index(first_num) or nil)
  local nums = xml.collect(root, { ["w:num"] = true }) -- (after the last w:num, as the schema wants)
  local last = nums[#nums] or xml.collect(root, { ["w:abstractNum"] = true })[1] and abstract
  xml.append(root, xml.elem("w:num", { { "w:numId", num_id } }, { xml.elem("w:abstractNumId", { { "w:val", abstract_id } }) }),
    last and last.parent == root and xml.index(last) + 1 or nil)
  ctx.dirty[NUMBERING] = true
  return num_id
end

-- The list of the paragraph just before p, if it is of this kind (a new item continues it).
local function previous_list(ctx, p, kind)
  for k = xml.index(p) - 1, 1, -1 do
    local sibling = p.parent.children[k]
    if sibling.tag == "w:p" then
      local list = list_of(ctx, xml.child(sibling, "w:pPr"))
      return list and list.kind == kind and list.id or nil
    elseif sibling.tag then
      return nil
    end
  end
end

-- Make paragraph p (its properties ppr, maybe nil) an item of a list like `want` ({ kind, level }),
-- or of no list. Returns ppr.
local function set_list(ctx, p, ppr, want)
  local have = list_of(ctx, ppr)
  if want then
    -- its list: the one of the item just before it (no blank line between), else its own (a new
    -- instance numbering again, if that is the one before it), else a new one
    local before = previous_list(ctx, p, want.kind)
    if want.continues then -- (the instance of the nearest item above at its level or higher: not a sub-item's,
      -- nor one of the other kind under it)
      for k = xml.index(p) - 1, 1, -1 do
        local sibling = p.parent.children[k]
        local list = sibling.tag == "w:p" and list_of(ctx, xml.child(sibling, "w:pPr"))
        if sibling.tag and not list then
          break
        elseif list and list.level <= want.level then
          before = list.kind == want.kind and list.id or nil
          break
        end
      end
    end
    local base
    if want.continues and before then
      base = before
    elseif have and have.kind == want.kind then
      local before_any = previous_list(ctx, p, have.kind)
      local _, num = abstract_of(ctx, have.id)
      local restarts = num and vim.iter(xml.collect(num, { ["w:startOverride"] = true })):any(function() return true end)
      local same = before_any and (before_any == have.id or abstract_of(ctx, before_any) == abstract_of(ctx, have.id) and not restarts)
      base = same and restart_num(ctx, have.id) or have.id -- (after a blank line: numbered again from 1)
    else
      base = new_list(ctx, want.kind)
    end
    local num_id = want.kind == "number" and num_for(ctx, p, base, want.level - 1, want.pattern) or base
    if have and have.id == num_id and have.level == want.level then
      return ppr
    end
    ppr = ppr or xml.elem("w:pPr")
    set_numbering(ppr, num_id, want.level - 1)
    if have and have.id ~= num_id then -- (the rest of its list after it comes along)
      for k = xml.index(p) + 1, #p.parent.children do
        local next_p = p.parent.children[k]
        if next_p.tag == "w:p" then
          local next_ppr = xml.child(next_p, "w:pPr")
          local list = list_of(ctx, next_ppr)
          if not list or list.id ~= have.id then
            break
          end
          set_numbering(next_ppr, num_for(ctx, next_p, num_id, list.level - 1, list.pattern), list.level - 1)
        elseif next_p.tag then
          break
        end
      end
    end
  else
    ppr = ppr or xml.elem("w:pPr")
    set_numbering(ppr, nil)
    if list_of(ctx, ppr) then -- its style is a list: this paragraph isn't
      set_numbering(ppr, "0")
    end
  end
  return ppr
end

local function find_style(ctx, id)
  return has(part(ctx, STYLES) or {}, "w:style", "w:styleId", id)
end

-- :Neoffice numbers: number the headings (through their styles, as Word does), or stop (also
-- numbering given to the headings themselves).
function M.number_headings(ctx, on)
  if not on then
    for _, p in ipairs(xml.collect(xml.find(ctx.root, "w:body"), { ["w:p"] = true })) do
      local ppr = xml.child(p, "w:pPr")
      if ppr and block_of(ctx, ppr).heading and xml.child(ppr, "w:numPr") then
        set_numbering(ppr, nil)
        ctx.dirty[M.MAIN] = true
      end
    end
  end
  ctx.heading_list = on and (ctx.heading_list or new_list(ctx, "headings")) or nil
  for id, level in pairs(ctx.level_of) do
    local style = find_style(ctx, id)
    if style then
      local ppr = xml.child(style, "w:pPr") or xml.append(style, xml.elem("w:pPr"))
      set_numbering(ppr, ctx.heading_list, level - 1)
      ctx.style_numbering[id] = xml.child(ppr, "w:numPr")
      ctx.dirty[STYLES] = true
    end
  end
  ctx.numbered = on
end

-- The style id of a title, adding a "Title" style if the document has none.
local function title_style(ctx)
  local styles = xml.find(part(ctx, STYLES) or {}, "w:styles")
  if not ctx.title_style and styles then
    xml.append(styles, xml.elem("w:style", { { "w:type", "paragraph" }, { "w:styleId", "Title" } }, {
      xml.elem("w:name", { { "w:val", "Title" } }),
      xml.elem("w:basedOn", { { "w:val", "Normal" } }),
      xml.elem("w:next", { { "w:val", "Normal" } }),
      xml.elem("w:qFormat"),
      xml.elem("w:pPr", {}, { xml.elem("w:spacing", { { "w:after", "240" } }) }),
      xml.elem("w:rPr", {}, { xml.elem("w:sz", { { "w:val", "56" } }) }),
    }))
    ctx.title_style, ctx.title_styles.Title, ctx.dirty[STYLES] = "Title", true, true
  end
  return ctx.title_style or "Title"
end

-- The style id of heading `level`, adding a "heading N" style if the document has none.
local function heading_style(ctx, level)
  local styles = xml.find(part(ctx, STYLES) or {}, "w:styles")
  if not ctx.style_of[level] and styles then
    local id = "Heading" .. level
    xml.append(styles, xml.elem("w:style", { { "w:type", "paragraph" }, { "w:styleId", id } }, {
      xml.elem("w:name", { { "w:val", "heading " .. level } }),
      xml.elem("w:basedOn", { { "w:val", "Normal" } }),
      xml.elem("w:next", { { "w:val", "Normal" } }),
      xml.elem("w:qFormat"),
      xml.elem("w:pPr", {}, { xml.elem("w:keepNext"), xml.elem("w:outlineLvl", { { "w:val", tostring(level - 1) } }) }),
      xml.elem("w:rPr", {}, { xml.elem("w:b"), xml.elem("w:sz", { { "w:val", tostring(40 - 4 * math.min(level, 5)) } }) }),
    }))
    ctx.style_of[level], ctx.level_of[id], ctx.dirty[STYLES] = id, level, true
    if ctx.numbered then -- (numbered like the other headings)
      local ppr = xml.child(find_style(ctx, id), "w:pPr")
      set_numbering(ppr, ctx.heading_list or new_list(ctx, "headings"), level - 1)
      ctx.style_numbering[id] = xml.child(ppr, "w:numPr")
    end
  end
  return ctx.style_of[level] or "Heading" .. level
end

-- The style id of a caption of this kind ("table", "figure"): the document's "table caption" or
-- "image caption" (pandoc's), else its "caption", else a "caption" style added.
local function caption_style(ctx, kind)
  local id = ctx.caption_style[kind == "table" and "table caption" or "image caption"] or ctx.caption_style.caption
  local styles = xml.find(part(ctx, STYLES) or {}, "w:styles")
  if not id and styles then
    id = "Caption"
    xml.append(styles, xml.elem("w:style", { { "w:type", "paragraph" }, { "w:styleId", id } }, {
      xml.elem("w:name", { { "w:val", "caption" } }),
      xml.elem("w:basedOn", { { "w:val", "Normal" } }),
      xml.elem("w:next", { { "w:val", "Normal" } }),
      xml.elem("w:qFormat"),
      xml.elem("w:pPr", {}, { xml.elem("w:spacing", { { "w:before", "120" }, { "w:after", "120" } }) }),
    }))
    ctx.caption_style.caption, ctx.caption_styles[id], ctx.dirty[STYLES] = id, true, true
  end
  return id or "Caption"
end

-- A caption's number: a field counting the captions of its word ("SEQ Table"), showing `n`.
function M.caption_number(_, label, n)
  return xml.elem("w:fldSimple", { { "w:instr", " SEQ " .. label .. " \\* ARABIC " } }, {
    xml.elem("w:r", {}, { xml.elem("w:rPr", {}, { xml.elem("w:noProof") }), xml.elem("w:t", {}, { { text = tostring(n or 1) } }) }),
  })
end

-- Every caption's number as it is in the document's order (Table 1, 2...: each word counts its
-- own), in the fields' shown results. Whether any changed.
function M.number_captions(ctx)
  local count, changed = {}, false
  local function show(name, texts)
    local key = name:lower()
    count[key] = (count[key] or 0) + 1
    local n = tostring(count[key])
    for k, t in ipairs(texts) do
      local want = k == 1 and n or ""
      if xml.text(t) ~= want then
        t.children, changed = { { text = want } }, true
      end
    end
  end
  for _, p in ipairs(xml.collect(xml.find(ctx.root, "w:body") or {}, { ["w:p"] = true })) do
    local state, name, texts = nil, nil, nil -- (in a field: "instr", then "result")
    for _, el in ipairs(xml.collect(p, { ["w:fldSimple"] = true, ["w:fldChar"] = true, ["w:instrText"] = true, ["w:t"] = true })) do
      if el.tag == "w:fldSimple" then
        local seq = M.seq_name(xml.attr(el, "w:instr"))
        if seq then
          show(seq, xml.collect(el, { ["w:t"] = true }))
        end
      elseif el.tag == "w:fldChar" then
        local kind = xml.attr(el, "w:fldCharType")
        if kind == "begin" then
          state, name, texts = "instr", "", {}
        elseif kind == "separate" and state then
          state = "result"
        elseif kind == "end" and state then
          local seq = M.seq_name(name)
          if seq then
            show(seq, texts)
          end
          state = nil
        end
      elseif el.tag == "w:instrText" and state == "instr" then
        name = name .. xml.text(el)
      elseif el.tag == "w:t" and state == "result" then
        texts[#texts + 1] = el
      end
    end
  end
  if changed then
    ctx.dirty[M.MAIN] = true
  end
  return changed
end

-- The style id of a quote, adding a "Quote" style if the document has none.
local function quote_style(ctx)
  local styles = xml.find(part(ctx, STYLES) or {}, "w:styles")
  if not ctx.quote_style and styles then
    xml.append(styles, xml.elem("w:style", { { "w:type", "paragraph" }, { "w:styleId", "Quote" } }, {
      xml.elem("w:name", { { "w:val", "Quote" } }),
      xml.elem("w:basedOn", { { "w:val", "Normal" } }),
      xml.elem("w:next", { { "w:val", "Normal" } }),
      xml.elem("w:qFormat"),
      xml.elem("w:pPr", {}, { xml.elem("w:ind", { { "w:left", "720" }, { "w:right", "720" } }) }),
      xml.elem("w:rPr", {}, { xml.elem("w:i") }),
    }))
    ctx.quote_style, ctx.quote_styles.Quote, ctx.dirty[STYLES] = "Quote", true, true
  end
  return ctx.quote_style or "Quote"
end

-- Schema order of paragraph properties (up to w:jc's neighbours), to put w:jc where Word wants it.
local PPR_ORDER = {}
for i, tag in ipairs(vim.split("pStyle keepNext keepLines pageBreakBefore framePr widowControl numPr suppressLineNumbers pBdr shd tabs suppressAutoHyphens kinsoku wordWrap overflowPunct topLinePunct autoSpaceDE autoSpaceDN bidi adjustRightInd snapToGrid spacing ind contextualSpacing mirrorIndents suppressOverlap jc textDirection textAlignment textboxTightWrap outlineLvl divId cnfStyle rPr sectPr pPrChange", " ")) do
  PPR_ORDER["w:" .. tag] = i
end

-- Kept on the same page as the paragraph after it (a table's caption, a figure).
function M.keep_with_next(_, p)
  local ppr = xml.child(p, "w:pPr") or xml.append(p, xml.elem("w:pPr"), 1)
  if not xml.child(ppr, "w:keepNext") then
    local at = #ppr.children + 1
    for k, c in ipairs(ppr.children) do
      if (PPR_ORDER[c.tag] or 0) > PPR_ORDER["w:keepNext"] then
        at = k
        break
      end
    end
    xml.append(ppr, xml.elem("w:keepNext"), at)
  end
end

-- A paragraph's own format's attributes set as `want` says (only those whose value changed).
local write_layout -- (below)
local function set_para(ctx, ppr, want)
  local paras = require("neoffice.paras")
  local have = M.para_attrs(ppr, ctx)
  want = vim.deepcopy(want) -- (a value its style gives already: none of its own)
  local id = xml.child(ppr, "w:pStyle") and xml.attr(xml.child(ppr, "w:pStyle"), "w:val")
  local style = style_attrs(ctx, id)
  for _, key in ipairs(paras.LAYOUT) do
    local base = style[key] or (key == "line" and "single" or "0cm")
    if want[key] and want[key] ~= "none" and paras.measure(key, want[key]) == paras.measure(key, base) then
      want[key] = nil
    end
  end
  write_layout(ppr, have, want)
end
-- Paragraph properties' indentation and spacing from `have` to `want` (only those that changed).
write_layout = function(ppr, have, want)
  local paras = require("neoffice.paras")
  local function changed(key) return paras.measure(key, have[key]) ~= paras.measure(key, want[key]) end
  local function twips(key) -- (a length, in twips)
    local cm = want[key] and want[key] ~= "none" and require("neoffice.markup").cm((want[key]:gsub("^%-", "")))
    return cm and math.floor(cm / 2.54 * 1440 + 0.5) * (want[key]:sub(1, 1) == "-" and -1 or 1)
  end
  local function element(tag)
    local el = xml.child(ppr, tag)
    if not el then
      el = xml.elem(tag)
      local at = #ppr.children + 1
      for k, c in ipairs(ppr.children) do
        if (PPR_ORDER[c.tag] or 0) > PPR_ORDER[tag] then
          at = k
          break
        end
      end
      xml.append(ppr, el, at)
    end
    return el
  end
  local function tidy(el) -- (an element left with no attributes goes)
    if el and #el.attrs == 0 then
      xml.remove(el)
    end
  end
  if changed("indent-left") or changed("indent-right") or changed("first") then
    local ind = element("w:ind")
    if changed("indent-left") then
      xml.set_attr(ind, "w:start", nil)
      local v = twips("indent-left")
      xml.set_attr(ind, "w:left", v and tostring(v) or nil)
    end
    if changed("indent-right") then
      xml.set_attr(ind, "w:end", nil)
      local v = twips("indent-right")
      xml.set_attr(ind, "w:right", v and tostring(v) or nil)
    end
    if changed("first") then
      local v = twips("first")
      xml.set_attr(ind, "w:firstLine", v and v >= 0 and tostring(v) or nil)
      xml.set_attr(ind, "w:hanging", v and v < 0 and tostring(-v) or nil)
    end
    tidy(ind)
  end
  if changed("before") or changed("after") or changed("line") then
    local spacing = element("w:spacing")
    for _, key in ipairs({ "before", "after" }) do
      if changed(key) then
        local v = twips(key)
        xml.set_attr(spacing, "w:" .. key, v and tostring(v) or nil)
      end
    end
    if changed("line") then
      local value = want.line
      local multiple = value and ({ single = 1, double = 2 })[value] or tonumber(value or "")
      local exact = not multiple and twips("line")
      xml.set_attr(spacing, "w:line", multiple and tostring(math.floor(multiple * 240 + 0.5)) or exact and tostring(exact) or nil)
      xml.set_attr(spacing, "w:lineRule", multiple and "auto" or exact and "exact" or nil)
    end
    tidy(spacing)
  end
end

-- (only when what it shows changed: "start", "left", "both"... stay as they are; nil: its style's)
local default_style_id -- (below)
local function set_align(ppr, align, style_align)
  local jc = xml.child(ppr, "w:jc")
  local have = jc and (JC[xml.attr(jc, "w:val") or ""] or "left") or style_align or "left"
  if have == (align or style_align or "left") then
    return
  end
  if jc then
    xml.remove(jc)
  end
  if align and align ~= (style_align or "left") then
    local at = #ppr.children + 1
    for k, c in ipairs(ppr.children) do
      if (PPR_ORDER[c.tag] or 0) > PPR_ORDER["w:jc"] then
        at = k
        break
      end
    end
    xml.append(ppr, xml.elem("w:jc", { { "w:val", align == "justify" and "both" or align } }), at)
  end
end
local function own_align(ctx, ppr) -- (a paragraph's alignment, if not its style's: "left" too, then)
  local jc = ppr and xml.child(ppr, "w:jc")
  local style = ppr and xml.child(ppr, "w:pStyle")
  local base = M.style_align(ctx, style and xml.attr(style, "w:val") or default_style_id(ctx)) or "left"
  local have = jc and (JC[xml.attr(jc, "w:val") or ""] or "left") or base
  return have ~= base and have or nil
end
M.own_align = own_align

-- ── Styles (the document's header: see init.lua) ───────────────────────────
-- A paragraph style's id, name (as Word shows it: "Heading 1"), and what it gives (following the
-- styles it is based on, and the defaults): { font, size, align, bold, italic, and paras.LAYOUT }.
local function styles_el(ctx)
  return part(ctx, STYLES) or {}
end
default_style_id = function(ctx)
  local styles = styles_el(ctx)
  local flagged = vim.iter(xml.collect(styles, { ["w:style"] = true })):find(function(st)
    return xml.attr(st, "w:type") == "paragraph" and xml.attr(st, "w:default") == "1"
  end)
  return flagged and xml.attr(flagged, "w:styleId") or (has(styles, "w:style", "w:styleId", "Normal") and "Normal") or nil
end
function M.style_of(ctx, p) -- (a paragraph's)
  local ppr = xml.child(p, "w:pPr")
  local style = ppr and xml.child(ppr, "w:pStyle")
  return style and xml.attr(style, "w:val") or default_style_id(ctx)
end
function M.style_name(ctx, id)
  local style = id and has(styles_el(ctx), "w:style", "w:styleId", id)
  local name = style and xml.child(style, "w:name")
  name = name and xml.attr(name, "w:val") or id or "Normal"
  return (name:gsub("^%l", string.upper))
end
function M.paragraph_styles(ctx) -- (the names of all its paragraph styles)
  local out = {}
  for _, st in ipairs(xml.collect(styles_el(ctx), { ["w:style"] = true })) do
    if xml.attr(st, "w:type") == "paragraph" then
      out[#out + 1] = M.style_name(ctx, xml.attr(st, "w:styleId"))
    end
  end
  return out
end
-- The paragraph styles in use: in the text, headers, footers and notes (ids).
function M.styles_in_use(ctx)
  local used = {}
  for name in pairs(ctx.entries or {}) do
    if name == M.MAIN or name:match("^word/header%d*%.xml$") or name:match("^word/footer%d*%.xml$") or name == "word/footnotes.xml" or name == "word/endnotes.xml" then
      local root = name == M.MAIN and ctx.root or part(ctx, name)
      for _, el in ipairs(root and xml.collect(root, { ["w:pStyle"] = true }) or {}) do
        used[xml.attr(el, "w:val") or ""] = true
      end
    end
  end
  used[""] = nil
  return vim.tbl_keys(used)
end
function M.style_id(ctx, name) -- (by its name, whatever its case; nil if it has none)
  for _, st in ipairs(xml.collect(styles_el(ctx), { ["w:style"] = true })) do
    local n = xml.child(st, "w:name")
    if xml.attr(st, "w:type") == "paragraph" and (n and xml.attr(n, "w:val") or ""):lower() == name:lower() then
      return xml.attr(st, "w:styleId")
    end
  end
end
local function chain_of(ctx, id)
  local chain, cur = {}, id
  for _ = 1, 10 do
    if not cur then
      break
    end
    chain[#chain + 1] = has(styles_el(ctx), "w:style", "w:styleId", cur)
    cur = ctx.based_on[cur]
  end
  return chain
end
function M.style_view(ctx, id)
  local out = vim.deepcopy(style_attrs(ctx, id))
  for _, key in ipairs(require("neoffice.paras").LAYOUT) do -- (the defaults: unsaid)
    if out[key] and require("neoffice.paras").measure(key, out[key]) == require("neoffice.paras").measure(key, key == "line" and "single" or "0cm") then
      out[key] = nil
    end
  end
  local font = style_font(ctx, id, true)
  out.font, out.size = font.font, font.size or "10pt"
  out.color = font.color and markup.color_hex(font.color) ~= "000000" and font.color or nil -- (black, automatic: unsaid)
  local defaults = xml.find(styles_el(ctx), "w:docDefaults")
  local function find(path, tag) -- (in the first style of the chain that says, else the defaults)
    for _, st in ipairs(chain_of(ctx, id)) do
      local el = xml.child(st, path)
      if el and xml.child(el, tag) then
        return xml.child(el, tag)
      end
    end
    local d = defaults and xml.find(defaults, path)
    return d and xml.child(d, tag)
  end
  local jc = find("w:pPr", "w:jc")
  local align = JC[jc and xml.attr(jc, "w:val") or ""]
  out.align = align
  out.bold = is_on(find("w:rPr", "w:b")) or nil
  out.italic = is_on(find("w:rPr", "w:i")) or nil
  return out
end
function M.style_align(ctx, id) -- (what alignment a style gives: nil for left)
  ctx.style_aligns = ctx.style_aligns or {}
  local key = id or false
  if ctx.style_aligns[key] == nil then
    ctx.style_aligns[key] = M.style_view(ctx, id).align or false
  end
  return ctx.style_aligns[key] or nil
end
-- A style's properties changed: { key = value, or false for the default }. Only those given.
function M.set_style(ctx, id, changes)
  local style = assert(has(styles_el(ctx), "w:style", "w:styleId", id), "no style " .. tostring(id))
  local paras = require("neoffice.paras")
  local function child(parent, tag, order)
    local el = xml.child(parent, tag)
    if not el then
      el = xml.elem(tag)
      local at = #parent.children + 1
      for k, c in ipairs(parent.children) do
        if (order[c.tag] or 0) > (order[tag] or 0) then
          at = k
          break
        end
      end
      xml.append(parent, el, at)
    end
    return el
  end
  local STYLE_ORDER = { ["w:name"] = 1, ["w:aliases"] = 2, ["w:basedOn"] = 3, ["w:next"] = 4, ["w:link"] = 5, ["w:autoRedefine"] = 6,
    ["w:hidden"] = 7, ["w:uiPriority"] = 8, ["w:semiHidden"] = 9, ["w:unhideWhenUsed"] = 10, ["w:qFormat"] = 11, ["w:locked"] = 12,
    ["w:personal"] = 13, ["w:rsid"] = 16, ["w:pPr"] = 17, ["w:rPr"] = 18, ["w:tblPr"] = 19, ["w:trPr"] = 20, ["w:tcPr"] = 21, ["w:tblStylePr"] = 22 }
  local layout_have, layout_want, any_layout = nil, nil, false
  for key, value in pairs(changes) do
    if vim.tbl_contains(paras.LAYOUT, key) then
      local ppr = child(style, "w:pPr", STYLE_ORDER)
      layout_have = layout_have or M.raw_para_attrs(ppr)
      layout_want = layout_want or vim.deepcopy(layout_have)
      layout_want[key] = value or (key == "line" and "single" or "0cm")
      any_layout = true
    elseif key == "align" then
      local ppr = child(style, "w:pPr", STYLE_ORDER)
      local jc = xml.child(ppr, "w:jc")
      if jc then
        xml.remove(jc)
      end
      local at = #ppr.children + 1
      for k, c in ipairs(ppr.children) do
        if (PPR_ORDER[c.tag] or 0) > PPR_ORDER["w:jc"] then
          at = k
          break
        end
      end
      xml.append(ppr, xml.elem("w:jc", { { "w:val", value == "justify" and "both" or value or "left" } }), at)
    else
      local rpr = child(style, "w:rPr", STYLE_ORDER)
      if key == "font" then
        set_font(rpr, value)
      elseif key == "size" then
        set_size(rpr, value)
      elseif key == "color" then -- (none: automatic, said, over what a style it is based on gives)
        set_flag(rpr, "w:color", value and markup.color_hex(value):upper() or "auto")
      elseif key == "bold" or key == "italic" then
        set_flag(rpr, key == "bold" and "w:b" or "w:i", value and true or "0")
      end
    end
  end
  if any_layout then
    write_layout(xml.child(style, "w:pPr"), layout_have, layout_want)
  end
  ctx.dirty[STYLES], ctx.style_attrs, ctx.style_fonts, ctx.style_aligns = true, nil, nil, nil -- (what styles give: read again)
end

-- The width of the text in cm: the page's, minus its margins (of the last section: the document's
-- own page). Word's default page (Letter, 1 inch margins) if the document doesn't say.
-- The page, for placing images: { text = the text's width, left, top = its margins } (cm).
function M.page(ctx)
  local section = xml.child(xml.find(ctx.root, "w:body") or {}, "w:sectPr") or {}
  local margins = xml.child(section, "w:pgMar")
  local function twips(attr, default)
    return (tonumber(margins and xml.attr(margins, attr) or "") or default) * 2.54 / 1440
  end
  return { text = M.text_width(ctx), left = twips("w:left", 1440), top = twips("w:top", 1440) }
end

-- Where a drawing (<w:drawing>) is: a geometry (see floats.lua), or nil if it is in the line.
local EMU = 360000 -- (in a centimetre)
function M.drawing_geometry(drawing)
  local anchor = xml.child(drawing, "wp:anchor")
  if not anchor then
    return nil
  end
  local function position(tag, rels)
    local el = xml.child(anchor, tag) or {}
    local align, offset = xml.child(el, "wp:align"), xml.child(el, "wp:posOffset")
    return {
      rel = rels[xml.attr(el, "relativeFrom") or ""] or rels[1],
      align = align and xml.text(align) or nil, offset = offset and (tonumber(xml.text(offset)) or 0) / EMU or nil,
    }
  end
  local h = position("wp:positionH", { "column", page = "page" })
  local v = position("wp:positionV", { "paragraph", page = "page", margin = "margin", topMargin = "margin", bottomMargin = "margin" })
  h.align = ({ left = "left", right = "right", center = "center", inside = "left", outside = "right" })[h.align or ""]
  v.align = ({ top = "top", bottom = "bottom", inside = "top", outside = "bottom" })[v.align or ""]
  local wrap = "front"
  for tag, kind in pairs({ ["wp:wrapSquare"] = "square", ["wp:wrapTight"] = "tight", ["wp:wrapThrough"] = "tight", ["wp:wrapTopAndBottom"] = "none" }) do
    if xml.child(anchor, tag) then
      wrap = kind
    end
  end
  if wrap == "front" and xml.attr(anchor, "behindDoc") == "1" then
    wrap = "behind"
  end
  local dist = tonumber(xml.attr(anchor, wrap == "none" and "distT" or "distL") or "") or 0
  local cx = tonumber(xml.attr(xml.child(anchor, "wp:extent") or {}, "cx") or "") or 0
  return { h = h, v = v, wrap = wrap, gap = dist / EMU, width = cx / EMU }
end

local place_drawing -- (below)

-- What a drawing or object that neoffice doesn't edit is, to show it by name (⟦chart⟧): nil for
-- others.
local GRAPHICS = {
  chart = "chart", diagram = "diagram", wordprocessingGroup = "drawing", wordprocessingShape = "shape",
  wordprocessingCanvas = "drawing", picture = "image",
}
function M.object_kind(el)
  if el.tag == "w:pict" then
    return xml.find(el, "v:textbox") and "box" or "drawing"
  elseif el.tag == "w:object" then
    return "object"
  elseif el.tag == "w:drawing" or el.tag == "mc:AlternateContent" then
    local data = xml.find(el, "a:graphicData")
    local uri = data and xml.attr(data, "uri") or ""
    if xml.find(el, "w:txbxContent") then
      return "box"
    end
    for key, name in pairs(GRAPHICS) do
      if uri:find(key, 1, true) then
        return name
      end
    end
    return "drawing"
  end
end

-- ── Boxes (text boxes: see floats.lua and M.load) ───────────────────────────

-- The paragraphs' container (<w:txbxContent>) of a box in a run's element (<mc:AlternateContent>
-- or <w:drawing>), if it is one that holds only paragraphs (no tables, no boxes in it).
function M.box_content(el)
  local choice = el.tag == "mc:AlternateContent" and xml.child(el, "mc:Choice") or el
  local drawing = choice and (choice.tag == "w:drawing" and choice or xml.child(choice, "w:drawing"))
  local content = drawing and xml.find(drawing, "w:txbxContent")
  if not content or #xml.collect(content, { ["w:tbl"] = true, ["w:txbxContent"] = true }) > 0 then
    return nil
  end
  return content
end

local function box_drawing(run)
  local alt = xml.child(run, "mc:AlternateContent")
  local choice = alt and xml.child(alt, "mc:Choice")
  return xml.child(choice or run, "w:drawing")
end

-- A box's attributes: its width and place (see floats.lua).
function M.box_attrs(ctx, run)
  local drawing = box_drawing(run)
  local cx = tonumber(xml.attr(xml.find(drawing or {}, "wp:extent") or {}, "cx") or "")
  local attrs = require("neoffice.floats").classify(drawing and M.drawing_geometry(drawing), M.page(ctx))
  attrs.width = cx and markup.width(cx / EMU) or nil
  return attrs
end

-- The container of a box's paragraphs (<w:txbxContent>), from its run.
function M.box_container(run)
  return xml.find(box_drawing(run) or run, "w:txbxContent")
end

-- The box a paragraph is in: its run, or nil.
function M.box_of(p)
  local up, content, run = p.parent, nil, nil
  while up do
    if up.tag == "w:txbxContent" then
      content = up
    elseif content and up.tag == "w:r" then
      run = up -- (the outermost: a new box's isn't in a paragraph yet)
    elseif run and up.tag == "w:p" then
      return run
    end
    up = up.parent
  end
  return run
end

-- A new box (empty: its paragraphs go in its content): { node = its run, content }.
function M.new_box(ctx, attrs)
  local document = xml.find(ctx.root, "w:document")
  for _, prefix in ipairs({ "mc", "wps", "wp", "a" }) do
    ensure_ns(document, prefix)
  end
  ctx.next_picture = ctx.next_picture or vim.iter(xml.collect(ctx.root, { ["wp:docPr"] = true }, {})):fold(0, function(acc, el)
    return math.max(acc, tonumber(xml.attr(el, "id")) or 0)
  end) + 1
  local id = ctx.next_picture
  ctx.next_picture = id + 1
  local cx = math.floor((markup.cm(attrs.width or "", M.text_width(ctx)) or M.text_width(ctx) / 2) * EMU)
  local cy = EMU -- (made to fit its content: see M.fit_box)
  local content = xml.elem("w:txbxContent")
  local run = xml.elem("w:r", {}, {
    xml.elem("mc:AlternateContent", {}, {
      xml.elem("mc:Choice", { { "Requires", "wps" } }, {
        xml.elem("w:drawing", {}, {
          xml.elem("wp:inline", {}, {
            xml.elem("wp:extent", { { "cx", tostring(cx) }, { "cy", tostring(cy) } }),
            xml.elem("wp:docPr", { { "id", tostring(id) }, { "name", "Text Box " .. id } }),
            xml.elem("a:graphic", {}, {
              xml.elem("a:graphicData", { { "uri", NS.wps } }, {
                xml.elem("wps:wsp", {}, {
                  xml.elem("wps:cNvSpPr", { { "txBox", "1" } }),
                  xml.elem("wps:spPr", {}, {
                    xml.elem("a:xfrm", {}, { xml.elem("a:off", { { "x", "0" }, { "y", "0" } }), xml.elem("a:ext", { { "cx", tostring(cx) }, { "cy", tostring(cy) } }) }),
                    xml.elem("a:prstGeom", { { "prst", "rect" } }, { xml.elem("a:avLst") }),
                    xml.elem("a:noFill"), xml.elem("a:ln", {}, { xml.elem("a:noFill") }),
                  }),
                  xml.elem("wps:txbx", {}, { content }),
                  xml.elem("wps:bodyPr", { { "wrap", "square" }, { "lIns", "0" }, { "tIns", "0" }, { "rIns", "0" }, { "bIns", "0" }, { "anchor", "t" } }, { xml.elem("a:spAutoFit") }),
                }),
              }),
            }),
          }),
        }),
      }),
    }),
  })
  return { node = run, content = content }
end

-- A box's width and place, as its attributes say (its copy for old programs, if any, goes: it
-- would be out of date).
function M.place_box(ctx, run, attrs)
  local drawing = box_drawing(run)
  if not drawing then
    return
  end
  local cx = attrs.width and markup.cm(attrs.width, M.text_width(ctx))
  if cx then
    for _, tag in ipairs({ "wp:extent", "a:ext" }) do
      local el = xml.find(drawing, tag)
      if el then
        xml.set_attr(el, "cx", tostring(math.floor(cx * EMU)))
      end
    end
  end
  place_drawing(ctx, drawing, attrs)
  M.box_changed(run)
end
function M.box_changed(run)
  local alt = xml.child(run, "mc:AlternateContent")
  local fallback = alt and xml.child(alt, "mc:Fallback")
  if fallback then
    xml.remove(fallback)
  end
end

-- A box as tall as its content (roughly: its images, and its text in lines).
function M.fit_box(ctx, run)
  local drawing = box_drawing(run)
  local content = drawing and xml.find(drawing, "w:txbxContent")
  if not content then
    return
  end
  local cx = tonumber(xml.attr(xml.find(drawing, "wp:extent") or {}, "cx") or "") or EMU
  local height = 0
  for _, p in ipairs(xml.collect(content, { ["w:p"] = true })) do
    for _, extent in ipairs(xml.collect(p, { ["wp:extent"] = true })) do
      height = height + (tonumber(xml.attr(extent, "cy") or "") or 0)
    end
    local chars = #table.concat(vim.tbl_map(xml.text, xml.collect(p, { ["w:t"] = true })))
    local per_line = math.max(cx / EMU / 0.19, 1)
    height = height + math.ceil(chars / per_line) * 0.5 * EMU + 0.2 * EMU
  end
  height = math.floor(math.max(height, 0.5 * EMU))
  for _, tag in ipairs({ "wp:extent", "a:ext" }) do
    local el = xml.find(drawing, tag)
    if el then
      xml.set_attr(el, "cy", tostring(height))
    end
  end
end

-- A box's run, at the start of paragraph `p` (after its properties), unless it is in it already.
function M.attach_box(_, run, p)
  if run.parent == p and xml.index(run) then -- (in it still: not when its paragraph was written again)
    return
  end
  if run.parent then
    xml.remove(run)
  end
  local ppr = xml.child(p, "w:pPr")
  xml.append(p, run, ppr and xml.index(ppr) + 1 or 1)
end

-- A drawing placed as its attributes say (see floats.lua): in the line, or anchored, floating.
place_drawing = function(ctx, drawing, attrs)
  local old = vim.iter(drawing.children):find(function(c) return c.tag == "wp:inline" or c.tag == "wp:anchor" end)
  if not old then
    return
  end
  local parts = {}
  for _, tag in ipairs({ "wp:extent", "wp:effectExtent", "wp:docPr", "wp:cNvGraphicFramePr", "a:graphic" }) do
    parts[tag] = xml.child(old, tag)
  end
  local cx = tonumber(xml.attr(parts["wp:extent"] or {}, "cx") or "") or 0
  local page = M.page(ctx)
  local geo = require("neoffice.floats").geometry(attrs, page, cx / EMU)
  if geo and geo.h.align then -- (as an offset: LibreOffice doesn't wrap text around an aligned text box)
    local w = cx / EMU
    geo.h = { rel = geo.h.rel, offset = ({ left = 0, right = page.text - w, center = (page.text - w) / 2 })[geo.h.align] }
  end
  local new
  if not geo then
    new = xml.elem("wp:inline", { { "distT", "0" }, { "distB", "0" }, { "distL", "0" }, { "distR", "0" } })
  else
    local function emu(n) return tostring(math.floor(n * EMU + 0.5)) end
    local function position(tag, pos, rel)
      return xml.elem(tag, { { "relativeFrom", rel } }, {
        pos.align and xml.elem("wp:align", {}, { { text = pos.align } }) or xml.elem("wp:posOffset", {}, { { text = emu(pos.offset or 0) } }),
      })
    end
    local gap = emu(geo.gap)
    local side = geo.wrap == "none" and "0" or gap
    ctx.next_height = (ctx.next_height or 251659264) + 1
    new = xml.elem("wp:anchor", {
      { "distT", gap }, { "distB", gap }, { "distL", side }, { "distR", side }, { "simplePos", "0" },
      { "relativeHeight", tostring(ctx.next_height) }, { "behindDoc", geo.wrap == "behind" and "1" or "0" },
      { "locked", "0" }, { "layoutInCell", "1" }, { "allowOverlap", "1" },
    }, {
      xml.elem("wp:simplePos", { { "x", "0" }, { "y", "0" } }),
      position("wp:positionH", geo.h, geo.h.rel == "page" and "page" or "column"),
      position("wp:positionV", geo.v, geo.v.rel),
    })
    parts["wp:effectExtent"] = parts["wp:effectExtent"] or xml.elem("wp:effectExtent", { { "l", "0" }, { "t", "0" }, { "r", "0" }, { "b", "0" } })
    parts.wrap = geo.wrap == "square" and xml.elem("wp:wrapSquare", { { "wrapText", "bothSides" } })
      or geo.wrap == "tight" and xml.elem("wp:wrapTight", { { "wrapText", "bothSides" } }, {
        xml.elem("wp:wrapPolygon", { { "edited", "0" } }, {
          xml.elem("wp:start", { { "x", "0" }, { "y", "0" } }), xml.elem("wp:lineTo", { { "x", "0" }, { "y", "21600" } }),
          xml.elem("wp:lineTo", { { "x", "21600" }, { "y", "21600" } }), xml.elem("wp:lineTo", { { "x", "21600" }, { "y", "0" } }),
          xml.elem("wp:lineTo", { { "x", "0" }, { "y", "0" } }),
        }),
      })
      or geo.wrap == "none" and xml.elem("wp:wrapTopAndBottom") or xml.elem("wp:wrapNone")
  end
  for _, tag in ipairs({ "wp:extent", "wp:effectExtent", "wrap", "wp:docPr", "wp:cNvGraphicFramePr", "a:graphic" }) do
    if parts[tag] and (geo or tag ~= "wrap") then
      xml.append(new, parts[tag])
    end
  end
  drawing.children[xml.index(old)] = new
  new.parent, old.parent = drawing, nil
  ensure_ns(xml.find(ctx.root, "w:document"), "wp")
end

-- ── Sections: page settings (see pages.lua) ─────────────────────────────────
-- Word keeps a section's settings (w:sectPr) in its last paragraph's properties, and the last
-- section's in the body.
local SECTION_ORDER = { "w:headerReference", "w:footerReference", "w:footnotePr", "w:endnotePr", "w:type", "w:pgSz", "w:pgMar",
  "w:paperSrc", "w:pgBorders", "w:lnNumType", "w:pgNumType", "w:cols", "w:formProt", "w:vAlign", "w:noEndnote", "w:titlePg",
  "w:textDirection", "w:bidi", "w:rtlGutter", "w:docGrid", "w:printerSettings", "w:sectPrChange" }
local START_OF = { nextPage = "page", continuous = "continuous", evenPage = "even", oddPage = "odd", nextColumn = "continuous" }
local TYPE_OF = { page = "nextPage", continuous = "continuous", even = "evenPage", odd = "oddPage" }
local TWIPS = { A3 = { 16838, 23811 }, A4 = { 11906, 16838 }, A5 = { 8391, 11906 }, B5 = { 9979, 14175 }, Letter = { 12240, 15840 }, Legal = { 12240, 20160 } }
local function twips_cm(v)
  local n = tonumber(v or "")
  return n and math.abs(n) * 2.54 / 1440
end
local function section_props(el)
  local size, margins, kind = xml.child(el or {}, "w:pgSz"), xml.child(el or {}, "w:pgMar"), xml.child(el or {}, "w:type")
  local props = {
    width = twips_cm(size and xml.attr(size, "w:w")) or 21.59, height = twips_cm(size and xml.attr(size, "w:h")) or 27.94,
    start = START_OF[kind and xml.attr(kind, "w:val") or "nextPage"] or "page",
  }
  for _, side in ipairs({ "top", "bottom", "left", "right" }) do
    props[side] = twips_cm(margins and xml.attr(margins, "w:" .. side)) or 2.54
  end
  return props
end
local function section_child(el, tag) -- (created where the schema wants it)
  local have = xml.child(el, tag)
  if have then
    return have
  end
  local rank = {}
  for i, t in ipairs(SECTION_ORDER) do
    rank[t] = i
  end
  local at
  for i, c in ipairs(el.children) do
    if not at and c.tag and (rank[c.tag] or 0) > rank[tag] then
      at = i
    end
  end
  return xml.append(el, xml.elem(tag), at)
end

-- The document's sections: { { start = the index in `paras` of its first paragraph, el = its
-- w:sectPr, props } } (see pages.lua).
function M.sections(ctx, paras)
  local body = xml.find(ctx.root, "w:body")
  local out, start = {}, 1
  ctx.section_els = {}
  for k, para in ipairs(paras) do
    local node = para.node or para
    local ppr = node.parent == body and node.tag == "w:p" and xml.child(node, "w:pPr")
    local el = ppr and xml.child(ppr, "w:sectPr")
    if el and k < #paras then -- (one in the last paragraph: an empty section after it, kept as it is)
      out[#out + 1] = { start = start, el = el, props = section_props(el) }
      ctx.section_els[el] = true
      start = k + 1
    end
  end
  local last = body and xml.child(body, "w:sectPr")
  out[#out + 1] = { start = start, el = last, props = section_props(last) }
  if last then
    ctx.section_els[last] = true
  end
  return out
end

-- Settings changed (see pages.changes) on a section's w:sectPr.
local function set_section(el, ch)
  if ch.size then
    local tw = TWIPS[ch.name or ""]
    local short, long = tw and tw[1] or math.floor(math.min(ch.width, ch.height) * 1440 / 2.54 + 0.5), tw and tw[2] or math.floor(math.max(ch.width, ch.height) * 1440 / 2.54 + 0.5)
    local size = section_child(el, "w:pgSz")
    xml.set_attr(size, "w:w", tostring(ch.landscape and long or short))
    xml.set_attr(size, "w:h", tostring(ch.landscape and short or long))
    xml.set_attr(size, "w:orient", ch.landscape and "landscape" or nil)
  end
  if next(ch.margins) then
    local margins = section_child(el, "w:pgMar")
    for side, cm in pairs(ch.margins) do
      xml.set_attr(margins, "w:" .. side, tostring(math.floor(cm * 1440 / 2.54 + 0.5)))
    end
    for _, attr in ipairs({ { "w:header", "720" }, { "w:footer", "720" }, { "w:gutter", "0" } }) do -- (all required)
      xml.set_attr(margins, attr[1], xml.attr(margins, attr[1]) or attr[2])
    end
    for _, side in ipairs({ "top", "bottom", "left", "right" }) do
      xml.set_attr(margins, "w:" .. side, xml.attr(margins, "w:" .. side) or "1440")
    end
  end
  if ch.start then
    xml.set_attr(section_child(el, "w:type"), "w:val", TYPE_OF[ch.start])
  end
end

-- The sections as the buffer has them: list = { { el = the w:sectPr it had (nil: a new one, from
-- the one before), want = its attributes, holder = its last paragraph (but the last's) } }. Changes
-- only what differs; returns whether anything did.
function M.set_sections(ctx, list)
  local pages = require("neoffice.pages")
  local body = xml.find(ctx.root, "w:body")
  local used, els, changed = {}, {}, false
  for j, s in ipairs(list) do
    local el = s.el
    if not el or used[el] then -- (a new section: as the one it was part of, but for what its line says)
      el = xml.copy(els[j - 1] or xml.elem("w:sectPr"))
      changed = true
    end
    used[el], els[j] = true, el
    local ch = pages.changes(pages.view(section_props(el)), s.want)
    if ch then
      set_section(el, ch)
      changed = true
    end
  end
  local function place(el, parent, at)
    if el.parent ~= parent then
      if el.parent then
        xml.remove(el)
      end
      xml.append(parent, el, at)
      changed = true
    end
  end
  for j, el in ipairs(els) do
    if j < #els then -- (in its last paragraph: one of the body's, else a new one after its table)
      local top = assert(list[j].holder, "a section has no paragraph")
      while top.parent and top.parent ~= body do
        top = top.parent
      end
      if top.tag ~= "w:p" then
        top = xml.append(body, xml.elem("w:p"), xml.index(top) + 1)
        changed = true
      end
      local ppr = xml.child(top, "w:pPr") or xml.append(top, xml.elem("w:pPr"), 1)
      local have = xml.child(ppr, "w:sectPr")
      if have and have ~= el then
        xml.remove(have)
        changed = true
      end
      local after = xml.child(ppr, "w:pPrChange") -- (sectPr comes last in pPr, but for this)
      place(el, ppr, after and xml.index(after) or nil)
    else
      local have = xml.child(body, "w:sectPr")
      if have and have ~= el then
        xml.remove(have)
        changed = true
      end
      place(el, body, nil)
    end
  end
  for el in pairs(ctx.section_els or {}) do -- (those of sections joined to the one before: gone)
    if not used[el] and el.parent then
      xml.remove(el)
      changed = true
    end
  end
  if changed then
    ctx.dirty[M.MAIN] = true
  end
  return changed
end

function M.text_width(ctx)
  local section = xml.child(xml.find(ctx.root, "w:body") or {}, "w:sectPr") or {}
  local size, margins = xml.child(section, "w:pgSz"), xml.child(section, "w:pgMar")
  local function twips(el, attr) return tonumber(el and xml.attr(el, attr) or "") end
  local w, left, right = twips(size, "w:w"), twips(margins, "w:left"), twips(margins, "w:right")
  return (w and left and right) and (w - left - right) * 2.54 / 1440 or 16.51
end

-- ── Tables (see tables.lua) ─────────────────────────────────────────────────

function M.new_table(ctx) -- one cell (rows and cells are added by copying; borders: see tables.lua)
  -- its paragraphs: compact (pandoc's style for them, if the document has it; else no spacing)
  local compact = vim.iter(xml.collect(part(ctx, STYLES) or {}, { ["w:style"] = true })):any(function(s)
    return xml.attr(s, "w:styleId") == "Compact" and xml.attr(s, "w:type") == "paragraph"
  end)
  local ppr = xml.elem("w:pPr", {}, {
    compact and xml.elem("w:pStyle", { { "w:val", "Compact" } }) or xml.elem("w:spacing", { { "w:before", "0" }, { "w:after", "0" } }),
  })
  return xml.elem("w:tbl", {}, {
    xml.elem("w:tblPr", {}, { xml.elem("w:tblW", { { "w:w", "0" }, { "w:type", "auto" } }) }),
    xml.elem("w:tblGrid", {}, { xml.elem("w:gridCol", { { "w:w", "2000" } }) }),
    xml.elem("w:tr", {}, { xml.elem("w:tc", {}, { xml.elem("w:tcPr", {}, { xml.elem("w:tcW", { { "w:w", "0" }, { "w:type", "auto" } }) }), xml.elem("w:p", {}, { ppr }) }) }),
  })
end

-- As many grid columns as the widest row has (after rows or cells were added or removed).
function M.fix_grid(_, tbl)
  local grid, n = xml.child(tbl, "w:tblGrid"), 0
  for _, tr in ipairs(xml.collect(tbl, { ["w:tr"] = true })) do
    local count = 0
    for _, el in ipairs(xml.collect(tr, { ["w:tc"] = true, ["w:gridBefore"] = true, ["w:gridAfter"] = true })) do
      local span = el.tag == "w:tc" and xml.child(xml.child(el, "w:tcPr") or {}, "w:gridSpan") or el
      count = count + (tonumber(span and xml.attr(span, "w:val") or "") or 1)
    end
    n = math.max(n, count)
  end
  local cols = grid and xml.collect(grid, { ["w:gridCol"] = true }) or {}
  while grid and #cols < n do
    cols[#cols + 1] = xml.append(grid, xml.copy(cols[#cols] or xml.elem("w:gridCol", { { "w:w", "2000" } })))
  end
  while #cols > math.max(n, 1) do
    xml.remove(table.remove(cols))
  end
end

-- A table's look (see tables.lua): { rules (a preset of tables.BORDERS, if its lines are one),
-- width, height (if its cell padding is one of neoffice's), position, typeset (by neoffice) }.
local TBLPR_ORDER, TCPR_ORDER = {}, {}
for i, tag in ipairs(vim.split("tblStyle tblpPr tblOverlap bidiVisual tblStyleRowBandSize tblStyleColBandSize tblW jc tblCellSpacing tblInd tblBorders shd tblLayout tblCellMar tblLook tblCaption tblDescription tblPrChange", " ")) do
  TBLPR_ORDER["w:" .. tag] = i
end
for i, tag in ipairs(vim.split("cnfStyle tcW gridSpan hMerge vMerge tcBorders shd noWrap tcMar textDirection tcFitText vAlign hideMark headers cellIns cellDel cellMerge tcPrChange", " ")) do
  TCPR_ORDER["w:" .. tag] = i
end
local function put(parent, el, order) -- (instead of the child with its tag, in schema order)
  local old = xml.child(parent, el.tag)
  if old then
    xml.remove(old)
  end
  local at = #parent.children + 1
  for k, c in ipairs(parent.children) do
    if (order[c.tag] or 0) > order[el.tag] then
      at = k
      break
    end
  end
  return xml.append(parent, el, at)
end

-- A row with cells added (`new`: tc -> { filled }): those not merged as the one they were copied
-- from; empty ones beyond the table's columns (padding, as a row with merged cells written with
-- all its cells) removed; one with text taking a column from a merged cell before it. Each cell
-- changed as wide as its columns.
function M.fit_row(_, tbl, tr, new)
  local grid = xml.child(tbl, "w:tblGrid")
  local widths = {}
  for _, c in ipairs(grid and xml.collect(grid, { ["w:gridCol"] = true }) or {}) do
    widths[#widths + 1] = tonumber(xml.attr(c, "w:w") or "") or 0
  end
  local function span(tc)
    local el = xml.child(xml.child(tc, "w:tcPr") or {}, "w:gridSpan")
    return tonumber(el and xml.attr(el, "w:val") or "") or 1
  end
  local touched = {}
  local function set_span(tc, n)
    local tcpr = xml.child(tc, "w:tcPr") or xml.append(tc, xml.elem("w:tcPr"), 1)
    local el = xml.child(tcpr, "w:gridSpan")
    if n <= 1 then
      if el then
        xml.remove(el)
      end
    else
      el = el or put(tcpr, xml.elem("w:gridSpan"), TCPR_ORDER)
      xml.set_attr(el, "w:val", tostring(n))
    end
    touched[tc] = true
  end
  local cells = vim.tbl_filter(function(c) return c.tag == "w:tc" end, tr.children)
  for _, tc in ipairs(cells) do
    if new[tc] then
      set_span(tc, 1)
      local merge = xml.child(xml.child(tc, "w:tcPr") or {}, "w:hMerge")
      if merge then
        xml.remove(merge)
      end
    end
  end
  local trpr = xml.child(tr, "w:trPr") or {}
  local function grid_skip(tag)
    local el = xml.child(trpr, tag)
    return tonumber(el and xml.attr(el, "w:val") or "") or 0
  end
  local total = grid_skip("w:gridBefore") + grid_skip("w:gridAfter")
  for _, tc in ipairs(cells) do
    total = total + span(tc)
  end
  if #widths == 0 then
    return
  end
  for k = #cells, 1, -1 do -- (padding: new empty cells, from the end)
    if total <= #widths then
      break
    end
    if new[cells[k]] and not new[cells[k]].filled then
      xml.remove(cells[k])
      table.remove(cells, k)
      total = total - 1
    end
  end
  for k, tc in ipairs(cells) do -- (a new cell with text: a column of the merged cell before it)
    if total <= #widths then
      break
    end
    if new[tc] then
      for j = k - 1, 1, -1 do
        if not new[cells[j]] and span(cells[j]) > 1 then
          set_span(cells[j], span(cells[j]) - 1)
          total = total - 1
          break
        end
      end
    end
  end
  if total <= #widths then
    local col = grid_skip("w:gridBefore") + 1
    for _, tc in ipairs(cells) do
      local w = 0
      for c = col, col + span(tc) - 1 do
        w = w + (widths[c] or 0)
      end
      local width = xml.child(xml.child(tc, "w:tcPr") or {}, "w:tcW")
      if touched[tc] and width and w > 0 then
        xml.set_attr(width, "w:w", tostring(w))
        xml.set_attr(width, "w:type", "dxa")
      end
      col = col + span(tc)
    end
  end
end

local function props_of(el, tag) -- (its w:tblPr or w:tcPr, made if needed: the first child)
  return xml.child(el, tag) or xml.append(el, xml.elem(tag), 1)
end
local function grid_of(tbl) -- its rows of cells
  return vim.tbl_map(function(tr) return xml.collect(tr, { ["w:tc"] = true }) end, xml.collect(tbl, { ["w:tr"] = true }))
end

local function line_width(borders, names) -- (0 if none, 1 thin, 2 thick; nil if not said)
  for _, name in ipairs(names) do
    local el = xml.child(borders, "w:" .. name)
    if el then
      local val = xml.attr(el, "w:val")
      local sz = tonumber(xml.attr(el, "w:sz") or "") or 4 -- (in eighths of a point)
      return (val == "nil" or val == "none") and 0 or sz >= 6 and 2 or 1 -- (thick from 3/4 of a point)
    end
  end
end
-- A table's style and those it is based on (the most specific first): the document's default
-- table style if it names none.
local function style_chain(ctx, tbl)
  if not ctx.table_styles then
    ctx.table_styles = {}
    for _, style in ipairs(xml.collect(part(ctx, STYLES) or {}, { ["w:style"] = true })) do
      if xml.attr(style, "w:type") == "table" then
        ctx.table_styles[xml.attr(style, "w:styleId") or ""] = style
        if xml.attr(style, "w:default") == "1" then
          ctx.table_styles[false] = style
        end
      end
    end
  end
  local ref = xml.child(xml.child(tbl, "w:tblPr") or {}, "w:tblStyle")
  local style, chain = ctx.table_styles[ref and xml.attr(ref, "w:val") or false], {}
  while style and #chain < 10 do
    chain[#chain + 1] = style
    local based = xml.child(style, "w:basedOn")
    style = based and ctx.table_styles[xml.attr(based, "w:val") or ""]
  end
  return chain
end
local function look(tbl, name, bit) -- (whether a table shows its style's formatting for its first row...)
  local el = xml.child(xml.child(tbl, "w:tblPr") or {}, "w:tblLook")
  local on = el and xml.attr(el, "w:" .. name)
  if on then
    return on == "1" or on == "true" or on == "on"
  end
  local val = el and tonumber(xml.attr(el, "w:val") or "", 16)
  return val ~= nil and math.floor(val / bit) % 2 == 1
end

function M.table_format(ctx, tbl)
  local tables = require("neoffice.tables")
  -- a cell's borders: its own; else its style's for the first or last row; else the table's (its
  -- own, then its style's)
  local chain = style_chain(ctx, tbl)
  local own = { xml.child(xml.child(tbl, "w:tblPr") or {}, "w:tblBorders") or {} }
  for _, style in ipairs(chain) do
    own[#own + 1] = xml.child(xml.child(style, "w:tblPr") or {}, "w:tblBorders") or {}
  end
  local function conditional(kind)
    local out = {}
    for _, style in ipairs(chain) do
      for _, pr in ipairs(xml.collect(style, { ["w:tblStylePr"] = true })) do
        if xml.attr(pr, "w:type") == kind then
          out[#out + 1] = xml.child(xml.child(pr, "w:tcPr") or {}, "w:tcBorders") or {}
        end
      end
    end
    return out
  end
  local first = look(tbl, "firstRow", 0x20) and conditional("firstRow") or {}
  local last = look(tbl, "lastRow", 0x40) and conditional("lastRow") or {}
  local rows = grid_of(tbl)
  local grid = {}
  for r, row in ipairs(rows) do
    grid[r] = {}
    for c, tc in ipairs(row) do
      local list = { xml.child(xml.child(tc, "w:tcPr") or {}, "w:tcBorders") or {} }
      vim.list_extend(list, r == 1 and first or {})
      vim.list_extend(list, r == #rows and last or {})
      local function side(names, outer, edge, inner)
        for _, borders in ipairs(list) do
          local w = line_width(borders, names)
          if w then
            return w
          end
        end
        for _, borders in ipairs(own) do
          local w = line_width(borders, edge and outer or inner)
          if w then
            return w
          end
        end
        return 0 -- (said nowhere: none)
      end
      grid[r][c] = {
        top = side({ "top" }, { "top" }, r == 1, { "insideH" }),
        bottom = side({ "bottom" }, { "bottom" }, r == #rows, { "insideH" }),
        left = side({ "left", "start" }, { "left", "start" }, c == 1, { "insideV" }),
        right = side({ "right", "end" }, { "right", "end" }, c == #row, { "insideV" }),
      }
    end
  end
  local tblpr = xml.child(tbl, "w:tblPr") or {}
  local format = { rules = tables.borders_of(grid), width = "auto", position = "left" }
  local width = xml.child(tblpr, "w:tblW")
  local w, kind = width and xml.attr(width, "w:w") or "0", width and xml.attr(width, "w:type")
  if kind == "pct" then
    local percent = w:match("%%$") and tonumber(w:sub(1, -2)) or (tonumber(w) or 0) / 50
    format.width = percent >= 100 and "full" or percent > 0 and ("%g%%"):format(percent) or "auto"
  elseif kind == "dxa" and (tonumber(w) or 0) > 0 then
    format.width = markup.width(tonumber(w) * 2.54 / 1440)
  end
  local jc = xml.child(tblpr, "w:jc")
  local align = jc and xml.attr(jc, "w:val")
  format.position = align == "center" and "center" or (align == "right" or align == "end") and "right" or "left"
  if format.width == "full" then -- (where it is doesn't matter)
    format.position = tables.defaults().position
  end
  local margins = xml.child(tblpr, "w:tblCellMar") or {}
  local function margin(side)
    local el = xml.child(margins, "w:" .. side)
    return el and (xml.attr(el, "w:type") or "dxa") == "dxa" and (tonumber(xml.attr(el, "w:w") or "") or 0) * 2.54 / 1440 or nil
  end
  local top, bottom = margin("top"), margin("bottom")
  format.height = top and bottom and math.abs(top - bottom) < 0.006 and tables.height_of(top) or nil
  format.typeset = format.rules ~= nil and format.height ~= nil -- (by neoffice: it keeps it so)
  return format
end

-- The size of the document's text (pt): its Normal style's, or its default.
local function font_size(ctx)
  local styles = part(ctx, STYLES) or {}
  for _, style in ipairs(xml.collect(styles, { ["w:style"] = true })) do
    if xml.attr(style, "w:type") == "paragraph" and xml.attr(style, "w:default") == "1" then
      local sz = xml.find(style, "w:sz")
      if sz and tonumber(xml.attr(sz, "w:val") or "") then
        return tonumber(xml.attr(sz, "w:val")) / 2
      end
    end
  end
  local defaults = xml.find(styles, "w:rPrDefault")
  local sz = defaults and xml.find(defaults, "w:sz")
  return sz and tonumber(xml.attr(sz, "w:val") or "") and tonumber(xml.attr(sz, "w:val")) / 2 or 10
end

local function twips(cm)
  return tostring(math.floor(cm / 2.54 * 1440 + 0.5))
end

-- Typeset a table with these options (see tables.lua; those not given stay as they are): its rules
-- (on every cell, and the table), column widths (from its content), row height (cell padding,
-- and paragraphs without spacing), position; and with `typeset`, its first row repeated on each
-- page and at the bottom of its cells, and rows that don't break across pages.
local SIZE = { "4", "8" } -- (thin and thick lines: half a point and a point)
local function border(tag, w)
  return xml.elem(tag, w > 0 and { { "w:val", "single" }, { "w:sz", SIZE[w] }, { "w:space", "0" }, { "w:color", "auto" } } or { { "w:val", "nil" } })
end
function M.set_table_format(ctx, tbl, set)
  local tables = require("neoffice.tables")
  local tblpr = props_of(tbl, "w:tblPr")
  local grid = grid_of(tbl)
  if set.rules then
    local preset = tables.BORDERS[set.rules]
    local borders = xml.elem("w:tblBorders")
    for _, side in ipairs({ "top", "left", "bottom", "right", "insideH", "insideV" }) do
      xml.append(borders, border("w:" .. side, preset[side] or 0)) -- (the line under the first row: the cells')
    end
    put(tblpr, borders, TBLPR_ORDER)
    for r, row in ipairs(grid) do
      for c, tc in ipairs(row) do
        local sides = tables.cell_borders(set.rules, r, #grid, c, #row)
        put(props_of(tc, "w:tcPr"), xml.elem("w:tcBorders", {}, {
          border("w:top", sides.top), border("w:left", sides.left), border("w:bottom", sides.bottom), border("w:right", sides.right),
        }), TCPR_ORDER)
      end
    end
  end
  if set.position then
    put(tblpr, xml.elem("w:jc", { { "w:val", set.position } }), TBLPR_ORDER)
  end
  if set.height then
    local tb, lr = twips(tables.HEIGHT[set.height]), twips(tables.PADDING)
    put(tblpr, xml.elem("w:tblCellMar", {}, {
      xml.elem("w:top", { { "w:w", tb }, { "w:type", "dxa" } }), xml.elem("w:left", { { "w:w", lr }, { "w:type", "dxa" } }),
      xml.elem("w:bottom", { { "w:w", tb }, { "w:type", "dxa" } }), xml.elem("w:right", { { "w:w", lr }, { "w:type", "dxa" } }),
    }), TBLPR_ORDER)
    -- the height is the padding's: paragraphs without spacing or indentation, single-spaced
    for _, p in ipairs(xml.collect(tbl, { ["w:p"] = true }, { ["w:tbl"] = true })) do
      local ppr = props_of(p, "w:pPr")
      put(ppr, xml.elem("w:spacing", { { "w:before", "0" }, { "w:after", "0" }, { "w:line", "240" }, { "w:lineRule", "auto" } }), PPR_ORDER)
      if not xml.child(ppr, "w:numPr") then
        put(ppr, xml.elem("w:ind", { { "w:left", "0" }, { "w:right", "0" }, { "w:firstLine", "0" } }), PPR_ORDER)
      end
    end
  end
  if set.width then
    -- the columns: as wide as their content (each cell spanning one column counts)
    local cols = xml.collect(xml.child(tbl, "w:tblGrid") or {}, { ["w:gridCol"] = true })
    local columns = {}
    for c = 1, #cols do
      columns[c] = {}
    end
    for _, row in ipairs(grid) do
      local c = 1
      for _, tc in ipairs(row) do
        local span = xml.child(xml.child(tc, "w:tcPr") or {}, "w:gridSpan")
        span = tonumber(span and xml.attr(span, "w:val") or "") or 1
        if span == 1 and columns[c] then
          columns[c][#columns[c] + 1] = vim.tbl_map(function(p)
            return table.concat(vim.tbl_map(xml.text, xml.collect(p, { ["w:t"] = true })))
          end, xml.collect(tc, { ["w:p"] = true }, { ["w:tbl"] = true }))
        end
        c = c + span
      end
    end
    local widths = tables.widths(columns, set.width, font_size(ctx) / 72 * 2.54, M.text_width(ctx))
    for c, col in ipairs(cols) do
      xml.set_attr(col, "w:w", twips(widths[c]))
    end
    for _, row in ipairs(grid) do
      local c = 1
      for _, tc in ipairs(row) do
        local span = xml.child(xml.child(tc, "w:tcPr") or {}, "w:gridSpan")
        span = tonumber(span and xml.attr(span, "w:val") or "") or 1
        local w = 0
        for k = c, math.min(c + span - 1, #widths) do
          w = w + widths[k]
        end
        put(props_of(tc, "w:tcPr"), xml.elem("w:tcW", { { "w:w", twips(w) }, { "w:type", "dxa" } }), TCPR_ORDER)
        c = c + span
      end
    end
    local percent = tonumber(set.width:match("^([%d.]+)%%$") or "")
    local cm = not percent and markup.cm(set.width)
    local w, kind = "0", "auto" -- (auto: Word's AutoFit to the content, from these widths)
    if set.width == "full" then
      w, kind = "5000", "pct"
    elseif percent then
      w, kind = tostring(math.floor(percent * 50 + 0.5)), "pct"
    elseif cm then
      w, kind = twips(cm), "dxa"
    end
    put(tblpr, xml.elem("w:tblW", { { "w:w", w }, { "w:type", kind } }), TBLPR_ORDER)
    local layout = xml.child(tblpr, "w:tblLayout") -- (AutoFit)
    if layout then
      xml.remove(layout)
    end
  end
  if set.typeset then
    local before = tbl.parent and tbl.parent.children[xml.index(tbl) - 1] -- (its caption: kept with it)
    if before and before.tag == "w:p" then
      put(props_of(before, "w:pPr"), xml.elem("w:keepNext"), PPR_ORDER)
    end
    for r, row in ipairs(grid) do
      local tr = row[1] and row[1].parent or nil
      if tr then
        local trpr = xml.child(tr, "w:trPr")
        if not trpr then
          local ex = xml.child(tr, "w:tblPrEx")
          trpr = xml.append(tr, xml.elem("w:trPr"), ex and xml.index(ex) + 1 or 1)
        end
        for _, tag in ipairs({ "w:tblHeader", "w:cantSplit" }) do
          local el = xml.child(trpr, tag)
          if el then
            xml.remove(el)
          end
        end
        xml.append(trpr, xml.elem("w:cantSplit"))
        if r == 1 then
          xml.append(trpr, xml.elem("w:tblHeader"))
        end
      end
      -- (a short table on one page: each row's paragraphs kept with the next row)
      local keep = #grid <= tables.KEEP_ROWS and r < #grid
      for _, tc in ipairs(row) do
        for _, p in ipairs(xml.collect(tc, { ["w:p"] = true }, { ["w:tbl"] = true })) do
          local ppr = props_of(p, "w:pPr")
          local old = xml.child(ppr, "w:keepNext")
          if old then
            xml.remove(old)
          end
          if keep then
            put(ppr, xml.elem("w:keepNext"), PPR_ORDER)
          end
        end
        local tcpr = props_of(tc, "w:tcPr")
        local v = xml.child(tcpr, "w:vAlign")
        if v then
          xml.remove(v)
        end
        if r == 1 and #grid > 1 then
          put(tcpr, xml.elem("w:vAlign", { { "w:val", "bottom" } }), TCPR_ORDER)
        end
      end
    end
  end
end

function M.after_table(tbl) -- (Word wants a paragraph after a table: once all paragraphs are in place;
  -- also between two tables, which would otherwise be one)
  local after = tbl.parent and tbl.parent.children[xml.index(tbl) + 1]
  local before = tbl.parent and tbl.parent.children[xml.index(tbl) - 1]
  if before and before.tag == "w:tbl" then
    xml.append(tbl.parent, xml.elem("w:p"), xml.index(tbl))
  end
  if tbl.parent and (not after or after.tag == "w:sectPr" or after.tag == "w:tbl") then
    xml.append(tbl.parent, xml.elem("w:p"), xml.index(tbl) + 1)
  end
end

-- ── Footnotes ───────────────────────────────────────────────────────────────

local FOOTNOTES_XML = '<w:footnotes xmlns:w="' .. NS.w .. '"><w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote></w:footnotes>'

local function footnote_body(note, text, model) -- (model: an existing footnote, to look like it)
  local first = model and xml.child(model, "w:p")
  local ppr = first and xml.child(first, "w:pPr")
  local ref = first and vim.iter(first.children):find(function(r) return xml.child(r, "w:footnoteRef") end)
  ref = ref and xml.copy(ref) or xml.elem("w:r", {}, {
    xml.elem("w:rPr", {}, { xml.elem("w:vertAlign", { { "w:val", "superscript" } }) }), xml.elem("w:footnoteRef"),
  })
  local t = xml.elem("w:t", { { "xml:space", "preserve" } }, { { text = " " .. text } })
  note.children = {}
  xml.append(note, xml.elem("w:p", {}, { ppr and xml.copy(ppr), ref, xml.elem("w:r", {}, { t }) }))
end

-- The reference run of a footnote piece: the same footnote with new text, or a new one.
local function footnote_node(ctx, piece)
  local original = piece.original and piece.original.footnote and piece.original
  local notes = xml.find(part(ctx, FOOTNOTES, FOOTNOTES_XML), "w:footnotes")
  ctx.dirty[FOOTNOTES] = true
  local note = original and has(notes, "w:footnote", "w:id", original.footnote.id)
  if note then
    footnote_body(note, piece.text, note)
    piece.id = original.footnote.id
    return original.node
  end
  local last, model = 0, nil -- (the highest id, and a footnote to copy the look of)
  for _, el in ipairs(xml.collect(notes, { ["w:footnote"] = true })) do
    local n = tonumber(xml.attr(el, "w:id")) or 0
    last, model = math.max(last, n), model or (n > 0 and el or nil)
  end
  local id = tostring(last + 1)
  note = xml.append(notes, xml.elem("w:footnote", { { "w:id", id } }))
  footnote_body(note, piece.text, model)
  add_rel(ctx, REL .. "footnotes", "footnotes.xml") -- (registered in the package, if it wasn't)
  content_type(ctx, "Override", "PartName", "/" .. FOOTNOTES, MIME .. "footnotes+xml")
  local old_ref = vim.iter(xml.collect(ctx.root, { ["w:footnoteReference"] = true })):next()
  local rpr = old_ref and xml.child(old_ref.parent, "w:rPr")
  rpr = rpr and xml.copy(rpr) or xml.elem("w:rPr", {}, { xml.elem("w:vertAlign", { { "w:val", "superscript" } }) })
  piece.id = id
  return xml.elem("w:r", {}, { rpr, xml.elem("w:footnoteReference", { { "w:id", id } }) })
end

-- <w:ins>/<w:del> for a change made in Neovim
local function revision(ctx, tag, ch)
  ctx.revision = (ctx.revision or 900000) + 1
  return xml.elem(tag, { { "w:id", tostring(ctx.revision) }, { "w:author", ch.author }, { "w:date", ch.date } })
end

local function comment_id(ctx, c)
  if not c.id then
    c.id, ctx.next_comment = tostring(ctx.next_comment), ctx.next_comment + 1
  end
  return c.id
end

-- An image piece (see init.lua): a new image, another file for an image, or new alt text.
local DRAWING = [[<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="%d" cy="%d"/><wp:docPr id="%d" name="Picture %d" descr=""/><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="0" name="%s"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed=""/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="%d" cy="%d"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>]]

local function image_node(ctx, img)
  local original = img.original and img.original.image and img.original
  local node
  local wanted = img.width and markup.cm(img.width, M.text_width(ctx)) * 360000 -- (in EMU)
  if original and original.image.path == img.path then -- only the alt text or the width changed
    node = xml.copy(original.node)
    local extent = xml.find(node, "wp:extent")
    local cx, cy = tonumber(xml.attr(extent, "cx")), tonumber(xml.attr(extent, "cy"))
    if wanted and img.width ~= original.image.width and cx and cy then -- (the same shape)
      for _, tag in ipairs({ "wp:extent", "a:ext" }) do
        local el = xml.find(node, tag)
        if el then
          xml.set_attr(el, "cx", tostring(math.floor(wanted)))
          xml.set_attr(el, "cy", tostring(math.floor(wanted * cy / cx)))
        end
      end
    end
  else
    local ext, width, height = images.info(img.bytes)
    local file = img.path
    if not ctx.entries[file] then -- a file from disk goes into word/media/
      local n = 1
      repeat
        file, n = ("word/media/neoffice%d.%s"):format(n, ext), n + 1
      until not ctx.entries[file]
      ctx.entries[file], ctx.files[file] = true, img.bytes
    end
    content_type(ctx, "Default", "Extension", ext, "image/" .. ext)
    local rid = add_rel(ctx, REL .. "image", (file:gsub("^word/", "")))
    local cx, cy = width * 9525, height * 9525 -- (96 pixels per inch)
    if original then -- same width as the image it replaces (unless another is given)
      local extent = xml.find(original.node, "wp:extent")
      cx = wanted or tonumber(xml.attr(extent, "cx")) or cx
      cy = math.floor(cx * height / width)
      node = xml.copy(original.node)
    else
      cx = wanted or math.min(cx, MAX_WIDTH)
      cy = math.floor(cx * height / width)
      ctx.next_picture = ctx.next_picture or vim.iter(xml.collect(ctx.root, { ["wp:docPr"] = true }, {})):fold(0, function(acc, el)
        return math.max(acc, tonumber(xml.attr(el, "id")) or 0)
      end) + 1
      node = xml.parse(DRAWING:format(cx, cy, ctx.next_picture, ctx.next_picture, xml.escape(vim.fs.basename(file), true), cx, cy)).children[1]
      ctx.next_picture = ctx.next_picture + 1
      local document = xml.find(ctx.root, "w:document")
      ensure_ns(document, "wp")
      ensure_ns(document, "r")
    end
    xml.set_attr(xml.find(node, "a:blip"), "r:embed", rid)
    for _, tag in ipairs({ "wp:extent", "a:ext" }) do
      local el = xml.find(node, tag)
      if el then
        xml.set_attr(el, "cx", tostring(math.floor(cx)))
        xml.set_attr(el, "cy", tostring(math.floor(cy)))
      end
    end
  end
  local drawing = xml.find(node, "w:drawing")
  local floats = require("neoffice.floats")
  local before = original and floats.crop(original.image) or nil
  local after = floats.crop(img)
  if drawing and vim.inspect(before) ~= vim.inspect(after) then -- its crop changed
    local fill = xml.find(drawing, "pic:blipFill")
    local rect = fill and xml.child(fill, "a:srcRect")
    if fill and after then
      rect = rect or xml.append(fill, xml.elem("a:srcRect"), (xml.index(xml.child(fill, "a:blip") or fill.children[1] or fill) or 0) + 1)
      for k, side in ipairs({ "t", "r", "b", "l" }) do
        xml.set_attr(rect, side, tostring(math.floor(after[k] * 100000 + 0.5)))
      end
    elseif rect then
      xml.remove(rect)
    end
    -- the same width; the height of what is shown
    local b, a = before or { 0, 0, 0, 0 }, after or { 0, 0, 0, 0 }
    local factor = (1 - a[1] - a[3]) / math.max(1 - b[1] - b[3], 0.01) * (1 - b[2] - b[4]) / math.max(1 - a[2] - a[4], 0.01)
    for _, tag in ipairs({ "wp:extent", "a:ext" }) do
      local el = xml.find(drawing, tag)
      local cy = el and tonumber(xml.attr(el, "cy") or "")
      if cy then
        xml.set_attr(el, "cy", tostring(math.floor(cy * factor + 0.5)))
      end
    end
  end
  if drawing and (not original or floats.place_key(original.image) ~= floats.place_key(img)) then -- (moved, or new)
    place_drawing(ctx, drawing, img)
  end
  local doc_pr = xml.find(node, "wp:docPr")
  if doc_pr then
    xml.set_attr(doc_pr, "descr", img.alt)
  end
  img.node = node
  return node
end

function M.fill(ctx, p, block, pieces)
  local ppr = xml.child(p, "w:pPr")
  local was_caption = false -- (it had a caption's number: its caption style goes if it no longer is one)
  for _, f in ipairs(xml.collect(p, { ["w:fldSimple"] = true, ["w:instrText"] = true })) do
    was_caption = was_caption or M.seq_name(f.tag == "w:fldSimple" and xml.attr(f, "w:instr") or xml.text(f)) ~= nil
  end
  p.children = {}
  -- comment references a deletion written back as it was holds (a reference run, deleted): not
  -- written again for its comment
  local held = {}
  for _, piece in ipairs(pieces) do
    local del = piece.deletion
    if del and del.node and piece.deleted == del.text then
      for _, ref in ipairs(xml.collect(del.node, { ["w:commentReference"] = true })) do
        held[xml.attr(ref, "w:id") or ""] = true
      end
    end
  end
  -- what kind of paragraph it is: a heading or title (by its style), a list item, or plain
  local style = ppr and xml.child(ppr, "w:pStyle")
  local current = style and xml.attr(style, "w:val")
  local styled = current and (ctx.level_of[current] or ctx.title_styles[current] or ctx.quote_styles[current]
    or was_caption and ctx.caption_styles[current])
  local want = block.heading and heading_style(ctx, block.heading) or block.title and title_style(ctx)
    or block.quote and quote_style(ctx) or block.caption and caption_style(ctx, block.caption)
  if want then
    local was_item = not styled and list_of(ctx, ppr)
    ppr = ppr or xml.elem("w:pPr")
    style = style or xml.append(ppr, xml.elem("w:pStyle"), 1)
    xml.set_attr(style, "w:val", want)
    if was_item then
      set_numbering(ppr, nil)
    end
  else
    if styled then
      xml.remove(style)
    end
    if block.list or list_of(ctx, ppr) then
      ppr = set_list(ctx, p, ppr, block.list)
    end
  end
  -- whether this paragraph's break is a tracked insertion or deletion (in its mark's properties)
  local mark = vim.iter(pieces):find(function(piece) return piece.paragraph_mark end)
  local mark_rpr = ppr and xml.child(ppr, "w:rPr")
  for _, tag in ipairs({ "w:ins", "w:del" }) do
    local old_mark = mark_rpr and xml.child(mark_rpr, tag)
    if old_mark then
      xml.remove(old_mark)
    end
  end
  if mark then
    local ch = mark.paragraph_mark
    local tag = ch.type == "delete" and "w:del" or "w:ins"
    ch.mark = ch.mark or (ch.node and ch.node.tag == tag and #ch.node.children == 0 and ch.node) or revision(ctx, tag, ch)
    ppr = ppr or xml.elem("w:pPr")
    if not mark_rpr then -- (rPr comes last in pPr, except for these)
      local after = xml.child(ppr, "w:sectPr") or xml.child(ppr, "w:pPrChange")
      mark_rpr = xml.append(ppr, xml.elem("w:rPr"), after and xml.index(after) or nil)
    end
    xml.append(mark_rpr, xml.elem(tag, vim.deepcopy(ch.mark.attrs)), 1) -- (first, as the schema wants)
  end
  if block.align or (ppr and xml.child(ppr, "w:jc")) then
    ppr = ppr or xml.elem("w:pPr")
    local style = xml.child(ppr, "w:pStyle")
    set_align(ppr, block.align, M.style_align(ctx, style and xml.attr(style, "w:val") or default_style_id(ctx)))
  end
  if block.para then -- (its own format's attributes: a plain paragraph's, see paras.lua)
    ppr = ppr or xml.elem("w:pPr")
    set_para(ctx, ppr, block.para)
  end
  if ppr then
    xml.append(p, ppr)
  end

  local base = M.font_base(ctx, ppr) -- (the font and size its style gives)
  local cur = { el = p } -- the open <w:hyperlink>/<w:ins> wrappers
  for _, piece in ipairs(pieces) do
    if piece.text then
      if piece.link ~= cur.link then
        cur = { link = piece.link, el = piece.link and xml.append(p, hyperlink(ctx, piece.link)) or p }
      end
      local ins = piece.insert
      if ins ~= cur.insert then
        cur.insert = ins
        if ins then
          ins.node = ins.node or revision(ctx, "w:ins", ins)
          cur.ins = xml.append(cur.el, xml.elem(ins.moved and "w:moveTo" or "w:ins", vim.deepcopy(ins.node.attrs)))
        else
          cur.ins = nil
        end
      end
      xml.append(cur.ins or cur.el, text_run(piece, ctx, base))
    elseif not (piece.insert_start or piece.insert_end) then
      cur = { el = p }
      local del = piece.deletion
      if del and (not del.node or piece.deleted ~= del.text) then -- (new, or its text changed: split, edited)
        local attrs = revision(ctx, "w:del", del).attrs -- (a fresh id: a deletion may now be in several paragraphs)
        del.node, del.text = xml.elem(del.moved and "w:moveFrom" or "w:del", attrs), piece.deleted
        for _, seg in ipairs(piece.segs or { { text = piece.deleted } }) do -- (formatted as it was)
          local run = text_run(vim.tbl_extend("force", seg, { text = (seg.text:gsub("\n", " ")) }))
          for _, el in ipairs(run.children) do
            el.tag = el.tag == "w:t" and "w:delText" or el.tag
          end
          xml.append(del.node, run)
        end
      end
      if piece.object and not piece.object.tag then -- (a field: its runs)
        for _, el in ipairs(piece.object.children) do
          xml.append(p, xml.copy(el))
        end
      elseif piece.object or del then
        if piece.note and not has(part(ctx, FOOTNOTES) or {}, "w:footnote", "w:id", piece.note.id) then
          -- (a footnote deleted by a save and brought back by undo: written again)
          local notes = xml.find(part(ctx, FOOTNOTES, FOOTNOTES_XML), "w:footnotes")
          local model = vim.iter(xml.collect(notes, { ["w:footnote"] = true })):find(function(el)
            return (tonumber(xml.attr(el, "w:id")) or 0) > 0
          end)
          footnote_body(xml.append(notes, xml.elem("w:footnote", { { "w:id", piece.note.id } })), piece.note.text, model)
          ctx.footnote_ids[piece.note.id], ctx.dirty[FOOTNOTES] = true, true
        end
        xml.append(p, piece.object or del.node)
      elseif piece.line_break or piece.page_break then
        xml.append(p, xml.elem("w:r", {}, { xml.elem("w:br", piece.page_break and { { "w:type", "page" } } or nil) }))
      elseif piece.footnote then
        piece.footnote.node = footnote_node(ctx, piece.footnote)
        xml.append(p, piece.footnote.node)
      elseif piece.image then
        xml.append(p, image_node(ctx, piece.image))
      elseif piece.math then
        xml.append(p, piece.math.node)
      elseif piece.comment_start then
        xml.append(p, xml.elem("w:commentRangeStart", { { "w:id", comment_id(ctx, piece.comment_start) } }))
      end
      local c = piece.comment_end or (piece.point and piece.comment_start)
      if c then
        xml.append(p, xml.elem("w:commentRangeEnd", { { "w:id", comment_id(ctx, c) } }))
        for _, x in ipairs({ c, unpack(c.replies) }) do
          if not held[comment_id(ctx, x)] then
            local ref = xml.elem("w:commentReference", { { "w:id", comment_id(ctx, x) } })
            xml.append(p, xml.elem("w:r", {}, { ref }))
          end
        end
      end
    end
  end
end

-- ── Writing comments ────────────────────────────────────────────────────────

-- A new paragraph id (8 hex digits, below 0x80000000) for comments.xml.
local function para_id(ctx)
  ctx.para_ids = ctx.para_ids or {}
  local id
  repeat
    id = ("%08X"):format(math.random(0x10000000, 0x7FFFFFFE))
  until not ctx.para_ids[id]
  ctx.para_ids[id] = true
  return id
end

local function comment_el(x)
  local el = xml.elem("w:comment", {
    { "w:id", x.id }, { "w:author", x.author }, { "w:date", x.date }, { "w:initials", vim.fn.strcharpart(x.author, 0, 2) },
  })
  local lines = vim.split(x.text, "\n")
  for k, line in ipairs(lines) do
    local attrs = k == #lines and { { "w14:paraId", x.para } } or {}
    xml.append(el, xml.elem("w:p", attrs, { xml.elem("w:r", {}, { xml.elem("w:t", { { "xml:space", "preserve" } }, { { text = line } }) }) }))
  end
  return el
end

-- An edited comment's text, in its element: its first paragraph's properties and its reference
-- mark (w:annotationRef) kept, its last paragraph's id (threads are linked by it) too.
local function set_comment_text(x)
  local paras = vim.tbl_filter(function(c) return c.tag == "w:p" end, x.node.children)
  local first = paras[1] or xml.elem("w:p")
  local ppr = xml.child(first, "w:pPr")
  local ref = vim.iter(xml.collect(first, { ["w:r"] = true })):find(function(r) return xml.child(r, "w:annotationRef") ~= nil end)
  local rpr = vim.iter(xml.collect(first, { ["w:r"] = true })):map(function(r) return not xml.child(r, "w:annotationRef") and xml.child(r, "w:rPr") or nil end):next()
  for _, p in ipairs(paras) do
    xml.remove(p)
  end
  local lines = vim.split(x.text, "\n")
  for k, line in ipairs(lines) do
    local p = xml.elem("w:p", k == #lines and x.para and { { "w14:paraId", x.para } } or {})
    if ppr then
      xml.append(p, xml.copy(ppr))
    end
    if k == 1 and ref then
      xml.append(p, xml.copy(ref))
    end
    xml.append(p, xml.elem("w:r", {}, { rpr and xml.copy(rpr), xml.elem("w:t", { { "xml:space", "preserve" } }, { { text = line } }) }))
    xml.append(x.node, p)
  end
  x.was_text = x.text
end

-- A deleted comment (or reply) leaves every part that mentions it.
local function forget(ctx, x)
  xml.remove(x.node)
  ctx.dirty[COMMENTS] = true
  if x.ex and x.ex.parent then
    xml.remove(x.ex)
    ctx.dirty[EXTENDED] = true
  end
  local ids = x.para and part(ctx, IDS)
  local entry = ids and has(ids, "w16cid:commentId", "w16cid:paraId", x.para)
  if entry then
    xml.remove(entry)
    ctx.dirty[IDS] = true
    local extensible = part(ctx, EXTENSIBLE)
    local durable = xml.attr(entry, "w16cid:durableId")
    local ext_entry = extensible and durable and has(extensible, "w16cex:commentExtensible", "w16cex:durableId", durable)
    if ext_entry then
      xml.remove(ext_entry)
      ctx.dirty[EXTENSIBLE] = true
    end
  end
end

-- Bring comments.xml and its companions in line with the buffer's comments, touching only the
-- comments that were added, replied to, resolved or deleted.
local function save_comments(ctx, live)
  if not ctx.comments_dirty then
    return
  end
  local keep = {}
  for _, c in ipairs(live) do
    for _, x in ipairs({ c, unpack(c.replies) }) do
      keep[x] = true
    end
  end
  for el, x in pairs(ctx.comment_of) do
    if not keep[x] and el.parent then
      forget(ctx, x)
    end
  end

  for _, c in ipairs(live) do
    for _, x in ipairs({ c, unpack(c.replies) }) do -- (edited: its text written again)
      if x.node and x.was_text ~= nil and x.text ~= x.was_text then
        set_comment_text(x)
        ctx.dirty[COMMENTS] = true
      end
    end
    local new_replies = vim.tbl_filter(function(x) return not x.node end, c.replies)
    if not c.node or #new_replies > 0 or (not c.resolved) ~= (not c.was_resolved) then
      local comments = xml.find(part(ctx, COMMENTS, '<w:comments xmlns:w="' .. NS.w .. '"/>'), "w:comments")
      local extended = xml.find(part(ctx, EXTENDED, '<w15:commentsEx xmlns:w15="' .. NS.w15 .. '"/>'), "w15:commentsEx")
      ensure_ns(comments, "w14")
      local thread = c.node and { c, unpack(new_replies) } or { c, unpack(c.replies) }
      for _, x in ipairs(thread) do
        if not x.para then -- (threads are linked by paragraph ids; older documents may lack them)
          x.para = para_id(ctx)
          if x.node then
            local paras = xml.collect(x.node, { ["w:p"] = true })
            xml.set_attr(paras[#paras], "w14:paraId", x.para)
          end
        end
        if not x.node then
          comment_id(ctx, x)
          x.node = xml.append(comments, comment_el(x))
          ctx.comment_of[x.node] = x
        end
        if not x.ex then
          local attrs = { { "w15:paraId", x.para } }
          if x ~= c then
            table.insert(attrs, { "w15:paraIdParent", c.para })
          end
          x.ex = xml.append(extended, xml.elem("w15:commentEx", attrs))
        end
      end
      if (not c.resolved) ~= (not c.was_resolved) then -- (on the thread's first comment, as Word does)
        for _, x in ipairs({ c, unpack(c.replies) }) do
          if x.ex then
            xml.set_attr(x.ex, "w15:done", c.resolved and x == c and "1" or "0")
          end
        end
        c.was_resolved = c.resolved
      end
      ctx.dirty[COMMENTS], ctx.dirty[EXTENDED] = true, true
    end
  end
  -- a document without comments until now needs the parts registered
  for name, rel_type in pairs({ comments = REL .. "comments", commentsExtended = REL_EX }) do
    local file = "word/" .. name .. ".xml"
    if ctx.dirty[file] and not ctx.entries[file] then
      content_type(ctx, "Override", "PartName", "/" .. file, MIME .. name .. "+xml")
      add_rel(ctx, rel_type, name .. ".xml")
      ctx.entries[file] = true
    end
  end
end

-- Formatting changes the text doesn't show (of paragraph marks, of runs inside fields...), so not
-- changes of their own: accepted or rejected with all the changes (of an author). Those in the
-- document, but those of `skip` (ids: shown as changes), as { id, author }.
function M.mark_changes(ctx, skip)
  local out = {}
  for _, change in ipairs(xml.collect(ctx.root, { ["w:rPrChange"] = true })) do
    local id = xml.attr(change, "w:id")
    if id and not (skip or {})[id] then
      out[#out + 1] = { id = id, author = xml.attr(change, "w:author") or "?" }
    end
  end
  return out
end

-- decisions: { action, ids } (the changes there were when it was made), in order: the last one
-- about a change decides
local function resolve_marks(ctx, decisions)
  for _, change in ipairs(#decisions > 0 and xml.collect(ctx.root, { ["w:rPrChange"] = true }) or {}) do
    local id = xml.attr(change, "w:id")
    for k = #decisions, 1, -1 do
      local d = decisions[k]
      if id and d.ids[id] then
        local rpr = change.parent
        if d.action == "reject" then -- (the formatting from before)
          local old = xml.child(change, "w:rPr")
          rpr.children = {}
          for _, c in ipairs(old and old.children or {}) do
            xml.append(rpr, xml.copy(c))
          end
        else
          xml.remove(change)
        end
        ctx.dirty[M.MAIN] = true
        break
      end
    end
  end
end

function M.save(ctx, live)
  resolve_marks(ctx, live.mark_decisions or {})
  save_comments(ctx, live.comments)
  -- footnotes whose reference was deleted go too
  local notes = ctx.dirty[M.MAIN] and part(ctx, FOOTNOTES)
  if notes then
    local used = {}
    for _, ref in ipairs(xml.collect(ctx.root, { ["w:footnoteReference"] = true })) do
      used[xml.attr(ref, "w:id")] = true
    end
    for id in pairs(ctx.footnote_ids) do
      local note = not used[id] and has(notes, "w:footnote", "w:id", id)
      if note then
        xml.remove(note)
        ctx.footnote_ids[id], ctx.dirty[FOOTNOTES] = nil, true
      end
    end
  end
  local entries, trees = {}, {}
  for name in pairs(ctx.dirty) do
    entries[name], trees[name] = xml.serialize(ctx.parts[name]), ctx.parts[name]
  end
  for name, bytes in pairs(ctx.files) do
    entries[name] = bytes
  end
  return entries, trees -- (trees: the XML each part was written from, to check: see init.lua)
end

return M
