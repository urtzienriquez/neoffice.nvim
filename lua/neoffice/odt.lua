-- .odt: content.xml (+ images) <-> paragraph spans.
local xml = require("neoffice.xml")
local markup = require("neoffice.markup")
local zip = require("neoffice.zip")
local images = require("neoffice.images")
local M = {}

M.MAIN = "content.xml"
local MANIFEST, STYLES = "META-INF/manifest.xml", "styles.xml"
local LISTS = { ["text:list"] = true, ["text:list-item"] = true, ["text:list-header"] = true }

local PARA = { ["text:p"] = true, ["text:h"] = true }
-- Elements shown as a "￼" placeholder even without text (others without text stay invisible).
local VISIBLE = { ["draw:frame"] = true, ["draw:a"] = true, ["text:line-break"] = true }

local function child_text(el, tag)
  return xml.text(xml.child(el, tag) or {})
end

local FLAGS = { "b", "i", "u", "s", "h", "sup", "sub" } -- (see markup.FLAGS)
local ALIGN = { center = "center", ["end"] = "right", right = "right", justify = "justify" } -- fo:text-align
local QUOTES = { Quotations = true, Block_20_Text = true, Quote = true, Intense_20_Quote = true }
local CAPTIONS = { Caption = true, Table = true, Figure = true, Illustration = true, Drawing = true }
local caption_style, para_attrs, style_attrs, automatic -- (below)

-- The formatting { b, i, u, sup, sub } of a text style, following its parent styles, starting
-- from `inherited` (that of the enclosing text).
local function merge(a, b) -- (a copy of a, with b's: tables of plain values)
  local out = {}
  for k, v in pairs(a) do
    out[k] = v
  end
  for k, v in pairs(b) do
    out[k] = v
  end
  return out
end
local function flags(ctx, name, inherited)
  local f = merge(inherited or {}, {}) -- (a copy: its values aren't changed, only replaced)
  for _, k in ipairs(FLAGS) do
    f[k] = f[k] or false
  end
  local style = name and ctx.styles[name]
  if not style then
    return f
  end
  f = flags(ctx, xml.attr(style, "style:parent-style-name"), f)
  local props = xml.child(style, "style:text-properties") or {}
  local weight, slant = xml.attr(props, "fo:font-weight"), xml.attr(props, "fo:font-style")
  local underline, position = xml.attr(props, "style:text-underline-style"), xml.attr(props, "style:text-position")
  local strike, background = xml.attr(props, "style:text-line-through-style"), xml.attr(props, "fo:background-color")
  if strike then
    f.s = strike ~= "none"
  end
  if background then
    f.h = background ~= "transparent" and (markup.color_of_rgb(background) or "yellow")
  end
  if weight then
    f.b = weight == "bold" or (tonumber(weight) or 0) >= 600
  end
  if slant then
    f.i = slant ~= "normal"
  end
  if underline then
    f.u = underline ~= "none"
  end
  if position then -- "super 58%", "sub 58%", "33% 58%" (raised), "-33% 58%" (lowered), "0% 100%"
    local raise = position:match("^super") and 1 or position:match("^sub") and -1 or tonumber(position:match("^(-?[%d.]+)%%")) or 0
    f.sup, f.sub = raise > 0, raise < 0
  end
  return f
end

-- ── Reading ─────────────────────────────────────────────────────────────────

local function annotation(el)
  local lines = {}
  for _, p in ipairs(xml.collect(el, PARA)) do
    lines[#lines + 1] = xml.text(p)
  end
  return {
    name = xml.attr(el, "office:name"),
    author = child_text(el, "dc:creator"),
    date = child_text(el, "dc:date"),
    text = table.concat(lines, "\n"),
    resolved = xml.attr(el, "loext:resolved") == "true",
    replies = {},
    node = el, -- (written back as it is, unless you resolve or reopen it)
  }
end

-- An image: its file in the package, its alt text (description, or else title) and its place.
local function image_of(ctx, frame)
  if xml.child(frame, "draw:object") then -- a formula or chart (its draw:image is only a preview)
    return nil
  end
  local image = xml.child(frame, "draw:image")
  local path = image and xml.attr(image, "xlink:href")
  if path and not path:match("^%a+:") and not path:match("^%.%./") then
    local alt = xml.child(frame, "svg:desc") or xml.child(frame, "svg:title")
    local cm = markup.cm(xml.attr(frame, "svg:width"))
    local out = { path = path, alt = alt and vim.trim(xml.text(alt):gsub("%s+", " ")) or "", width = cm and markup.width(cm) or nil }
    for key, value in pairs(require("neoffice.floats").classify(M.frame_geometry(ctx, frame), M.page(ctx))) do
      out[key] = value
    end
    for key, value in pairs(M.frame_crop(ctx, frame, path)) do
      out[key] = value
    end
    return out
  end
end

local function element_children(el)
  return vim.tbl_filter(function(c) return c.tag end, el.children)
end

-- Whether level `level` of a list style has bullets or numbers.
local function list_kind(ctx, style_name, level)
  local style = style_name and ctx.list_styles[style_name]
  local levels = style and element_children(style) or {}
  local level_style = vim.iter(levels):find(function(c) return xml.attr(c, "text:level") == tostring(level) end)
  return level_style and level_style.tag == "text:list-level-style-number" and xml.attr(level_style, "style:num-format") ~= "" and "number" or "bullet"
end

-- A numbered level's pattern (see markup.list_marker): the numbers it shows (its parents', as
-- display-levels says, then its own), each in its level's format, and its suffix; "1." if it isn't
-- one that can be written so (a prefix, "(1)"; no number).
local ODF_SAMPLE = { ["1"] = "1", a = "a", A = "A", i = "i", I = "I" }
local function level_style(ctx, style_name, level)
  local style = style_name and ctx.list_styles[style_name]
  for _, c in ipairs(style and element_children(style) or {}) do
    if xml.attr(c, "text:level") == tostring(level) then
      return c
    end
  end
end
local function list_pattern(ctx, style_name, level)
  local own = level_style(ctx, style_name, level)
  local fmt = own and xml.attr(own, "style:num-format")
  if not own or own.tag ~= "text:list-level-style-number" or not ODF_SAMPLE[fmt or ""] or (xml.attr(own, "style:num-prefix") or "") ~= "" then
    return "1."
  end
  local display = tonumber(xml.attr(own, "text:display-levels") or "1") or 1
  local parts = {}
  for l = level - display + 1, level - 1 do
    local parent = level_style(ctx, style_name, l)
    local pf = parent and xml.attr(parent, "style:num-format")
    if not ODF_SAMPLE[pf or ""] then
      return "1."
    end
    parts[#parts + 1] = ODF_SAMPLE[pf] .. "."
  end
  local pattern = table.concat(parts) .. ODF_SAMPLE[fmt] .. (xml.attr(own, "style:num-suffix") or "")
  return markup.list_marker(pattern .. " x") == pattern and pattern or "1."
end

-- The list a paragraph is an item of: { kind = "bullet" | "number", level = 1.., id = its outermost
-- list, pattern (numbered: see list_pattern), style (its list style's name) }.
local function list_of(ctx, p)
  local level, style, top, node = 0, nil, nil, p.parent
  while node and LISTS[node.tag] do
    if node.tag == "text:list-header" then -- (an unnumbered heading of the list)
      return nil
    elseif node.tag == "text:list-item" and level == 0 then -- (its own list style, as LibreOffice writes one)
      style = xml.attr(node, "text:style-override") or style
    elseif node.tag == "text:list" then
      level, top = level + 1, node
      style = style or xml.attr(node, "text:style-name")
    end
    node = node.parent
  end
  if not top then
    return nil
  end
  local kind = list_kind(ctx, style, level)
  return { kind = kind, level = level, id = top, style = style, pattern = kind == "number" and list_pattern(ctx, style, level) or nil }
end

-- A list's level numbered as a pattern says ("1.a.": its own number in letters, after its parent's).
-- Only where it is: the list it is in (its sub-list, for a sub-item: the sub-items under another
-- item may be numbered their own way) gets a copy of its list style.
local function set_pattern(ctx, p, pattern)
  local list = list_of(ctx, p)
  if not list or list.kind ~= "number" or list.pattern == pattern then
    return
  end
  local target = list.level > 1 and p.parent and p.parent.parent or list.id -- (its sub-list, or its list)
  local original = list.style and ctx.list_styles[list.style]
  if not original then
    return
  end
  local copy = xml.copy(original)
  local name, n = nil, 1
  repeat
    name, n = "NeoList" .. n, n + 1
  until not ctx.list_styles[name]
  xml.set_attr(copy, "style:name", name)
  xml.set_attr(copy, "style:display-name", nil)
  xml.append(automatic(ctx), copy)
  ctx.list_styles[name] = copy
  xml.set_attr(target, "text:style-name", name)
  for _, l in ipairs(xml.collect(target, { ["text:list"] = true })) do -- (the lists in it that said the same)
    if l ~= target and xml.attr(l, "text:style-name") == list.style then
      xml.set_attr(l, "text:style-name", name)
    end
  end
  local own = level_style(ctx, name, list.level)
  if not own or own.tag ~= "text:list-level-style-number" then
    return
  end
  local parts = {}
  for part in pattern:gmatch("[%w]+[.)]") do
    parts[#parts + 1] = part
  end
  xml.set_attr(own, "style:num-format", parts[#parts]:sub(1, -2))
  xml.set_attr(own, "style:num-suffix", parts[#parts]:sub(-1))
  xml.set_attr(own, "style:num-prefix", nil)
  xml.set_attr(own, "text:display-levels", #parts > 1 and tostring(#parts) or nil)
end

local function based_on(ctx, p, names) -- in one of these styles, or in a style based on one
  local style = xml.attr(p, "text:style-name")
  for _ = 1, 10 do
    if names[style] then
      return true
    end
    style = style and ctx.parent_style[style]
  end
  return false
end

local function is_title(ctx, p)
  return based_on(ctx, p, { Title = true })
end

-- The format's attributes (see paras.lua) of paragraph properties.
-- (a paragraph style's, following the styles it is based on, and the default style)
local function raw_attrs(props)
  local paras, out = require("neoffice.paras"), {}
  if not props then
    return out
  end
  local function cm(attr)
    local value = xml.attr(props, attr)
    local sign, rest = (value or ""):match("^(%-?)(.*)$")
    local n = markup.cm(rest)
    return n and (sign == "-" and -n or n)
  end
  local left, right, first = cm("fo:margin-left"), cm("fo:margin-right"), cm("fo:text-indent")
  out["indent-left"] = left and left ~= 0 and paras.cm(left) or nil
  out["indent-right"] = right and right ~= 0 and paras.cm(right) or nil
  out.first = first and first ~= 0 and paras.cm(first) or nil
  local before, after = cm("fo:margin-top"), cm("fo:margin-bottom")
  out.before = before and paras.pt(before / 2.54 * 72) or nil
  out.after = after and paras.pt(after / 2.54 * 72) or nil
  local line = xml.attr(props, "fo:line-height") or xml.attr(props, "style:line-height-at-least")
  local percent = line and tonumber(line:match("^([%d.]+)%%$") or "")
  if percent then
    out.line = paras.multiple(percent / 100)
  elseif line and markup.cm(line) then
    out.line = paras.pt(markup.cm(line) / 2.54 * 72)
  end
  return out
end
style_attrs = function(ctx, name)
  if not ctx.para_style_el then
    ctx.para_style_el, ctx.style_attrs = {}, {}
    for _, root in ipairs({ ctx.styles_doc, ctx.root }) do
      for _, style in ipairs(xml.collect(root, { ["style:style"] = true, ["style:default-style"] = true }, { ["office:body"] = true })) do
        if xml.attr(style, "style:family") == "paragraph" then
          local name = xml.attr(style, "style:name")
          if style.tag == "style:default-style" then
            ctx.para_style_el[false] = style
          elseif name then
            ctx.para_style_el[name] = style
          end
        end
      end
    end
  end
  local key = name or false
  if not ctx.style_attrs[key] then
    local chain, cur = {}, name
    for _ = 1, 10 do
      if not cur or not ctx.para_style_el[cur] then
        break
      end
      chain[#chain + 1] = ctx.para_style_el[cur]
      cur = xml.attr(ctx.para_style_el[cur], "style:parent-style-name")
    end
    local default = ctx.para_style_el[false]
    local out = raw_attrs(default and xml.child(default, "style:paragraph-properties"))
    for k = #chain, 1, -1 do
      for attr, value in pairs(raw_attrs(xml.child(chain[k], "style:paragraph-properties"))) do
        out[attr] = value
      end
    end
    ctx.style_attrs[key] = out
  end
  return ctx.style_attrs[key]
end

-- A paragraph's own format's attributes (see paras.lua): those of its automatic style that the
-- style it is based on doesn't give it already.
para_attrs = function(ctx, p)
  local paras = require("neoffice.paras")
  local style = ctx.auto_paras[xml.attr(p, "text:style-name") or ""]
  local out = raw_attrs(style and xml.child(style, "style:paragraph-properties"))
  local base = style_attrs(ctx, style and xml.attr(style, "style:parent-style-name") or xml.attr(p, "text:style-name"))
  for _, key in ipairs(paras.LAYOUT) do
    if out[key] and paras.measure(key, out[key]) == paras.measure(key, base[key] or (key == "line" and "single" or "0cm")) then
      out[key] = nil
    end
  end
  return out
end

-- A paragraph's own alignment and whether it starts on a new page (from its automatic style).
local function layout(ctx, p)
  local style = ctx.auto_paras[xml.attr(p, "text:style-name") or ""]
  local props = style and xml.child(style, "style:paragraph-properties") or { attrs = {} }
  return ALIGN[xml.attr(props, "fo:text-align") or ""], xml.attr(props, "fo:break-before") == "page"
end

-- Fonts (see fonts_of in init.lua): a span's own font and size (its text styles', following their
-- parents), and what a paragraph's style gives (following its parents, and the default style).
-- A font by its family's name (style:font-name names a font-face, which says its family); sizes as
-- shown ("11pt").
local function unquote(s)
  return (s:gsub("^%s*['\"](.*)['\"]%s*$", "%1"))
end
local function first_family(family) -- (of a list, the first preferred: "'Times New Roman', serif")
  local first = vim.trim(unquote(family:match("^%s*('[^']*')") or family:match('^%s*("[^"]*")') or family:match("^([^,]*)")))
  return first ~= "" and first or nil
end
local function font_families(ctx)
  if not ctx.font_family then
    ctx.font_family = {}
    for _, root in ipairs({ ctx.styles_doc, ctx.root }) do
      for _, face in ipairs(xml.collect(root, { ["style:font-face"] = true }, { ["office:body"] = true })) do
        local name, family = xml.attr(face, "style:name"), xml.attr(face, "svg:font-family")
        if name then
          ctx.font_family[name] = family and first_family(family) or name
        end
      end
    end
  end
  return ctx.font_family
end
local function props_font(ctx, props, out) -- (text properties, over what `out` says)
  if props then
    local paras = require("neoffice.paras")
    local name, family = xml.attr(props, "style:font-name"), xml.attr(props, "fo:font-family")
    if name then
      out.font = font_families(ctx)[name] or name
    elseif family then
      out.font = first_family(family) or out.font
    end
    local size = xml.attr(props, "fo:font-size")
    local percent = size and tonumber(size:match("^([%d.]+)%%$") or "")
    if percent then -- (of what it had)
      local had = out.size and markup.cm(out.size)
      out.size = had and paras.pt(had * percent / 100 * 72 / 2.54) or out.size
    elseif size and markup.cm(size) then
      out.size = paras.pt(markup.cm(size) * 72 / 2.54)
    end
  end
  return out
end
local function text_style_font(ctx, name, out, depth)
  local style = name and ctx.styles[name]
  if not style or (depth or 0) > 10 then
    return out
  end
  out = text_style_font(ctx, xml.attr(style, "style:parent-style-name"), out, (depth or 0) + 1)
  return props_font(ctx, xml.child(style, "style:text-properties"), out)
end
function M.run_font(ctx, props) -- (props: a span's style, or the styles of the spans it is in)
  local chain = type(props) == "table" and props or { props }
  local key = table.concat(chain, "\0")
  ctx.run_fonts = ctx.run_fonts or {}
  if not ctx.run_fonts[key] then -- (a style's fonts don't change: new styles get new names)
    local out = {}
    for _, name in ipairs(chain) do
      out = text_style_font(ctx, name, out)
    end
    ctx.run_fonts[key] = out
  end
  return ctx.run_fonts[key] -- (as it is: not to be changed)
end
function M.font_default(ctx) -- (the document's: its default paragraph style's, Standard)
  return M.font_base(ctx, nil)
end
function M.font_base(ctx, p)
  local name = p and M.style_of(ctx, p) or "Standard" -- (its named style's: an automatic one is its own formatting)
  ctx.font_bases = ctx.font_bases or {}
  if not ctx.font_bases[name] then
    ctx.font_bases[name] = M.font_base_of(ctx, name)
  end
  return ctx.font_bases[name] -- (as it is: not to be changed)
end
-- (what its text shows where it says nothing: its automatic style's font, if that has one, else
-- its named style's, as LibreOffice does)
function M.font_inherited(ctx, p)
  local name = xml.attr(p, "text:style-name") or "Standard"
  ctx.font_inherits = ctx.font_inherits or {}
  if not ctx.font_inherits[name] then
    ctx.font_inherits[name] = M.font_base_of(ctx, name)
  end
  return ctx.font_inherits[name] -- (as it is: not to be changed)
end
function M.font_base_of(ctx, name)
  style_attrs(ctx, nil) -- (the paragraph styles, by name)
  local chain, cur = {}, name
  for _ = 1, 10 do
    if not cur or not ctx.para_style_el[cur] then
      break
    end
    chain[#chain + 1] = ctx.para_style_el[cur]
    cur = xml.attr(ctx.para_style_el[cur], "style:parent-style-name")
  end
  local default = ctx.para_style_el[false]
  local out = props_font(ctx, default and xml.child(default, "style:text-properties"), {})
  for k = #chain, 1, -1 do
    out = props_font(ctx, xml.child(chain[k], "style:text-properties"), out)
  end
  return out
end
-- (a font-face for a font, so a style can name it: the document's own if it has one)
local function font_face(ctx, family, in_styles) -- (in content.xml, or in styles.xml for a style there)
  local root
  if in_styles then
    root = xml.find(ctx.styles_doc, "office:document-styles")
  else
    root = automatic(ctx).parent -- (<office:document-content>)
  end
  local decls = xml.child(root, "office:font-face-decls")
  for _, face in ipairs(decls and xml.collect(decls, { ["style:font-face"] = true }) or {}) do
    local name = xml.attr(face, "style:name")
    if name and (font_families(ctx)[name] or name):lower() == family:lower() then
      return name
    end
  end
  if not decls then -- (before the styles)
    local first = xml.child(root, "office:styles") or xml.child(root, "office:automatic-styles")
    decls = xml.append(root, xml.elem("office:font-face-decls"), first and xml.index(first) or nil)
  end
  local name = family -- (its name: the family's, unless another font has it)
  if font_families(ctx)[name] and font_families(ctx)[name]:lower() ~= family:lower() then
    name = family .. " (neoffice)"
  end
  xml.append(decls, xml.elem("style:font-face", { { "style:name", name },
    { "svg:font-family", family:find("[%s,;]") and "'" .. family .. "'" or family } }))
  ctx.font_family[name] = family
  return name
end
M.font_face = font_face

-- A paragraph's named style (not the automatic one LibreOffice gives its own formatting).
function M.style_of(ctx, p)
  local name = xml.attr(p, "text:style-name")
  local auto = name and ctx.auto_paras[name]
  if auto then
    name = xml.attr(auto, "style:parent-style-name")
  end
  return name or "Standard"
end
local function text_align(props) -- (fo:text-align, as a name: "left" for start/left)
  local v = props and xml.attr(props, "fo:text-align")
  return v and (ALIGN[v] or "left") or nil
end
function M.style_align(ctx, name) -- (what alignment a style gives: nil for left)
  ctx.style_aligns = ctx.style_aligns or {}
  local key = name or false
  if ctx.style_aligns[key] == nil then
    style_attrs(ctx, nil)
    local found, cur = nil, name
    for _ = 1, 10 do
      local el = cur and ctx.para_style_el[cur]
      if not el then
        break
      end
      found = found or text_align(xml.child(el, "style:paragraph-properties"))
      cur = xml.attr(el, "style:parent-style-name")
    end
    local default = ctx.para_style_el[false]
    found = found or text_align(default and xml.child(default, "style:paragraph-properties"))
    ctx.style_aligns[key] = found ~= "left" and found or false
  end
  return ctx.style_aligns[key] or nil
end
local function own_align(ctx, p) -- (its alignment, if not its style's: "left" too, then)
  local auto = ctx.auto_paras[xml.attr(p, "text:style-name") or ""]
  local base = M.style_align(ctx, M.style_of(ctx, p)) or "left"
  local have = text_align(auto and xml.child(auto, "style:paragraph-properties")) or base
  return have ~= base and have or nil
end

-- The block of a paragraph (see markup.prefix).
local function block_of(ctx, p)
  local list, block = list_of(ctx, p), {}
  if p.tag == "text:h" then
    block = { heading = tonumber(xml.attr(p, "text:outline-level")) or 1 }
  elseif is_title(ctx, p) then
    block = { title = true }
  elseif based_on(ctx, p, QUOTES) then
    block = { quote = true }
  elseif list then
    block = { list = list }
  end
  block.align = own_align(ctx, p)
  block.caption_style = based_on(ctx, p, CAPTIONS) or nil -- (see captions in init.lua)
  block.font_base, block.font_inherit, block.font_default = M.font_base(ctx, p), M.font_inherited(ctx, p), M.font_default(ctx)
  if not (block.heading or block.title or block.quote or block.list) then -- (a plain one: its own format, see paras.lua)
    block.para = para_attrs(ctx, p)
  end
  return block
end

-- The text of an element as one line (spaces, tabs; not a footnote's number).
local function plain(el)
  local out = {}
  local function walk(node)
    for _, c in ipairs(node.children) do
      if c.text then
        out[#out + 1] = (c.text:gsub("%s+", " "))
      elseif c.tag == "text:s" then
        out[#out + 1] = (" "):rep(tonumber(xml.attr(c, "text:c")) or 1)
      elseif c.tag == "text:tab" then
        out[#out + 1] = "\t"
      elseif c.tag and c.tag ~= "text:note-citation" then
        walk(c)
        out[#out + 1] = PARA[c.tag] and " " or nil
      end
    end
  end
  walk(el)
  return vim.trim(table.concat(out))
end

local function read_para(ctx, p)
  local spans = {}
  local function add(span)
    spans[#spans + 1] = span
  end
  local function change(el)
    local id = xml.attr(el, "text:change-id")
    ctx.changes[id] = ctx.changes[id] or { id = id, type = "insert", author = "?", date = "" }
    return ctx.changes[id]
  end

  if select(2, layout(ctx, p)) then -- (it starts on a new page)
    add({ page_break = true })
  end
  local after_space = true -- (whitespace at the start of a paragraph is ignored, and after a space)
  -- f = { style, link, b, i, u } inherited from the enclosing spans/links; style: the span's style,
  -- or the styles of the spans it is in, outermost first, when they are nested
  local function walk(node, f)
    local function text(s)
      if s ~= "" then
        local own = M.run_font(ctx, f.style)
        add(merge(flags(ctx, nil, f), { text = s, props = f.style, link = f.link, run_font = own.font, run_size = own.size }))
      end
    end
    local k = 0
    while k < #node.children do
      k = k + 1
      local el = node.children[k]
      local tag = el.tag
      if tag == "field:fieldmark-start" then -- a field (a citation...): one thing, showing its result
        local group, depth = { children = {} }, 0
        for j = k, #node.children do
          local t = node.children[j].tag
          depth = depth + (t == "field:fieldmark-start" and 1 or t == "field:fieldmark-end" and -1 or 0)
          xml.append(group, xml.copy(node.children[j]))
          k = j
          if depth == 0 then
            break
          end
        end
        add({ object = group, visible = true, field = { text = plain(group) } })
        after_space = false
      elseif tag == "text:reference-mark-start" and (xml.attr(el, "text:name") or ""):match("^ZOTERO_") and (function()
        for j = k + 1, #node.children do -- (its end, beside it: a citation is one thing, as in Word)
          local c = node.children[j]
          if c.tag == "text:reference-mark-end" and xml.attr(c, "text:name") == xml.attr(el, "text:name") then
            return true
          end
        end
      end)() then
        local group = { children = {} }
        for j = k, #node.children do
          local c = node.children[j]
          xml.append(group, xml.copy(c))
          k = j
          if c.tag == "text:reference-mark-end" and xml.attr(c, "text:name") == xml.attr(el, "text:name") then
            break
          end
        end
        add({ object = group, visible = true, field = { text = plain(group) } })
        after_space = false
      elseif el.text then -- ODF collapses whitespace, also across elements; real spaces are <text:s/>
        local s = el.text:gsub("%s+", " ")
        text(after_space and s:gsub("^ ", "") or s)
        after_space = s:sub(-1) == " " or (after_space and s == "")
      elseif tag == "text:span" then
        local style = xml.attr(el, "text:style-name")
        local outer = type(f.style) == "table" and f.style or { f.style }
        local chain = style and vim.list_extend(vim.list_slice(outer), { style }) or outer
        walk(el, merge(flags(ctx, style, f), { style = #chain > 1 and chain or chain[1], link = f.link }))
      elseif tag == "text:a" then
        local url = xml.attr(el, "xlink:href")
        ctx.link_attrs[url] = ctx.link_attrs[url] or el.attrs
        walk(el, merge(f, { link = url }))
      elseif tag == "text:s" then
        text((" "):rep(tonumber(xml.attr(el, "text:c")) or 1))
        spans[#spans].explicit, after_space = true, false -- (a real space: kept)
      elseif tag == "text:tab" then
        text("\t")
        after_space = false
      elseif tag == "text:line-break" then
        add({ line_break = true })
        after_space = false
      elseif tag == "text:note" and xml.attr(el, "text:note-class") ~= "endnote" then
        add({ object = el, visible = true, footnote = { text = plain(xml.child(el, "text:note-body") or el) } })
        after_space = false
      elseif tag == "text:change-start" then
        add({ insert_start = change(el) })
      elseif tag == "text:change-end" then
        add({ insert_end = change(el) })
      elseif tag == "text:change" then
        add({ deletion = change(el) })
      elseif tag == "office:annotation" then
        local c, parent = annotation(el), xml.attr(el, "loext:parent-name")
        if parent and ctx.comments[parent] then
          table.insert(ctx.comments[parent].replies, c)
        else
          add({ comment_start = c })
          if c.name then
            ctx.comments[c.name] = c
          else
            add({ comment_end = c })
          end
        end
      elseif tag == "office:annotation-end" then
        local c = ctx.comments[xml.attr(el, "office:name")]
        if c then
          add({ comment_end = c })
        end
      elseif (tag == "draw:frame" or tag == "draw:custom-shape") and M.box_content(el) then -- a box: its paragraphs show before this one's
        add({ object = el, visible = true, box = { node = el, content = M.box_content(el), attrs = M.box_attrs(ctx, el) } })
        after_space = false
      elseif tag then
        local image = tag == "draw:frame" and image_of(ctx, el) or nil
        local text = not tag:match("^draw:") and plain(el) or ""
        local seq = tag == "text:sequence" and xml.attr(el, "text:name") or nil -- (a caption's number)
        local kind = not image and M.object_kind(ctx, el) -- (shown by name: ⟦chart⟧, kept)
        add({ object = el, visible = VISIBLE[tag] or xml.text(el) ~= "" or kind and true, image = image,
          field = kind and { text = kind } or text ~= "" and { text = text, seq = seq } or nil })
        after_space = after_space and not spans[#spans].visible
      end
    end
  end

  walk(p, flags(ctx, nil, {}))
  for k = #spans, 1, -1 do -- (and at its end)
    local sp = spans[k]
    if sp.text and not sp.explicit then
      sp.text = sp.text:gsub(" $", "")
      if sp.text ~= "" then
        break
      end
    elseif sp.object and sp.visible then
      break
    end
  end
  return spans, block_of(ctx, p)
end

function M.load(path)
  local ctx = {
    path = path,
    entries = zip.list(path), -- names of all files in the package
    root = xml.parse(assert(zip.read(path, M.MAIN), "no content.xml")),
    link_attrs = {}, -- url -> attributes of its first <text:a>
    styles = {}, -- text style name -> <style:style>
    headings = {}, -- heading level -> paragraph style name
    parent_style = {}, -- paragraph style name -> the style it is based on
    list_styles = {}, -- list style name -> <text:list-style>
    changes = {}, -- tracked change id -> change
    comments = {}, -- annotation name -> comment
    made = {}, -- styles created while saving, by "base|b|i|u"
    auto_paras = {}, -- automatic paragraph style name -> <style:style> (direct paragraph formatting)
  }
  ctx.auto = xml.find(ctx.root, "office:automatic-styles")
  ctx.styles_doc = xml.parse(zip.read(path, STYLES) or "")
  for _, root in ipairs({ ctx.styles_doc, ctx.root }) do
    for _, list_style in ipairs(xml.collect(root, { ["text:list-style"] = true }, { ["office:body"] = true })) do
      ctx.list_styles[xml.attr(list_style, "style:name")] = list_style
    end
  end
  local outline = xml.find(ctx.styles_doc, "text:outline-style")
  local first = outline and element_children(outline)[1]
  ctx.numbered = first ~= nil and (xml.attr(first, "style:num-format") or "") ~= ""
  for _, root in ipairs({ ctx.styles_doc, ctx.root }) do
    for _, style in ipairs(xml.collect(root, { ["style:style"] = true }, { ["office:body"] = true })) do
      local family, name = xml.attr(style, "style:family"), xml.attr(style, "style:name")
      local level = tonumber(xml.attr(style, "style:default-outline-level") or "")
      if family == "text" then
        ctx.styles[name] = style
      elseif family == "paragraph" and level and root ~= ctx.root then
        ctx.headings[level] = ctx.headings[level] or name
      end
      if family == "paragraph" then
        ctx.parent_style[name] = xml.attr(style, "style:parent-style-name") or false
        ctx.auto_paras[name] = root == ctx.root and style or nil
      end
      if family == "paragraph" and QUOTES[name] and root ~= ctx.root then -- (LibreOffice's is "Quotations")
        ctx.quote_style = name == "Quotations" and name or ctx.quote_style or name
      end
    end
  end
  local tracked = xml.find(ctx.root, "text:tracked-changes")
  ctx.tracking = tracked ~= nil and xml.attr(tracked, "text:track-changes") ~= "false"
  for _, region in ipairs(xml.collect(ctx.root, { ["text:changed-region"] = true })) do
    local kind = vim.iter(region.children):find(function(c)
      return c.tag
    end)
    local info = xml.find(kind, "office:change-info") or {}
    local deleted = {}
    for _, p in ipairs(xml.collect(kind, PARA)) do
      deleted[#deleted + 1] = xml.text(p)
    end
    local id = xml.attr(region, "text:id") or xml.attr(region, "xml:id")
    ctx.changes[id] = {
      id = id,
      region = region,
      type = ({ ["text:insertion"] = "insert", ["text:deletion"] = "delete" })[kind.tag] or "format",
      author = child_text(info, "dc:creator"),
      date = child_text(info, "dc:date"),
      text = table.concat(deleted, "\n"),
    }
  end
  local paras = {}
  local body = assert(xml.find(ctx.root, "office:text"), "no office:text")
  for _, p in ipairs(xml.collect(body, PARA, { ["text:tracked-changes"] = true })) do
    local spans, block = read_para(ctx, p)
    -- its boxes: their paragraphs first (not in a table's cell: there, a box stays as it is)
    local in_cell = false
    local up = p.parent
    while up do
      in_cell = in_cell or up.tag == "table:table-cell"
      up = up.parent
    end
    local own = {}
    for _, s in ipairs(spans) do
      if s.box and not in_cell then
        for _, bp in ipairs(xml.collect(s.box.content, PARA)) do
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
  for _, c in pairs(ctx.comments) do
    c.was_resolved = c.resolved
  end
  -- (or the headings are numbered by being in a list)
  ctx.numbered = ctx.numbered or vim.iter(paras):any(function(para) return para.node.tag == "text:h" and list_of(ctx, para.node) end)
  -- (which items one after another are one list: those of one list, also when LibreOffice wrote it
  -- as lists one after another going on with the numbering, "continue-numbering" or "continue-list")
  local prev_top, group, groups = nil, nil, 0
  for _, para in ipairs(paras) do
    local l = para.block.list
    -- (an item numbered again from a value, not the list's first: another list from there)
    local item = l and para.node and para.node.parent
    local restarts = l and l.level == 1 and item and item.tag == "text:list-item" and xml.attr(item, "text:start-value") ~= nil
      and element_children(l.id)[1] ~= item
    if restarts then
      groups = groups + 1
      group = groups
    end
    if l and l.id then
      if l.id ~= prev_top then
        local goes_on = xml.attr(l.id, "text:continue-numbering") == "true" or xml.attr(l.id, "text:continue-list") ~= nil
        if not (group and goes_on) then
          groups = groups + 1
          group = groups
        end
        prev_top = l.id
      end
      l.group = group
    else
      group, prev_top = nil, nil -- (a paragraph between: what follows is another)
    end
  end
  return ctx, paras
end

-- ── Writing ─────────────────────────────────────────────────────────────────

function M.new_para(template, ctx)
  local attrs = vim.tbl_filter(function(kv)
    return kv[1] ~= "xml:id" and kv[1] ~= "text:id"
  end, vim.deepcopy(template.attrs))
  local p = xml.elem(template.tag, attrs)
  -- not a style that starts a new page (or page style): the paragraph it comes after has it
  local auto = ctx and ctx.auto_paras[xml.attr(p, "text:style-name") or ""]
  local props = auto and xml.child(auto, "style:paragraph-properties")
  if auto and (xml.attr(auto, "style:master-page-name") or props and xml.attr(props, "fo:break-before")) then
    xml.set_attr(p, "text:style-name", xml.attr(auto, "style:parent-style-name"))
  end
  return p
end

function M.first_para(ctx)
  return xml.append(xml.find(ctx.root, "office:text"), xml.elem("text:p"))
end

automatic = function(ctx) -- <office:automatic-styles> (created if needed)
  if not ctx.auto then
    local body = xml.find(ctx.root, "office:body")
    ctx.auto = xml.append(body.parent, xml.elem("office:automatic-styles"), xml.index(body))
  end
  return ctx.auto
end

-- ── Lists and numbering ─────────────────────────────────────────────────────

-- A list style for new lists ("bullet" or "number"), made once.
local function list_style(ctx, kind)
  local name = kind == "bullet" and "NeoBullets" or "NeoNumbers"
  if not ctx.list_styles[name] then
    local style = xml.elem("text:list-style", { { "style:name", name } })
    for level = 1, 10 do
      local attrs = { { "text:level", tostring(level) } }
      if kind == "bullet" then
        table.insert(attrs, { "text:bullet-char", ({ "•", "◦", "▪" })[(level - 1) % 3 + 1] })
      else
        vim.list_extend(attrs, { { "style:num-suffix", "." }, { "style:num-format", "1" } })
      end
      local indent = ("%.2fin"):format(0.25 * (level + 1))
      xml.append(style, xml.elem(kind == "bullet" and "text:list-level-style-bullet" or "text:list-level-style-number", attrs, {
        xml.elem("style:list-level-properties", { { "text:list-level-position-and-space-mode", "label-alignment" } }, {
          xml.elem("style:list-level-label-alignment", {
            { "text:label-followed-by", "listtab" }, { "text:list-tab-stop-position", indent },
            { "fo:text-indent", "-0.25in" }, { "fo:margin-left", indent },
          }),
        }),
      }))
    end
    xml.append(automatic(ctx), style)
    ctx.list_styles[name] = style
  end
  return name
end

-- Make paragraph p an item of a list like `want` ({ kind, level }), or of no list: lift it out of
-- its list (splitting the list around it), then put it in the list just before it, if that is of
-- the same kind, or in a new one. List items that don't change are left as they are.
-- (a list after a blank line is a list of its own; without one, it goes on with the list above it)
local function join_or_split(ctx, p, want)
  local list = list_of(ctx, p)
  if not list or list.level ~= 1 or not want then
    return
  end
  local item, top = p.parent, list.id
  local items = element_children(top)
  local first = items[1] == item and element_children(item)[1] == p
  if want.continues and not first and xml.attr(item, "text:start-value") then -- (numbered on: no more starting again)
    xml.set_attr(item, "text:start-value", nil)
    return
  end
  local siblings = element_children(top.parent)
  local at = vim.iter(ipairs(siblings)):find(function(_, c) return c == top end)
  local before = at and siblings[at - 1]
  if want.continues and first and before and before.tag == "text:list" and list_kind(ctx, xml.attr(before, "text:style-name"), 1) == want.kind then
    for _, c in ipairs(vim.list_slice(top.children)) do -- (its items join the list above it)
      xml.remove(c)
      xml.append(before, c)
    end
    xml.remove(top)
  elseif not want.continues and not first and element_children(item)[1] == p then
    local attrs = vim.tbl_filter(function(kv) return kv[1] ~= "xml:id" and kv[1] ~= "text:continue-numbering" and kv[1] ~= "text:continue-list" end,
      vim.deepcopy(top.attrs))
    local rest = xml.append(top.parent, xml.elem("text:list", attrs), xml.index(top) + 1)
    for _, c in ipairs(vim.list_slice(top.children, xml.index(item))) do -- (it and what follows: a list of their own)
      xml.remove(c)
      xml.append(rest, c)
    end
  end
end

local function set_list(ctx, p, want)
  local have = list_of(ctx, p)
  local own_item = not have or element_children(p.parent)[1] == p -- (a new paragraph may share one)
  if own_item and (have and have.kind) == (want and want.kind) and (have and have.level) == (want and want.level) then
    join_or_split(ctx, p, want)
    if want and want.kind == "number" and want.pattern then
      set_pattern(ctx, p, want.pattern)
    end
    return
  end
  while p.parent and LISTS[p.parent.tag] do
    local parent = p.parent
    local after = vim.list_slice(parent.children, xml.index(p) + 1)
    if vim.iter(after):any(function(c) return c.tag end) then -- what follows it goes in a copy
      local attrs = vim.tbl_filter(function(kv) return kv[1] ~= "xml:id" end, vim.deepcopy(parent.attrs))
      local rest = xml.append(parent.parent, xml.elem(parent.tag, attrs), xml.index(parent) + 1)
      for _, c in ipairs(after) do
        xml.remove(c)
        xml.append(rest, c)
      end
    end
    xml.remove(p)
    xml.append(parent.parent, p, xml.index(parent) + 1)
    if #element_children(parent) == 0 then
      xml.remove(parent)
    end
  end
  if want then
    local siblings = element_children(p.parent)
    local index = vim.iter(ipairs(siblings)):find(function(_, c) return c == p end)
    local prev = siblings[index - 1]
    -- (the list above it, if it goes on with it: of its kind, or any for a sub-item, which goes in its last item)
    local container = want.continues and prev and prev.tag == "text:list"
      and (want.level > 1 or list_kind(ctx, xml.attr(prev, "text:style-name"), 1) == want.kind) and prev
    container = container or xml.append(p.parent, xml.elem("text:list", { { "text:style-name", list_style(ctx, want.kind) } }), xml.index(p))
    local top = container
    local function style_of(list) -- (a list's style: its own, else the list's it is in)
      local node = list
      while node and LISTS[node.tag] do
        local name = node.tag == "text:list" and xml.attr(node, "text:style-name")
        if name then
          return name
        end
        node = node.parent
      end
    end
    for level = 2, want.level do -- down to its level, in the last item
      local items = element_children(container)
      local item = items[#items] and items[#items].tag == "text:list-item" and items[#items] or xml.append(container, xml.elem("text:list-item"))
      local inside = element_children(item)
      local nested = inside[#inside] and inside[#inside].tag == "text:list" and inside[#inside]
      if nested and level == want.level and list_kind(ctx, style_of(nested), level) ~= want.kind then
        nested = nil -- (a sub-list of the other kind: this one is another, after it)
      end
      if not nested then
        nested = xml.append(item, xml.elem("text:list"))
        if list_kind(ctx, style_of(container), level) ~= want.kind then -- (of its own kind: bullets under a number...)
          xml.set_attr(nested, "text:style-name", list_style(ctx, want.kind))
        end
      end
      container = nested
    end
    xml.remove(p)
    xml.append(xml.append(container, xml.elem("text:list-item")), p)
    -- the rest of a list it was taken out of (when it was a new paragraph in it) joins again
    local after = element_children(top.parent)
    local next_list = after[(vim.iter(ipairs(after)):find(function(_, c) return c == top end) or #after) + 1]
    if next_list and next_list.tag == "text:list" and list_kind(ctx, xml.attr(next_list, "text:style-name"), 1) == want.kind then
      for _, c in ipairs(vim.list_slice(next_list.children)) do
        xml.remove(c)
        xml.append(top, c)
      end
      xml.remove(next_list)
    end
    if want.kind == "number" and want.pattern then
      set_pattern(ctx, p, want.pattern)
    end
  end
end

-- :Neoffice numbers: number the headings (the document's outline numbering), or stop (also
-- headings numbered by being in a list).
function M.number_headings(ctx, on)
  if not on then
    for _, h in ipairs(xml.collect(xml.find(ctx.root, "office:text"), { ["text:h"] = true }, { ["text:tracked-changes"] = true })) do
      if list_of(ctx, h) then
        set_list(ctx, h, nil)
        ctx.dirty[M.MAIN] = true
      end
    end
  end
  local outline = xml.find(ctx.styles_doc, "text:outline-style")
  if not outline then
    local styles = assert(xml.find(ctx.styles_doc, "office:styles"), "no office:styles in styles.xml")
    outline = xml.append(styles, xml.elem("text:outline-style", { { "style:name", "Outline" } }))
    for level = 1, 10 do
      xml.append(outline, xml.elem("text:outline-level-style", { { "text:level", tostring(level) }, { "style:num-format", "" } }))
    end
  end
  for _, level_style in ipairs(element_children(outline)) do
    local level = xml.attr(level_style, "text:level")
    xml.set_attr(level_style, "style:num-format", on and "1" or "")
    xml.set_attr(level_style, "text:display-levels", on and level ~= "1" and level or nil)
    local alignment = xml.find(level_style, "style:list-level-label-alignment")
    if on and alignment then -- (a space between the number and the heading)
      xml.set_attr(alignment, "text:label-followed-by", "space")
    end
  end
  -- heading styles with an empty list style opt out of the outline numbering: not any more
  for _, style in ipairs(xml.collect(ctx.styles_doc, { ["style:style"] = true })) do
    if on and xml.attr(style, "style:default-outline-level") and xml.attr(style, "style:list-style-name") == "" then
      xml.set_attr(style, "style:list-style-name", nil)
    end
  end
  ctx.numbered, ctx.dirty[STYLES] = on, true
end

-- The style to use for a piece: its original one, or a copy with bold/italic/underline changed.
-- The styles of the spans for a piece (nested, outermost first): its original ones, or with the
-- innermost a copy with bold/italic/... changed.
local function style_for(ctx, piece, base, inherit)
  local chain = type(piece.props) == "table" and piece.props or { piece.props }
  local outer, inner, inherited = vim.list_slice(chain, 1, #chain - 1), chain[#chain], {}
  for _, name in ipairs(outer) do
    inherited = flags(ctx, name, inherited)
  end
  local have = flags(ctx, inner, inherited)
  -- its font and size: only if what it shows changes (to the paragraph style's for none, said)
  local paras, fonts = require("neoffice.paras"), {}
  local own = M.run_font(ctx, chain)
  base, inherit = base or {}, inherit or base or {}
  for _, key in ipairs({ "font", "size" }) do
    local want = piece[key]
    if want ~= nil and paras.measure(key, want or base[key]) ~= paras.measure(key, own[key] or inherit[key] or base[key]) then
      fonts[key] = want or base[key] or false
    end
  end
  if vim.iter(FLAGS):all(function(k) return have[k] == piece[k] end) and not fonts.font and not fonts.size then
    return chain
  end
  local key = table.concat(vim.tbl_map(function(k) return tostring(piece[k]) end, FLAGS), "|") .. "|" .. (inner or "")
    .. "|" .. tostring(fonts.font) .. "|" .. tostring(fonts.size)
  if not ctx.made[key] then
    automatic(ctx)
    local base = inner and ctx.styles[inner]
    local style
    if base and base.parent == ctx.auto then -- automatic styles can't be inherited: copy it
      style = xml.copy(base)
    else
      style = xml.elem("style:style", { { "style:family", "text" } })
      xml.set_attr(style, "style:parent-style-name", inner)
    end
    local name, n = nil, 1
    repeat
      name, n = "Neo" .. n, n + 1
    until not ctx.styles[name]
    xml.set_attr(style, "style:name", name)
    local props = xml.child(style, "style:text-properties") or xml.append(style, xml.elem("style:text-properties"))
    for _, suffix in ipairs({ "", "-asian", "-complex" }) do
      xml.set_attr(props, (suffix == "" and "fo:" or "style:") .. "font-weight" .. suffix, piece.b and "bold" or "normal")
      xml.set_attr(props, (suffix == "" and "fo:" or "style:") .. "font-style" .. suffix, piece.i and "italic" or "normal")
    end
    xml.set_attr(props, "style:text-underline-style", piece.u and "solid" or "none")
    xml.set_attr(props, "style:text-underline-width", piece.u and "auto" or nil)
    xml.set_attr(props, "style:text-underline-color", piece.u and "font-color" or nil)
    xml.set_attr(props, "style:text-position", piece.sup and "super 58%" or piece.sub and "sub 58%" or "0% 100%")
    xml.set_attr(props, "style:text-line-through-style", piece.s and "solid" or "none")
    xml.set_attr(props, "style:text-line-through-type", piece.s and "single" or nil)
    xml.set_attr(props, "fo:background-color", piece.h and "#" .. (markup.COLORS[piece.h] or piece.h:gsub("^#", "")) or "transparent")
    if fonts.font ~= nil then -- (false: none known to say: the paragraph style's)
      xml.set_attr(props, "style:font-name", fonts.font and font_face(ctx, fonts.font) or nil)
      xml.set_attr(props, "fo:font-family", nil)
    end
    if fonts.size ~= nil then
      for _, attr in ipairs({ "fo:font-size", "style:font-size-asian", "style:font-size-complex" }) do
        xml.set_attr(props, attr, fonts.size or nil)
      end
    end
    xml.append(ctx.auto, style)
    ctx.styles[name], ctx.made[key] = style, name
  end
  return vim.list_extend(outer, { ctx.made[key] })
end

-- Append text, turning tabs into <text:tab/> and repeated or leading spaces into <text:s/>.
local function put_text(el, s, after_space)
  local pending = {}
  local function flush()
    if #pending > 0 then
      xml.append(el, { text = table.concat(pending) })
      pending = {}
    end
  end
  for ch in s:gmatch(".") do
    if ch == "\t" or (ch == " " and after_space) then
      flush()
      xml.append(el, xml.elem(ch == "\t" and "text:tab" or "text:s"))
    else
      pending[#pending + 1] = ch
    end
    after_space = ch == " "
  end
  flush()
  return after_space
end

local function annotation_el(c, parent_name)
  if c.node then -- as it was (only "resolved" can have changed)
    if (not c.resolved) ~= (not c.was_resolved) then
      xml.set_attr(c.node, "loext:resolved", c.resolved and "true" or "false")
      c.was_resolved = c.resolved
    end
    return c.node
  end
  local el = xml.elem("office:annotation", {
    parent_name and { "loext:parent-name", parent_name } or { "office:name", c.name },
    { "loext:resolved", c.resolved and "true" or "false" },
  }, {
    xml.elem("dc:creator", {}, { { text = c.author } }),
    xml.elem("dc:date", {}, { { text = c.date } }),
  })
  for _, line in ipairs(vim.split(c.text, "\n")) do
    xml.append(el, xml.elem("text:p", {}, { { text = line } }))
  end
  return el
end

local function tracked_changes(ctx)
  local found = xml.find(ctx.root, "text:tracked-changes")
  return found or xml.append(xml.find(ctx.root, "office:text"), xml.elem("text:tracked-changes"), 1)
end

-- The <text:changed-region> describing a change made in Neovim.
local function region(ctx, ch)
  if ch.region then
    return
  end
  ch.id = ch.id or ("ctneo%d%d"):format(os.time(), math.random(1e6))
  local info = xml.elem("office:change-info", {}, {
    xml.elem("dc:creator", {}, { { text = ch.author } }),
    xml.elem("dc:date", {}, { { text = ch.date } }),
  })
  local kind = xml.elem(ch.type == "delete" and "text:deletion" or "text:insertion", {}, { info })
  if ch.type == "delete" then
    for _, line in ipairs(vim.split(ch.text, "\n")) do
      xml.append(kind, xml.elem("text:p", {}, { { text = line } }))
    end
  end
  ch.region = xml.append(tracked_changes(ctx), xml.elem("text:changed-region", { { "text:id", ch.id } }, { kind }))
  ctx.changes[ch.id] = ch
end

-- A length like "6.5in" or "16.51cm": number and unit.
local function length(s)
  local n, unit = (s or ""):match("^([%d.]+)(%a*)$")
  return tonumber(n), unit
end

-- An image piece (see init.lua): a new image, another file for an image, or new alt text.
local function image_node(ctx, img)
  local original = img.original and img.original.image and img.original
  local frame
  local wanted = img.width and markup.cm(img.width, M.text_width(ctx))
  local function size(el, cm, ratio) -- (height from the width: the same shape)
    xml.set_attr(el, "svg:width", ("%.3fcm"):format(cm))
    xml.set_attr(el, "svg:height", ("%.3fcm"):format(cm * ratio))
  end
  if original and original.image.path == img.path then -- only the alt text or the width changed
    frame = xml.copy(original.node)
    local w, h = markup.cm(xml.attr(frame, "svg:width")), markup.cm(xml.attr(frame, "svg:height"))
    if wanted and img.width ~= original.image.width and w and h then
      size(frame, wanted, h / w)
    end
  else
    local ext, width, height = images.info(img.bytes)
    local file = img.path
    if not ctx.entries[file] then -- a file from disk goes into Pictures/
      local n = 1
      repeat
        file, n = ("Pictures/neoffice%d.%s"):format(n, ext), n + 1
      until not ctx.entries[file]
      ctx.entries[file], ctx.files[file] = true, img.bytes
      ctx.manifest = ctx.manifest or xml.parse(assert(zip.read(ctx.path, MANIFEST), "no manifest"))
      local entry = { { "manifest:full-path", file }, { "manifest:media-type", "image/" .. ext } }
      xml.append(xml.find(ctx.manifest, "manifest:manifest"), xml.elem("manifest:file-entry", entry))
      ctx.dirty[MANIFEST] = true
    end
    if original then -- same width as the image it replaces (unless another is given)
      frame = xml.copy(original.node)
      local w, unit = length(xml.attr(frame, "svg:width"))
      if wanted then
        size(frame, wanted, height / width)
      elseif w then
        xml.set_attr(frame, "svg:height", ("%.4f%s"):format(w * height / width, unit))
      end
    else
      local inches = math.min(width / 96, 6)
      frame = xml.elem("draw:frame", {
        { "draw:name", vim.fs.basename(file) }, { "text:anchor-type", "as-char" },
        { "svg:width", ("%.4fin"):format(inches) }, { "svg:height", ("%.4fin"):format(inches * height / width) },
      }, { xml.elem("draw:image", { { "xlink:type", "simple" }, { "xlink:show", "embed" }, { "xlink:actuate", "onLoad" } }) })
      if wanted then
        size(frame, wanted, height / width)
      end
    end
    local image = xml.child(frame, "draw:image")
    xml.set_attr(image, "xlink:href", file)
    xml.set_attr(image, "draw:mime-type", xml.attr(image, "draw:mime-type") and "image/" .. ext or nil)
  end
  local desc = xml.child(frame, "svg:desc") or xml.append(frame, xml.elem("svg:desc"))
  desc.children = {}
  xml.append(desc, { text = img.alt })
  local floats = require("neoffice.floats")
  local before = original and floats.crop(original.image) or nil
  if vim.inspect(before) ~= vim.inspect(floats.crop(img)) then -- its crop changed
    M.crop_frame(ctx, frame, img, before)
  end
  if not original or floats.place_key(original.image) ~= floats.place_key(img) then -- (moved, or new)
    M.place_frame(ctx, frame, img)
  end
  img.node = frame
  return frame
end

-- A paragraph's own format's attributes set as `want` says (only those whose value changed),
-- through an automatic style.
local ODF_ATTR = {
  ["indent-left"] = "fo:margin-left", ["indent-right"] = "fo:margin-right", first = "fo:text-indent",
  before = "fo:margin-top", after = "fo:margin-bottom", line = "fo:line-height",
}
local function set_para(ctx, p, want)
  local paras = require("neoffice.paras")
  local have = para_attrs(ctx, p)
  local auto = ctx.auto_paras[xml.attr(p, "text:style-name") or ""]
  local base = style_attrs(ctx, auto and xml.attr(auto, "style:parent-style-name") or xml.attr(p, "text:style-name"))
  want = vim.deepcopy(want) -- (a value its style gives already: none of its own)
  for _, key in ipairs(paras.LAYOUT) do
    if want[key] and want[key] ~= "none" and paras.measure(key, want[key]) == paras.measure(key, base[key] or (key == "line" and "single" or "0cm")) then
      want[key] = nil
    end
  end
  local changes = {}
  for _, key in ipairs(paras.LAYOUT) do
    if paras.measure(key, have[key]) ~= paras.measure(key, want[key]) then
      local value = want[key] ~= "none" and want[key] or nil
      if key == "line" and value then
        local multiple = ({ single = 1, double = 2 })[value] or tonumber(value)
        value = multiple and ("%g%%"):format(multiple * 100) or value
      end
      changes[#changes + 1] = { ODF_ATTR[key], value or false }
    end
  end
  if #changes == 0 then
    return
  end
  local name = xml.attr(p, "text:style-name")
  local parts = { "para-format", name or "" }
  for _, c in ipairs(changes) do
    parts[#parts + 1] = c[1] .. "=" .. tostring(c[2])
  end
  local key = table.concat(parts, "|")
  if not ctx.made[key] then
    local auto = ctx.auto_paras[name or ""]
    local style = auto and xml.copy(auto) or xml.elem("style:style", { { "style:family", "paragraph" } })
    if not auto and name then
      xml.set_attr(style, "style:parent-style-name", name)
    end
    local new, n = nil, 1
    repeat
      new, n = "NeoP" .. n, n + 1
    until ctx.parent_style[new] == nil
    xml.set_attr(style, "style:name", new)
    local props = xml.child(style, "style:paragraph-properties") or xml.append(style, xml.elem("style:paragraph-properties"), 1)
    for _, c in ipairs(changes) do
      xml.set_attr(props, c[1], c[2] or nil)
      if c[1] == "fo:line-height" then
        xml.set_attr(props, "style:line-height-at-least", nil)
      end
    end
    xml.append(automatic(ctx), style)
    ctx.made[key], ctx.auto_paras[new] = new, style
    ctx.parent_style[new] = xml.attr(style, "style:parent-style-name") or false
    if ctx.para_style_el then -- (known as a paragraph style: its font, see font_base)
      ctx.para_style_el[new] = style
    end
  end
  xml.set_attr(p, "text:style-name", ctx.made[key])
end

-- Give a paragraph this alignment and page break (when they changed), through an automatic style.
local ALIGN_VALUE = { center = "center", right = "end", justify = "justify", left = "start" }
local function set_layout(ctx, p, align, page) -- (align: nil for its style's)
  local _, have_page = layout(ctx, p)
  local have_align = own_align(ctx, p)
  if have_align == align and have_page == page then
    return
  end
  local name = xml.attr(p, "text:style-name")
  local key = table.concat({ "para", name or "", tostring(align), tostring(page) }, "|")
  if not ctx.made[key] then
    local auto = ctx.auto_paras[name or ""]
    local style = auto and xml.copy(auto) or xml.elem("style:style", { { "style:family", "paragraph" } })
    if not auto and name then
      xml.set_attr(style, "style:parent-style-name", name)
    end
    local new, n = nil, 1
    repeat
      new, n = "NeoP" .. n, n + 1
    until ctx.parent_style[new] == nil
    xml.set_attr(style, "style:name", new)
    local props = xml.child(style, "style:paragraph-properties") or xml.append(style, xml.elem("style:paragraph-properties"), 1)
    if have_align ~= align then -- (what it shows: its own, or its style's again)
      xml.set_attr(props, "fo:text-align", ALIGN_VALUE[align or M.style_align(ctx, M.style_of(ctx, p)) or "left"])
    end
    xml.set_attr(props, "fo:break-before", page and "page" or nil)
    xml.append(automatic(ctx), style)
    ctx.made[key], ctx.auto_paras[new] = new, style
    ctx.parent_style[new] = xml.attr(style, "style:parent-style-name") or false
    if ctx.para_style_el then -- (known as a paragraph style: its font, see font_base)
      ctx.para_style_el[new] = style
    end
  end
  xml.set_attr(p, "text:style-name", ctx.made[key])
end

-- ── Styles (the document's header: see init.lua) ───────────────────────────
-- A paragraph style's name (as LibreOffice shows it: "Text Body"), and what it gives (following
-- its parents, and the default style): { font, size, align, bold, italic, and paras.LAYOUT }.
function M.style_name(ctx, name)
  style_attrs(ctx, nil)
  local el = name and ctx.para_style_el[name]
  local shown = el and xml.attr(el, "style:display-name")
  shown = shown or (name or "Standard"):gsub("_(%x%x)_", function(hex) return string.char(tonumber(hex, 16)) end)
  return shown
end
function M.paragraph_styles(ctx) -- (the names of all its named paragraph styles)
  style_attrs(ctx, nil)
  local out = {}
  for name in pairs(ctx.para_style_el) do
    if name and not ctx.auto_paras[name] then
      out[#out + 1] = M.style_name(ctx, name)
    end
  end
  table.sort(out)
  return out
end
-- The named paragraph styles in use: those of the paragraphs in the text and in the page styles
-- (headers, footers), an automatic style's for its paragraphs.
function M.styles_in_use(ctx)
  local auto = {} -- (the automatic paragraph styles: of the text, and of the page styles)
  for _, root in ipairs({ ctx.root, ctx.styles_doc }) do
    local list = xml.find(root, "office:automatic-styles")
    for _, st in ipairs(list and list.children or {}) do
      if st.tag == "style:style" and xml.attr(st, "style:family") == "paragraph" then
        auto[root == ctx.root and "c|" or "s|"] = auto[root == ctx.root and "c|" or "s|"] or {}
        auto[root == ctx.root and "c|" or "s|"][xml.attr(st, "style:name") or ""] = st
      end
    end
  end
  local used = {}
  for _, root in ipairs({ ctx.root, ctx.styles_doc }) do
    local mine = auto[root == ctx.root and "c|" or "s|"] or {}
    for _, el in ipairs(xml.collect(root, { ["text:p"] = true, ["text:h"] = true })) do
      local name = xml.attr(el, "text:style-name")
      if name and mine[name] then
        name = xml.attr(mine[name], "style:parent-style-name")
      end
      if name then
        used[name] = true
      end
    end
  end
  return vim.tbl_keys(used)
end
function M.style_id(ctx, shown) -- (by its name, whatever its case: a named style, in styles.xml)
  style_attrs(ctx, nil)
  for name, el in pairs(ctx.para_style_el) do
    if name and not ctx.auto_paras[name] and (M.style_name(ctx, name):lower() == shown:lower() or name:lower() == shown:lower()) then
      return name, el
    end
  end
end
function M.style_view(ctx, name)
  local paras = require("neoffice.paras")
  local out = vim.deepcopy(style_attrs(ctx, name))
  for _, key in ipairs(paras.LAYOUT) do -- (the defaults: unsaid)
    if out[key] and paras.measure(key, out[key]) == paras.measure(key, key == "line" and "single" or "0cm") then
      out[key] = nil
    end
  end
  local font = M.font_base_of(ctx, name)
  out.font, out.size = font.font, font.size
  out.align = M.style_align(ctx, name)
  local weight, slant -- (the first of its chain that says, else the default style)
  local cur = name
  for _ = 1, 10 do
    local el = cur and ctx.para_style_el[cur]
    if not el then
      break
    end
    local props = xml.child(el, "style:text-properties")
    weight = weight or (props and xml.attr(props, "fo:font-weight"))
    slant = slant or (props and xml.attr(props, "fo:font-style"))
    cur = xml.attr(el, "style:parent-style-name")
  end
  local default = ctx.para_style_el[false]
  local dprops = default and xml.child(default, "style:text-properties")
  weight = weight or (dprops and xml.attr(dprops, "fo:font-weight"))
  slant = slant or (dprops and xml.attr(dprops, "fo:font-style"))
  out.bold = weight and (weight == "bold" or (tonumber(weight) or 0) >= 600) or nil
  out.italic = slant and slant ~= "normal" or nil
  return out
end
-- A style's properties changed: { key = value, or false for the default }. Only those given.
function M.set_style(ctx, name, changes)
  style_attrs(ctx, nil)
  local style = assert(ctx.para_style_el[name], "no style " .. tostring(name))
  local function props_of(tag, first)
    return xml.child(style, tag) or xml.append(style, xml.elem(tag), first and 1 or nil)
  end
  for key, value in pairs(changes) do
    if ODF_ATTR[key] then
      local v = value or (key == "line" and "single" or "0cm")
      if key == "line" then
        local multiple = ({ single = 1, double = 2 })[v] or tonumber(v)
        v = multiple and ("%g%%"):format(multiple * 100) or v
      end
      xml.set_attr(props_of("style:paragraph-properties", true), ODF_ATTR[key], v)
    elseif key == "align" then
      xml.set_attr(props_of("style:paragraph-properties", true), "fo:text-align", ALIGN_VALUE[value or "left"])
    else
      local props = props_of("style:text-properties")
      if key == "font" then
        xml.set_attr(props, "style:font-name", value and font_face(ctx, value, true) or nil)
        xml.set_attr(props, "fo:font-family", nil)
      elseif key == "size" then
        for _, attr in ipairs({ "fo:font-size", "style:font-size-asian", "style:font-size-complex" }) do
          xml.set_attr(props, attr, value or nil)
        end
      elseif key == "bold" or key == "italic" then
        for _, suffix in ipairs({ "", "-asian", "-complex" }) do
          local attr = (suffix == "" and "fo:" or "style:") .. (key == "bold" and "font-weight" or "font-style") .. suffix
          xml.set_attr(props, attr, key == "bold" and (value and "bold" or "normal") or (value and "italic" or "normal"))
        end
      end
    end
  end
  ctx.dirty[STYLES] = true -- (what styles give: read again)
  ctx.style_attrs, ctx.font_bases, ctx.font_inherits, ctx.style_aligns = {}, nil, nil, nil
end

-- The style of a caption of this kind ("table", "figure"): LibreOffice's "Table" or "Figure"
-- (or "Illustration") if the document has it, else its "Caption", else a "Caption" style added.
caption_style = function(ctx, kind)
  for _, name in ipairs(kind == "table" and { "Table", "Caption" } or { "Figure", "Illustration", "Caption" }) do
    if ctx.parent_style[name] ~= nil then
      return name
    end
  end
  local styles = xml.find(ctx.styles_doc, "office:styles")
  if styles then
    xml.append(styles, xml.elem("style:style", { { "style:name", "Caption" }, { "style:family", "paragraph" }, { "style:parent-style-name", "Standard" }, { "style:class", "extra" } }, {
      xml.elem("style:paragraph-properties", { { "fo:margin-top", "0.212cm" }, { "fo:margin-bottom", "0.212cm" } }),
    }))
    ctx.parent_style.Caption, ctx.dirty[STYLES] = "Standard", true
  end
  return "Caption"
end

-- A caption's number: a sequence counting the captions of its word, showing `n` (declared in the
-- document's sequences if it isn't).
function M.caption_number(ctx, label, n)
  local text = xml.find(ctx.root, "office:text")
  local decls = text and xml.child(text, "text:sequence-decls")
  if text and not decls then
    decls = xml.append(text, xml.elem("text:sequence-decls"), 1)
  end
  local declared = decls and vim.iter(xml.collect(decls, { ["text:sequence-decl"] = true })):any(function(d)
    return xml.attr(d, "text:name") == label
  end)
  if decls and not declared then
    xml.append(decls, xml.elem("text:sequence-decl", { { "text:display-outline-level", "0" }, { "text:name", label } }))
  end
  ctx.sequences_made = (ctx.sequences_made or 0) + 1
  return xml.elem("text:sequence", {
    { "text:ref-name", ("refNeo%s%d_%d"):format(label, os.time(), ctx.sequences_made) }, { "text:name", label },
    { "text:formula", "ooow:" .. label .. "+1" }, { "style:num-format", "1" },
  }, { { text = tostring(n or 1) } })
end

-- Every caption's number as it is in the document's order (each word counts its own). Whether
-- any changed.
function M.number_captions(ctx)
  local count, changed = {}, false
  for _, seq in ipairs(xml.collect(xml.find(ctx.root, "office:text") or {}, { ["text:sequence"] = true })) do
    local name = xml.attr(seq, "text:name") or ""
    count[name] = (count[name] or 0) + 1
    if xml.text(seq) ~= tostring(count[name]) then
      seq.children, changed = { { text = tostring(count[name]) } }, true
    end
  end
  if changed then
    ctx.dirty[M.MAIN] = true
  end
  return changed
end

-- A footnote piece: the same <text:note> with new text, or a new one.
local function footnote_node(ctx, piece)
  local original = piece.original and piece.original.footnote and piece.original.node
  local note = original or xml.elem("text:note", { { "text:note-class", "footnote" } }, {
    xml.elem("text:note-citation", {}, { { text = tostring(#xml.collect(ctx.root, { ["text:note"] = true }) + 1) } }),
    xml.elem("text:note-body"),
  })
  if not original then
    ctx.note_n = (ctx.note_n or 0) + 1
    xml.set_attr(note, "text:id", ("ftnneo%d_%d"):format(os.time(), ctx.note_n))
  end
  local body = xml.child(note, "text:note-body")
  local first = vim.iter(body.children):find(function(c) return PARA[c.tag] end)
  local para = xml.elem("text:p", first and vim.deepcopy(first.attrs) or { { "text:style-name", "Footnote" } })
  put_text(para, piece.text, true)
  body.children = {}
  xml.append(body, para)
  piece.node = note
  return note
end

-- The width of the text in cm: the page's, minus its margins (of the "Standard" page style, or the
-- first). Letter with 1 inch margins if the document doesn't say.
function M.text_width(ctx)
  local masters = xml.collect(ctx.styles_doc, { ["style:master-page"] = true })
  local master = vim.iter(masters):find(function(m) return xml.attr(m, "style:name") == "Standard" end) or masters[1]
  local name = master and xml.attr(master, "style:page-layout-name")
  local layout = name and vim.iter(xml.collect(ctx.styles_doc, { ["style:page-layout"] = true })):find(function(l)
    return xml.attr(l, "style:name") == name
  end)
  local props = layout and xml.child(layout, "style:page-layout-properties") or { attrs = {} }
  local w, left, right = markup.cm(xml.attr(props, "fo:page-width")), markup.cm(xml.attr(props, "fo:margin-left")), markup.cm(xml.attr(props, "fo:margin-right"))
  return (w and left and right) and w - left - right or 16.51
end

-- ── Tables (see tables.lua) ─────────────────────────────────────────────────

function M.new_table(ctx) -- one cell (rows and cells are added by copying; borders: see tables.lua)
  if not ctx.styles_made_cell then
    xml.append(automatic(ctx), xml.elem("style:style", { { "style:name", "NeoCell" }, { "style:family", "table-cell" } }, {
      xml.elem("style:table-cell-properties", { { "fo:padding", "0.04in" } }),
    }))
    if ctx.named then -- (see style_named)
      ctx.named["table-cell|NeoCell"] = { el = ctx.auto.children[#ctx.auto.children], automatic = true }
    end
    ctx.styles_made_cell = true
  end
  ctx.tables_made = (ctx.tables_made or 0) + 1
  return xml.elem("table:table", { { "table:name", ("NeoTable%d_%d"):format(os.time(), ctx.tables_made) } }, {
    xml.elem("table:table-column"),
    xml.elem("table:table-row", {}, {
      xml.elem("table:table-cell", { { "table:style-name", "NeoCell" }, { "office:value-type", "string" } }, {
        -- (in LibreOffice's style for them, if the document has it)
        xml.elem("text:p", ctx.parent_style.Table_20_Contents ~= nil and { { "text:style-name", "Table_20_Contents" } } or {}),
      }),
    }),
  })
end

-- The style of a family ("table", "table-cell") with this name: { el, automatic = bool }, or nil.
local function style_named(ctx, family, name)
  if not ctx.named then
    ctx.named = {}
    for _, root in ipairs({ ctx.styles_doc, ctx.root }) do
      for _, style in ipairs(xml.collect(root, { ["style:style"] = true }, { ["office:body"] = true })) do
        ctx.named[xml.attr(style, "style:family") .. "|" .. (xml.attr(style, "style:name") or "")] = { el = style, automatic = root == ctx.root }
      end
    end
  end
  return name and ctx.named[family .. "|" .. name]
end
-- A new automatic style for an element, from its own: a copy if that one is automatic, else based
-- on it. Returns its properties element (`props`), and sets it as the element's style.
local function own_style(ctx, el, family, props)
  local base = style_named(ctx, family, xml.attr(el, "table:style-name"))
  local style
  if base and base.automatic then
    style = xml.copy(base.el)
  else
    style = xml.elem("style:style", { { "style:family", family } })
    xml.set_attr(style, "style:parent-style-name", base and xml.attr(base.el, "style:name") or nil)
  end
  local name, n = nil, 1
  repeat
    name, n = (({ table = "NeoTableStyle", ["table-column"] = "NeoColumn", ["table-row"] = "NeoRow" })[family] or "NeoCell") .. n, n + 1
  until not style_named(ctx, family, name)
  xml.set_attr(style, "style:name", name)
  xml.append(automatic(ctx), style)
  ctx.named[family .. "|" .. name] = { el = style, automatic = true }
  xml.set_attr(el, "table:style-name", name)
  return xml.child(style, props) or xml.append(style, xml.elem(props))
end

-- ── Floating images and frames (see floats.lua) ─────────────────────────────

-- ── Sections: page settings (see pages.lua) ─────────────────────────────────
-- A section starts at a paragraph (or table) whose style names a master page ("page style"); the
-- master page's page layout has its size, orientation and margins.
local function masters(ctx)
  local out = {}
  for _, m in ipairs(xml.collect(ctx.styles_doc, { ["style:master-page"] = true })) do
    out[xml.attr(m, "style:name") or ""] = m
  end
  return out
end
local function page_layout(ctx, master)
  local name = master and xml.attr(master, "style:page-layout-name")
  return name and vim.iter(xml.collect(ctx.styles_doc, { ["style:page-layout"] = true })):find(function(l)
    return xml.attr(l, "style:name") == name
  end)
end
local function layout_props(layout)
  local props = layout and xml.child(layout, "style:page-layout-properties") or { attrs = {} }
  local out = { width = markup.cm(xml.attr(props, "fo:page-width") or "") or 21, height = markup.cm(xml.attr(props, "fo:page-height") or "") or 29.7, start = "page" }
  for _, side in ipairs({ "top", "bottom", "left", "right" }) do
    out[side] = markup.cm(xml.attr(props, "fo:margin-" .. side) or "") or 2
  end
  return out
end
local function starts_at(node) -- (what a section starting at a paragraph starts at: its table, in one)
  local el = node
  while el.parent and el.tag ~= "table:table" do
    el = el.parent
  end
  return el.tag == "table:table" and el or node
end
local function master_of(ctx, el) -- (the master page an element's style names: nil if none)
  local name
  if el.tag == "table:table" then
    local st = style_named(ctx, "table", xml.attr(el, "table:style-name"))
    name = st and xml.attr(st.el, "style:master-page-name")
  else
    local style = xml.attr(el, "text:style-name")
    local auto = ctx.auto_paras[style or ""]
    name = auto and xml.attr(auto, "style:master-page-name")
    if name == nil then -- (its named style's, then)
      local named = style_named(ctx, "paragraph", auto and xml.attr(auto, "style:parent-style-name") or style)
      name = named and not named.automatic and xml.attr(named.el, "style:master-page-name") or nil
    end
  end
  return name ~= "" and name or nil
end
local function default_master(ctx)
  local all = masters(ctx)
  if all.Standard then
    return "Standard"
  end
  local first = xml.find(ctx.styles_doc, "style:master-page")
  return first and xml.attr(first, "style:name")
end

local function in_frame(node) -- (a box's paragraph: shown before the one it is anchored in)
  local el = node.parent
  while el do
    if el.tag and el.tag:match("^draw:") then -- (a frame, a text box, a shape with text)
      return true
    end
    el = el.parent
  end
  return false
end

function M.sections(ctx, paras)
  local all = masters(ctx)
  local out, prev, boxes = {}, nil, nil -- (boxes: where the boxes shown before a paragraph start)
  ctx.section_starts = {}
  for k, para in ipairs(paras) do
    local node = para.node or para
    if in_frame(node) then
      boxes = boxes or k
    else
      local el = starts_at(node)
      if el ~= prev then
        local m = master_of(ctx, el)
        if #out == 0 then -- (the first: its page style, named or the default)
          m = m and all[m] and m or default_master(ctx)
          out[1] = { start = 1, el = m, props = layout_props(page_layout(ctx, all[m or ""])) }
        elseif m and all[m] then
          out[#out + 1] = { start = boxes or k, el = m, props = layout_props(page_layout(ctx, all[m])) }
          ctx.section_starts[el] = true
        end
      end
      prev, boxes = el, nil
    end
  end
  if #out == 0 then
    local m = default_master(ctx)
    out[1] = { start = 1, el = m, props = layout_props(page_layout(ctx, all[m or ""])) }
  end
  return out
end

-- An element's style, naming this master page (nil: none; "" over a named style that names one).
local function set_master(ctx, el, name)
  if master_of(ctx, el) == name then
    return false
  end
  if el.tag == "table:table" then
    local props = own_style(ctx, el, "table", "style:table-properties")
    local base = style_named(ctx, "table", xml.attr(props.parent, "style:parent-style-name"))
    xml.set_attr(props.parent, "style:master-page-name", name or (base and xml.attr(base.el, "style:master-page-name") and "") or nil)
    return true
  end
  local style = xml.attr(el, "text:style-name")
  local key = "master|" .. (style or "") .. "|" .. (name or "")
  if not ctx.made[key] then
    local auto = ctx.auto_paras[style or ""]
    local new_style = auto and xml.copy(auto) or xml.elem("style:style", { { "style:family", "paragraph" } })
    if not auto and style then
      xml.set_attr(new_style, "style:parent-style-name", style)
    end
    local new, n = nil, 1
    repeat
      new, n = "NeoP" .. n, n + 1
    until ctx.parent_style[new] == nil
    xml.set_attr(new_style, "style:name", new)
    xml.set_attr(new_style, "style:master-page-name", nil)
    local auto_new = ctx.auto_paras
    auto_new[new] = new_style
    local inherited = master_of(ctx, { tag = "text:p", attrs = { { "text:style-name", new } } })
    xml.set_attr(new_style, "style:master-page-name", name or (inherited and "") or nil)
    xml.append(automatic(ctx), new_style)
    ctx.made[key] = new
    ctx.parent_style[new] = xml.attr(new_style, "style:parent-style-name") or false
    if ctx.para_style_el then
      ctx.para_style_el[new] = new_style
    end
  end
  xml.set_attr(el, "text:style-name", ctx.made[key])
  return true
end

-- Settings changed (see pages.changes) on a page layout.
local function set_layout_props(layout, ch)
  local props = xml.child(layout, "style:page-layout-properties") or xml.append(layout, xml.elem("style:page-layout-properties"), 1)
  local pages = require("neoffice.pages")
  if ch.size then
    xml.set_attr(props, "fo:page-width", pages.length(ch.width))
    xml.set_attr(props, "fo:page-height", pages.length(ch.height))
    xml.set_attr(props, "style:print-orientation", ch.landscape and "landscape" or "portrait")
  end
  for side, cm in pairs(ch.margins) do
    xml.set_attr(props, "fo:margin-" .. side, pages.length(cm))
  end
end

-- The sections as the buffer has them: list = { { el = the master page it had (nil: as the one
-- before), want = its attributes, first = its first paragraph } }. A master page whose sections all
-- want other settings gets them; one that only some do, a copy for those. Returns whether
-- anything changed.
function M.set_sections(ctx, list)
  local pages = require("neoffice.pages")
  local all = masters(ctx)
  local changed = false
  local base, groups = {}, {}
  for j, s in ipairs(list) do
    assert(s.want.start == "page", "a section of an .odt starts on a new page: start = page (continuous is Word's)")
    base[j] = s.el and all[s.el] and s.el or base[j - 1] or default_master(ctx)
    groups[base[j]] = groups[base[j]] or {}
    table.insert(groups[base[j]], j)
  end
  local used_layouts = {}
  for _, m in pairs(all) do
    local name = xml.attr(m, "style:page-layout-name") or ""
    used_layouts[name] = (used_layouts[name] or 0) + 1
  end
  local styles_auto = xml.find(ctx.styles_doc, "office:automatic-styles")
  local function new_layout(layout) -- (a copy of a page layout, of its own)
    local copy = xml.copy(layout or xml.elem("style:page-layout"))
    local name, n = nil, 1
    local names = {}
    for _, l in ipairs(xml.collect(ctx.styles_doc, { ["style:page-layout"] = true })) do
      names[xml.attr(l, "style:name") or ""] = true
    end
    repeat
      name, n = "NeoPL" .. n, n + 1
    until not names[name]
    xml.set_attr(copy, "style:name", name)
    xml.append(styles_auto, copy)
    return copy
  end
  local master, copies = {}, {}
  for m, js in pairs(groups) do
    local layout = page_layout(ctx, all[m])
    local have = pages.view(layout_props(layout))
    local wants = {}
    for _, j in ipairs(js) do
      wants[pages.render(list[j].want)] = list[j].want
    end
    local keys = vim.tbl_keys(wants)
    if #keys == 1 and keys[1] ~= pages.render(have) then -- (all of them: the master page itself)
      if not layout or (used_layouts[xml.attr(layout, "style:name") or ""] or 0) > 1 then
        layout = new_layout(layout)
        xml.set_attr(all[m], "style:page-layout-name", xml.attr(layout, "style:name"))
      end
      set_layout_props(layout, pages.changes(have, wants[keys[1]]))
      changed = true
    end
    for _, j in ipairs(js) do
      local key = pages.render(list[j].want)
      if #keys == 1 or key == pages.render(have) then
        master[j] = m
      else -- (only some: a copy of it for them)
        if not copies[m .. "|" .. key] then
          local copy = xml.copy(all[m])
          local name, n = nil, 1
          repeat
            name, n = "NeoPage" .. n, n + 1
          until not all[name]
          xml.set_attr(copy, "style:name", name)
          local want = list[j].want
          xml.set_attr(copy, "style:display-name", (xml.attr(all[m], "style:display-name") or m) .. " (" .. want.size .. ", " .. want.orientation .. ")")
          local own = new_layout(layout)
          set_layout_props(own, pages.changes(have, want))
          xml.set_attr(copy, "style:page-layout-name", xml.attr(own, "style:name"))
          xml.append(all[m].parent, copy)
          all[name], copies[m .. "|" .. key] = copy, name
          changed = true
        end
        master[j] = copies[m .. "|" .. key]
      end
    end
  end
  local starts = {}
  for j, s in ipairs(list) do
    local el = starts_at(s.first)
    starts[el] = true
    local now = master_of(ctx, el)
    if j > 1 or now or master[j] ~= default_master(ctx) then
      changed = set_master(ctx, el, master[j]) or changed
    end
  end
  for el in pairs(ctx.section_starts or {}) do -- (sections joined to the one before: no longer starting one)
    if not starts[el] and el.parent and master_of(ctx, el) then
      changed = set_master(ctx, el, nil) or changed
    end
  end
  if changed then
    ctx.dirty[M.MAIN], ctx.dirty[STYLES] = true, true
  end
  return changed
end

-- The page, for placing images: { text = the text's width, left, top = its margins } (cm).
function M.page(ctx)
  local master = xml.find(ctx.styles_doc, "style:master-page")
  local name = master and xml.attr(master, "style:page-layout-name")
  local layout = vim.iter(xml.collect(ctx.styles_doc, { ["style:page-layout"] = true })):find(function(l)
    return xml.attr(l, "style:name") == name
  end)
  local props = layout and xml.child(layout, "style:page-layout-properties") or { attrs = {} }
  return { text = M.text_width(ctx), left = markup.cm(xml.attr(props, "fo:margin-left")) or 2, top = markup.cm(xml.attr(props, "fo:margin-top")) or 2 }
end

-- A graphic style's property, following the styles it is based on.
local function graphic_prop(ctx, name, attr)
  for _ = 1, 10 do
    local style = style_named(ctx, "graphic", name)
    if not style then
      return nil
    end
    local value = xml.attr(xml.child(style.el, "style:graphic-properties") or { attrs = {} }, attr)
    if value then
      return value
    end
    name = xml.attr(style.el, "style:parent-style-name")
  end
end

-- Where a frame is: a geometry (see floats.lua), or nil if it is in the line (as a character).
function M.frame_geometry(ctx, frame)
  local anchor = xml.attr(frame, "text:anchor-type") or "paragraph"
  if anchor == "as-char" then
    return nil
  end
  local style = xml.attr(frame, "draw:style-name")
  local function prop(attr) return graphic_prop(ctx, style, attr) end
  local wrap = prop("style:wrap") or "none"
  if wrap == "run-through" then
    wrap = prop("style:run-through") == "background" and "behind" or "front"
  elseif wrap ~= "none" then
    wrap = prop("style:wrap-contour") == "true" and "tight" or "square"
  end
  local hpos, hrel = prop("style:horizontal-pos") or "center", prop("style:horizontal-rel") or "paragraph"
  local vpos, vrel = prop("style:vertical-pos") or "top", prop("style:vertical-rel") or "paragraph"
  local h = { rel = (anchor == "page" or hrel == "page" or hrel:match("^page%-")) and hrel ~= "page-content" and "page" or "column" }
  h.align = ({ left = "left", right = "right", center = "center", inside = "left", outside = "right" })[hpos]
  h.offset = not h.align and (markup.cm(xml.attr(frame, "svg:x")) or 0) or nil
  local v = { rel = (anchor == "page" or vrel == "page") and "page" or vrel == "page-content" and "margin" or "paragraph" }
  v.align = (vpos == "top" or vpos == "bottom") and v.rel ~= "paragraph" and vpos or nil
  v.offset = not v.align and (vpos == "from-top" and markup.cm(xml.attr(frame, "svg:y")) or 0) or nil
  local gap = markup.cm(prop(wrap == "none" and "fo:margin-top" or "fo:margin-left") or "") or 0
  return { h = h, v = v, wrap = wrap, gap = gap, width = markup.cm(xml.attr(frame, "svg:width")) or 0 }
end

-- An image's size before it is cropped (cm: from its pixels and resolution, as LibreOffice takes
-- it), which its crop (fo:clip) is measured on.
local function natural(data)
  local images = require("neoffice.images")
  local _, w, h = images.info(data or "")
  local dpi = w and images.dpi(data)
  return w and w / dpi * 2.54, h and h / dpi * 2.54
end

-- A frame's crop (its style's fo:clip: "rect(top, right, bottom, left)"), as attributes:
-- { ["crop-top"] = "39.3%", ... } (only the sides cut).
function M.frame_crop(ctx, frame, path)
  local clip = graphic_prop(ctx, xml.attr(frame, "draw:style-name"), "fo:clip") or ""
  local sides = {}
  for value in (clip:match("^rect%((.*)%)$") or ""):gmatch("[^,%s]+") do
    sides[#sides + 1] = markup.cm(value) or 0
  end
  if #sides ~= 4 or (sides[1] + sides[2] + sides[3] + sides[4]) <= 0 then
    return {}
  end
  local w, h = natural(zip.read(ctx.path, path)) -- (read only for a cropped image)
  if not w then
    return {}
  end
  local floats, out = require("neoffice.floats"), {}
  for k, side in ipairs({ sides[1] / h, sides[2] / w, sides[3] / h, sides[4] / w }) do
    out[floats.CROP[k]] = floats.percent(side)
  end
  return out
end

-- A frame cropped as its attributes say (keeping its width: its height, that of what is shown).
function M.crop_frame(ctx, frame, img, before)
  local function length_cm(n) return ("%.5fcm"):format(n) end -- (precise: a crop is a share of it)
  local floats = require("neoffice.floats")
  local after = floats.crop(img)
  local image = xml.child(frame, "draw:image")
  local path = image and xml.attr(image, "xlink:href")
  local w, h = natural(img.bytes or (path and zip.read(ctx.path, path)))
  if not w then
    return
  end
  local a = after or { 0, 0, 0, 0 }
  M.style_frame(ctx, frame, {
    ["fo:clip"] = after and ("rect(%s, %s, %s, %s)"):format(length_cm(a[1] * h), length_cm(a[2] * w), length_cm(a[3] * h), length_cm(a[4] * w)) or false,
  })
  local b = before or { 0, 0, 0, 0 }
  local factor = (1 - a[1] - a[3]) / math.max(1 - b[1] - b[3], 0.01) * (1 - b[2] - b[4]) / math.max(1 - a[2] - a[4], 0.01)
  local height = markup.cm(xml.attr(frame, "svg:height"))
  if height then
    xml.set_attr(frame, "svg:height", length_cm(height * factor))
  end
end

-- A frame placed as its attributes say (see floats.lua): in the line, or anchored, floating.
local function length_cm(n)
  return ("%.3fcm"):format(n)
end
function M.place_frame(ctx, frame, attrs, extra)
  local geo = require("neoffice.floats").geometry(attrs, M.page(ctx), markup.cm(xml.attr(frame, "svg:width")) or 0)
  if not geo then
    xml.set_attr(frame, "text:anchor-type", "as-char")
    xml.set_attr(frame, "svg:x", nil)
    xml.set_attr(frame, "svg:y", nil)
    return
  end
  xml.set_attr(frame, "text:anchor-type", "paragraph")
  local wrap = ({ square = "parallel", tight = "parallel", none = "none", front = "run-through", behind = "run-through" })[geo.wrap]
  local gap, side = length_cm(geo.gap), geo.wrap == "none" and "0cm" or length_cm(geo.gap)
  local props = {
    ["style:wrap"] = wrap, ["style:wrap-contour"] = geo.wrap == "tight" and "true" or "false",
    ["style:wrap-contour-mode"] = geo.wrap == "tight" and "outside" or false,
    ["style:run-through"] = geo.wrap == "behind" and "background" or "foreground",
    ["fo:margin-top"] = gap, ["fo:margin-bottom"] = gap, ["fo:margin-left"] = side, ["fo:margin-right"] = side,
    ["style:horizontal-pos"] = geo.h.align or "from-left", ["style:horizontal-rel"] = geo.h.rel == "page" and "page" or "paragraph",
    ["style:vertical-pos"] = geo.v.align or "from-top",
    ["style:vertical-rel"] = geo.v.rel == "page" and "page" or geo.v.rel == "margin" and "page-content" or "paragraph",
  }
  xml.set_attr(frame, "svg:x", not geo.h.align and length_cm(geo.h.offset or 0) or nil)
  xml.set_attr(frame, "svg:y", not geo.v.align and length_cm(geo.v.offset or 0) or nil)
  M.style_frame(ctx, frame, vim.tbl_extend("force", props, extra or {}))
end

-- A frame's graphic properties set (false: removed), in a style of its own: a copy of its
-- automatic style, or based on its common one.
function M.style_frame(ctx, frame, props)
  local base = style_named(ctx, "graphic", xml.attr(frame, "draw:style-name"))
  local style
  if base and base.automatic then
    style = xml.copy(base.el)
  else
    style = xml.elem("style:style", { { "style:family", "graphic" } })
    -- (based on its style: also one LibreOffice has built in, as "Frame", that the document doesn't list)
    xml.set_attr(style, "style:parent-style-name", base and xml.attr(base.el, "style:name") or xml.attr(frame, "draw:style-name"))
  end
  local name, n = nil, 1
  repeat
    name, n = "NeoGraphic" .. n, n + 1
  until not style_named(ctx, "graphic", name)
  xml.set_attr(style, "style:name", name)
  local el = xml.child(style, "style:graphic-properties") or xml.append(style, xml.elem("style:graphic-properties"))
  for attr, value in pairs(props) do
    xml.set_attr(el, attr, value or nil)
  end
  xml.append(automatic(ctx), style)
  ctx.named["graphic|" .. name] = { el = style, automatic = true }
  xml.set_attr(frame, "draw:style-name", name)
end

-- What a drawing or object that neoffice doesn't edit is, to show it by name (⟦chart⟧): nil for
-- others.
local SHAPES = { ["draw:custom-shape"] = "shape", ["draw:rect"] = "shape", ["draw:ellipse"] = "shape", ["draw:line"] = "shape",
  ["draw:path"] = "shape", ["draw:polygon"] = "shape", ["draw:polyline"] = "shape", ["draw:connector"] = "shape", ["draw:g"] = "drawing" }
function M.object_kind(ctx, el)
  if el.tag == "draw:frame" then
    local object = xml.child(el, "draw:object") or xml.child(el, "draw:object-ole")
    local href = object and (xml.attr(object, "xlink:href") or ""):gsub("^%./", "")
    local content = href and href ~= "" and zip.read(ctx.path, href .. "/content.xml") or ""
    return object and (content:find("<math", 1, true) and "equation" or content:find("<chart:chart", 1, true) and "chart" or "object")
      or xml.child(el, "draw:text-box") and "box" or nil
  end
  return SHAPES[el.tag] and (xml.find(el, "text:p") and "box" or SHAPES[el.tag]) or nil
end

-- ── Boxes (frames with text: see floats.lua and M.load) ─────────────────────

-- A box's paragraphs' container: a frame's text box (as LibreOffice makes a figure with its
-- caption), or a shape with text in it (as it reads Word's text boxes); if it holds only
-- paragraphs (no tables, no boxes in it).
function M.box_content(frame)
  local box = frame.tag == "draw:frame" and xml.child(frame, "draw:text-box")
    or frame.tag == "draw:custom-shape" and xml.child(frame, "text:p") and frame
  if not box or #xml.collect(box, { ["table:table"] = true, ["draw:text-box"] = true, ["draw:custom-shape"] = true }) > 0 then
    return nil
  end
  return box
end

function M.box_attrs(ctx, frame)
  local attrs = require("neoffice.floats").classify(M.frame_geometry(ctx, frame), M.page(ctx))
  local cm = markup.cm(xml.attr(frame, "svg:width"))
  attrs.width = cm and markup.width(cm) or nil
  return attrs
end

function M.box_container(frame)
  return frame.tag == "draw:custom-shape" and frame or xml.child(frame, "draw:text-box")
end

-- The box a paragraph is in: its frame, or nil.
function M.box_of(p)
  local up = p.parent
  while up do
    if up.tag == "draw:text-box" then
      return up.parent
    elseif up.tag == "draw:custom-shape" then
      return up
    end
    up = up.parent
  end
end

-- A new box: a rectangle with text in it (as LibreOffice reads Word's text boxes, and draws them
-- well), its paragraphs before its geometry.
function M.new_box(ctx, attrs)
  ctx.boxes_made = (ctx.boxes_made or 0) + 1
  local width = markup.cm(attrs.width or "", M.text_width(ctx)) or M.text_width(ctx) / 2
  local shape = xml.elem("draw:custom-shape", {
    -- (in a style based on "Frame": LibreOffice then shows the images and fields in it, as in a frame)
    { "draw:name", ("NeoBox%d_%d"):format(os.time(), ctx.boxes_made) }, { "draw:style-name", "Frame" }, { "text:anchor-type", "paragraph" },
    { "svg:width", length_cm(width) }, { "svg:height", "1cm" }, { "draw:z-index", "0" },
  }, {
    xml.elem("draw:enhanced-geometry", { { "svg:viewBox", "0 0 21600 21600" }, { "draw:type", "rectangle" }, { "draw:enhanced-path", "M 0 0 L 21600 0 21600 21600 0 21600 Z N" } }),
  })
  return { node = shape, content = shape }
end

function M.place_box(ctx, frame, attrs)
  local cm = attrs.width and markup.cm(attrs.width, M.text_width(ctx))
  if cm then
    xml.set_attr(frame, "svg:width", length_cm(cm))
  end
  local new = (xml.attr(frame, "draw:name") or ""):match("^NeoBox") -- (one neoffice made: no fill, line or padding)
  M.place_frame(ctx, frame, attrs, new and {
    ["draw:fill"] = "none", ["draw:stroke"] = "none", ["fo:padding"] = "0cm", ["draw:auto-grow-height"] = "true",
    ["draw:textarea-vertical-align"] = "top", ["fo:wrap-option"] = "wrap",
  } or nil)
end
function M.box_changed() end

-- A box as tall as its content (roughly: its images, and its text in lines), growing with it.
function M.fit_box(_, frame)
  local box = M.box_container(frame)
  if not box then
    return
  end
  local width = markup.cm(xml.attr(frame, "svg:width")) or 5
  local height = 0
  for _, p in ipairs(xml.collect(box, PARA)) do
    for _, image in ipairs(xml.collect(p, { ["draw:frame"] = true })) do
      height = height + (markup.cm(xml.attr(image, "svg:height")) or 0)
    end
    local chars = #(xml.text(p))
    height = height + math.ceil(chars / math.max(width / 0.19, 1)) * 0.5 + 0.2
  end
  if box == frame then -- (a shape: its height)
    xml.set_attr(frame, "svg:height", length_cm(math.max(height, 0.5)))
  else
    xml.set_attr(frame, "svg:height", nil)
    xml.set_attr(box, "fo:min-height", length_cm(math.max(height, 0.5)))
  end
end

function M.attach_box(_, frame, p)
  if frame.parent == p and xml.index(frame) then -- (in it still: not when its paragraph was written again)
    return
  end
  if frame.parent then
    xml.remove(frame)
  end
  xml.append(p, frame, 1)
end

local function grid_of(tbl) -- its rows of cells
  return vim.tbl_map(function(tr) return xml.collect(tr, { ["table:table-cell"] = true }) end, xml.collect(tbl, { ["table:table-row"] = true }))
end

-- The texts of a table's columns, for their widths: for each column, its cells' paragraphs (cells
-- spanning one column); and how many columns it has.
local function column_texts(tbl)
  local columns, n = {}, 0
  for _, row in ipairs(grid_of(tbl)) do
    local c = 1
    for _, cell in ipairs(row) do
      local span = tonumber(xml.attr(cell, "table:number-columns-spanned") or "") or 1
      if span == 1 then
        columns[c] = columns[c] or {}
        table.insert(columns[c], vim.tbl_map(xml.text, xml.collect(cell, { ["text:p"] = true, ["text:h"] = true }, { ["table:table"] = true })))
      end
      c = c + span
    end
    n = math.max(n, c - 1)
  end
  for c = 1, n do
    columns[c] = columns[c] or {}
  end
  return columns, n
end

local function font_size(ctx) -- the size of the document's text (pt)
  for _, style in ipairs(xml.collect(ctx.styles_doc, { ["style:default-style"] = true, ["style:style"] = true })) do
    local name = xml.attr(style, "style:name")
    if xml.attr(style, "style:family") == "paragraph" and (style.tag == "style:default-style" or name == "Standard") then
      local props = xml.child(style, "style:text-properties")
      local size = props and xml.attr(props, "fo:font-size")
      local pt = size and tonumber(size:match("^([%d.]+)pt$") or "")
      if pt then
        return pt
      end
    end
  end
  return 12
end
local function auto_widths(ctx, tbl, width)
  return require("neoffice.tables").widths((column_texts(tbl)), width, font_size(ctx) / 72 * 2.54, M.text_width(ctx))
end

-- A table's look (see tables.lua): { rules (a preset of tables.BORDERS, if its lines are one),
-- width, height (if its cell padding is one of neoffice's), position, typeset (by neoffice) }.
function M.table_format(ctx, tbl)
  local tables = require("neoffice.tables")
  local paddings = {}
  local grid = vim.tbl_map(function(row)
    return vim.tbl_map(function(cell)
      local style = style_named(ctx, "table-cell", xml.attr(cell, "table:style-name"))
      local props = style and xml.child(style.el, "style:table-cell-properties") or { attrs = {} }
      local sides = {}
      for _, side in ipairs({ "top", "bottom", "left", "right" }) do
        local value = xml.attr(props, "fo:border-" .. side) or xml.attr(props, "fo:border") or "none"
        local width = markup.cm(value:match("^%s*([%d.]+%a+)") or "")
        sides[side] = (value == "none" or not width or width == 0) and 0 or width >= 0.026 and 2 or 1 -- (thick from 3/4 pt)
      end
      for _, side in ipairs({ "top", "bottom" }) do
        paddings[#paddings + 1] = markup.cm(xml.attr(props, "fo:padding-" .. side) or xml.attr(props, "fo:padding") or "") or -1
      end
      return sides
    end, row)
  end, grid_of(tbl))
  local format = { rules = tables.borders_of(grid), width = "auto", position = "left" }
  local same = #paddings > 0
  for _, p in ipairs(paddings) do
    same = same and math.abs(p - paddings[1]) < 0.006
  end
  format.height = same and tables.height_of(paddings[1]) or nil
  local style = style_named(ctx, "table", xml.attr(tbl, "table:style-name"))
  local props = style and xml.child(style.el, "style:table-properties")
  if props then
    local align, rel, width = xml.attr(props, "table:align"), xml.attr(props, "style:rel-width"), markup.cm(xml.attr(props, "style:width"))
    format.position = (align == "center" or align == "right") and align or "left"
    if align == "margins" or rel == "100%" then
      format.width, format.position = "full", tables.defaults().position -- (where it is doesn't matter)
    elseif rel then
      format.width = rel
    elseif width then -- (as wide as its content: auto; else that wide)
      local sum = 0
      for _, w in ipairs(auto_widths(ctx, tbl, "auto")) do
        sum = sum + w
      end
      format.width = math.abs(sum - width) < 0.02 * width + 0.05 and "auto" or markup.width(width)
    end
  end
  format.typeset = format.rules ~= nil and format.height ~= nil -- (by neoffice: it keeps it so)
  return format
end

-- An element's style changed: these properties (of `tag`, the style's properties element) set, or
-- removed (false). The same change of the same style makes one style.
local function restyle(ctx, el, family, tag, attrs)
  ctx.restyled = ctx.restyled or {}
  local parts = {}
  for _, a in ipairs(attrs) do
    parts[#parts + 1] = a[1] .. "=" .. tostring(a[2])
  end
  local key = family .. "|" .. (xml.attr(el, "table:style-name") or "") .. "|" .. table.concat(parts, ";")
  if ctx.restyled[key] then
    return xml.set_attr(el, "table:style-name", ctx.restyled[key])
  end
  local props = own_style(ctx, el, family, tag)
  for _, a in ipairs(attrs) do
    xml.set_attr(props, a[1], a[2] or nil)
  end
  ctx.restyled[key] = xml.attr(el, "table:style-name")
end

-- A paragraph's style with these paragraph properties: "compact" (without spacing or indentation,
-- single-spaced: in a typeset table's cell) or "keep" (kept with the next: a table's caption).
local PARA_PROPS = {
  compact = { ["fo:margin-top"] = "0cm", ["fo:margin-bottom"] = "0cm", ["fo:margin-left"] = "0cm", ["fo:margin-right"] = "0cm", ["fo:text-indent"] = "0cm", ["fo:line-height"] = "100%" },
  keep = { ["fo:keep-with-next"] = "always" },
}
local function para_style(ctx, p, kind)
  local name = xml.attr(p, "text:style-name")
  local key = kind .. "|" .. (name or "")
  if not ctx.made[key] then
    local auto = ctx.auto_paras[name or ""]
    local style = auto and xml.copy(auto) or xml.elem("style:style", { { "style:family", "paragraph" } })
    if not auto and name then
      xml.set_attr(style, "style:parent-style-name", name)
    end
    local new, n = nil, 1
    repeat
      new, n = "NeoP" .. n, n + 1
    until ctx.parent_style[new] == nil
    xml.set_attr(style, "style:name", new)
    local props = xml.child(style, "style:paragraph-properties") or xml.append(style, xml.elem("style:paragraph-properties"), 1)
    for attr, value in pairs(PARA_PROPS[kind]) do
      xml.set_attr(props, attr, value)
    end
    xml.append(automatic(ctx), style)
    ctx.made[key], ctx.auto_paras[new] = new, style
    ctx.parent_style[new] = xml.attr(style, "style:parent-style-name") or false
    if ctx.para_style_el then -- (known as a paragraph style: its font, see font_base)
      ctx.para_style_el[new] = style
    end
    ctx.made[kind .. "|" .. new] = new -- (already so)
  end
  xml.set_attr(p, "text:style-name", ctx.made[key])
end

-- Kept on the same page as the paragraph after it (a table's caption, a figure).
function M.keep_with_next(ctx, p)
  para_style(ctx, p, "keep")
end

-- Typeset a table with these options (see tables.lua; those not given stay as they are): its rules
-- (on every cell), column widths (from its content), row height (cell padding, and paragraphs
-- without spacing), position; and with `typeset`, its first row repeated on each page and at the
-- bottom of its cells, and rows that don't break across pages.
local LINE = { [0] = "none", "0.5pt solid #000000", "1pt solid #000000" }
local function cm(n) -- (a length, to a hundredth of a millimetre)
  return (("%.3f"):format(n):gsub("0+$", ""):gsub("%.$", "")) .. "cm"
end
function M.set_table_format(ctx, tbl, set)
  local tables = require("neoffice.tables")
  local grid = grid_of(tbl)
  for r, row in ipairs(grid) do
    for c, cell in ipairs(row) do
      local attrs = {}
      if set.rules then
        local sides = tables.cell_borders(set.rules, r, #grid, c, #row)
        attrs[#attrs + 1] = { "fo:border", false }
        for _, side in ipairs({ "top", "bottom", "left", "right" }) do
          attrs[#attrs + 1] = { "fo:border-" .. side, LINE[sides[side]] }
        end
      end
      if set.height then
        local tb, lr = cm(tables.HEIGHT[set.height]), cm(tables.PADDING)
        vim.list_extend(attrs, { { "fo:padding", false }, { "fo:padding-top", tb }, { "fo:padding-bottom", tb }, { "fo:padding-left", lr }, { "fo:padding-right", lr } })
      end
      if set.typeset then
        attrs[#attrs + 1] = { "style:vertical-align", r == 1 and #grid > 1 and "bottom" or false }
      end
      if #attrs > 0 then
        restyle(ctx, cell, "table-cell", "style:table-cell-properties", attrs)
      end
    end
  end
  if set.height then
    for _, p in ipairs(xml.collect(tbl, { ["text:p"] = true }, { ["table:table"] = true })) do
      para_style(ctx, p, "compact")
    end
  end
  if set.width or set.position then
    local width = set.width or M.table_format(ctx, tbl).width
    local widths = auto_widths(ctx, tbl, width)
    local sum = 0
    for _, w in ipairs(widths) do
      sum = sum + w
    end
    local percent = tonumber(width:match("^([%d.]+)%%$") or "")
    restyle(ctx, tbl, "table", "style:table-properties", {
      { "style:width", cm(sum) },
      { "style:rel-width", width == "full" and "100%" or percent and width or false },
      { "table:align", width == "full" and "margins" or set.position or M.table_format(ctx, tbl).position },
      { "fo:margin-left", false }, { "fo:margin-right", false },
    })
    if set.width then -- the columns: one each, as wide as computed
      local old = xml.collect(tbl, { ["table:table-column"] = true }, { ["table:table-row"] = true, ["table:table-rows"] = true, ["table:table-header-rows"] = true })
      local parent = old[1] and old[1].parent or tbl
      local at = old[1] and xml.index(old[1]) or 1
      for _, col in ipairs(old) do
        xml.remove(col)
      end
      for k, w in ipairs(widths) do
        local col = xml.append(parent, xml.elem("table:table-column"), at + k - 1)
        restyle(ctx, col, "table-column", "style:table-column-properties", { { "style:column-width", cm(w) } })
      end
    end
  end
  -- the table styles neoffice made that nothing uses any more (those of earlier saves)
  local used = {}
  local function walk(n)
    for _, c in ipairs(n.children or {}) do
      if c.tag then
        if c.tag:match("^table:") then
          used[xml.attr(c, "table:style-name") or ""] = true
        end
        walk(c)
      end
    end
  end
  walk(xml.find(ctx.root, "office:body") or {})
  for _, style in ipairs(vim.tbl_filter(function(c) return c.tag == "style:style" end, automatic(ctx).children)) do
    local name, family = xml.attr(style, "style:name") or "", xml.attr(style, "style:family") or ""
    if family:match("^table") and name:match("^Neo%a+%d+$") and not used[name] then
      xml.remove(style)
      ctx.named[family .. "|" .. name] = nil
    end
  end
  ctx.restyled = nil
  if set.typeset then
    -- the first row in the header rows (repeated on each page), and only it
    local rows = xml.collect(tbl, { ["table:table-row"] = true })
    local first = rows[1]
    if first then
      local head = xml.child(tbl, "table:table-header-rows")
      if not head then
        local anchor = first.parent == tbl and first or first.parent
        head = xml.append(tbl, xml.elem("table:table-header-rows"), xml.index(anchor))
      end
      if first.parent ~= head then
        xml.remove(first)
        xml.append(head, first, 1)
      end
      local moved = vim.tbl_filter(function(r) return r ~= first and r.parent == head end, rows)
      for k = #moved, 1, -1 do -- (after it, in order)
        xml.remove(moved[k])
        local nxt = tbl.children[xml.index(head) + 1]
        if nxt and nxt.tag == "table:table-rows" then
          xml.append(nxt, moved[k], 1)
        else
          xml.append(tbl, moved[k], xml.index(head) + 1)
        end
      end
      for _, row in ipairs(rows) do -- (rows that don't break across pages)
        restyle(ctx, row, "table-row", "style:table-row-properties", { { "fo:keep-together", "always" } })
      end
      local before = tbl.parent and tbl.parent.children[xml.index(tbl) - 1] -- (its caption: kept with it)
      if before and (before.tag == "text:p" or before.tag == "text:h") then
        para_style(ctx, before, "keep")
      end
      -- (a short table on one page)
      restyle(ctx, tbl, "table", "style:table-properties", { { "style:may-break-between-rows", #rows > tables.KEEP_ROWS and "true" or "false" } })
    end
  end
end

-- As many columns as the widest row has (after rows or cells were added or removed).
-- A row with cells added (see docx.fit_row): in an .odt, a merged cell is followed by the cells it
-- covers (table:covered-table-cell); a new cell goes after those.
function M.fit_row(_, tbl, tr, new)
  local function count(el, attr)
    return tonumber(xml.attr(el, attr) or "") or 1
  end
  local columns = 0
  for _, col in ipairs(xml.collect(tbl, { ["table:table-column"] = true }, { ["table:table-row"] = true, ["table:table-rows"] = true })) do
    columns = columns + count(col, "table:number-columns-repeated")
  end
  -- (each new cell after the cells covered by the cell of the document's before it, in order)
  local groups, anchors, anchor = {}, {}, false
  for _, c in ipairs(vim.list_slice(tr.children)) do
    if new[c] then
      xml.set_attr(c, "table:number-columns-spanned", nil)
      xml.set_attr(c, "table:number-rows-spanned", nil)
      if not groups[anchor] then
        groups[anchor], anchors[#anchors + 1] = {}, anchor
      end
      table.insert(groups[anchor], c)
      xml.remove(c)
    elseif c.tag == "table:table-cell" then
      anchor = c
    end
  end
  for _, a in ipairs(anchors) do
    local at = a and xml.index(a) + count(a, "table:number-columns-spanned") or 1
    for i, c in ipairs(groups[a]) do
      xml.append(tr, c, at + i - 1)
    end
  end
  local function total()
    local n = 0
    for _, c in ipairs(tr.children) do
      if c.tag == "table:table-cell" or c.tag == "table:covered-table-cell" then
        n = n + count(c, "table:number-columns-repeated")
      end
    end
    return n
  end
  if columns == 0 then
    return
  end
  for k = #tr.children, 1, -1 do -- (padding: new empty cells, from the end)
    if total() <= columns then
      break
    end
    local c = tr.children[k]
    if new[c] and not new[c].filled then
      xml.remove(c)
    end
  end
  local k = 1
  while k <= #tr.children and total() > columns do -- (a new cell with text: a column of the merged cell before it)
    local c = tr.children[k]
    if new[c] then
      for j = k - 1, 1, -1 do
        local before = tr.children[j]
        if before.tag == "table:table-cell" and not new[before] and count(before, "table:number-columns-spanned") > 1 then
          local n = count(before, "table:number-columns-spanned") - 1
          xml.set_attr(before, "table:number-columns-spanned", n > 1 and tostring(n) or nil)
          local last = j + n -- (its last covered cell: no longer covered)
          if tr.children[last] and tr.children[last].tag == "table:covered-table-cell" then
            xml.remove(tr.children[last])
            k = k - 1
          end
          break
        end
      end
    end
    k = k + 1
  end
end

function M.fix_grid(_, tbl)
  local n, have = 0, 0
  local function count(el, attr)
    return tonumber(xml.attr(el, attr) or "") or 1
  end
  for _, tr in ipairs(xml.collect(tbl, { ["table:table-row"] = true })) do
    local cells = 0
    for _, c in ipairs(xml.collect(tr, { ["table:table-cell"] = true, ["table:covered-table-cell"] = true })) do
      cells = cells + count(c, "table:number-columns-repeated")
    end
    n = math.max(n, cells)
  end
  local cols = xml.collect(tbl, { ["table:table-column"] = true }, { ["table:table-row"] = true, ["table:table-rows"] = true })
  for _, col in ipairs(cols) do
    have = have + count(col, "table:number-columns-repeated")
  end
  if have < n and #cols > 0 then
    local last = cols[#cols]
    local r = count(last, "table:number-columns-repeated") + n - have
    xml.set_attr(last, "table:number-columns-repeated", tostring(r))
  end
  while have > math.max(n, 1) and #cols > 0 do
    local last = cols[#cols]
    local r = count(last, "table:number-columns-repeated")
    local drop = math.min(r, have - n)
    if drop == r then
      xml.remove(table.remove(cols))
    else
      xml.set_attr(last, "table:number-columns-repeated", r - drop > 1 and tostring(r - drop) or nil)
    end
    have = have - drop
  end
end

function M.fill(ctx, p, block, pieces)
  -- what kind of paragraph it is: heading, title, list item or plain
  local level, old_level = block.heading, p.tag == "text:h" and (tonumber(xml.attr(p, "text:outline-level")) or 1) or nil
  if level ~= old_level then
    p.tag = level and "text:h" or "text:p"
    xml.set_attr(p, "text:outline-level", level and tostring(level))
    xml.set_attr(p, "text:style-name", level and (ctx.headings[level] or "Heading_20_" .. level))
  end
  if (block.title or false) ~= is_title(ctx, p) and not level then
    xml.set_attr(p, "text:style-name", block.title and "Title" or nil)
  end
  if not not block.quote ~= based_on(ctx, p, QUOTES) and not (level or block.title) then
    xml.set_attr(p, "text:style-name", block.quote and (ctx.quote_style or "Quotations") or nil)
  end
  -- a caption: in a caption style (and no longer, if it had a caption's number and isn't one now)
  local in_caption = based_on(ctx, p, CAPTIONS)
  if block.caption and not in_caption then
    xml.set_attr(p, "text:style-name", caption_style(ctx, block.caption))
  elseif not block.caption and in_caption and #xml.collect(p, { ["text:sequence"] = true }) > 0 and not (level or block.title or block.quote) then
    xml.set_attr(p, "text:style-name", nil)
  end
  local page = vim.iter(pieces):any(function(piece) return piece.page_break end) -- (.odt: before it)
  set_layout(ctx, p, block.align, page)
  if block.para then -- (its own format's attributes: a plain paragraph's, see paras.lua)
    set_para(ctx, p, block.para)
  end
  set_list(ctx, p, block.list)
  p.children = {}

  local base, inherit = M.font_base(ctx, p), M.font_inherited(ctx, p) -- (its named style's; what its text shows)
  local cur, after_space = { el = p }, true -- cur: the open <text:a>, if any
  for _, piece in ipairs(pieces) do
    if piece.text then
      if piece.link ~= cur.link then
        local attrs = ctx.link_attrs[piece.link] or { { "xlink:type", "simple" }, { "xlink:href", piece.link } }
        local a = piece.link and xml.elem("text:a", vim.deepcopy(attrs))
        cur = { link = piece.link, el = a and xml.append(p, a) or p }
      end
      local target = cur.el
      for _, style in ipairs(style_for(ctx, piece, base, inherit)) do -- (nested spans, outermost first)
        target = xml.append(target, xml.elem("text:span", { { "text:style-name", style } }))
      end
      after_space = put_text(target, piece.text, after_space)
    else
      cur = { el = p }
      local ch = piece.insert_start or piece.insert_end or piece.deletion
      local c = piece.comment_start
      if piece.part and piece.part > 0 then -- (another part of a deletion: a change of its own)
        local key = tostring(p) .. ":" .. piece.part
        ch.parts = ch.parts or {}
        ch.parts[key] = ch.parts[key] or { type = ch.type, author = ch.author, date = ch.date, parent = ch }
        ch = ch.parts[key]
      end
      if piece.deletion and ch.region and ch.text ~= piece.deleted then -- (its text changed: split, edited)
        local kind = xml.find(ch.region, "text:deletion")
        for _, c in ipairs(vim.tbl_filter(function(c) return PARA[c.tag] end, kind.children)) do
          xml.remove(c)
        end
        for _, line in ipairs(vim.split(piece.deleted, "\n")) do
          xml.append(kind, xml.elem("text:p", {}, { { text = line } }))
        end
      end
      if piece.deletion then
        ch.text = piece.deleted
      end
      if ch then
        region(ctx, ch)
        local tag = piece.deletion and "text:change" or piece.insert_start and "text:change-start" or "text:change-end"
        xml.append(p, xml.elem(tag, { { "text:change-id", ch.id } }))
      elseif piece.object and not piece.object.tag then -- (a field: its elements)
        for _, el in ipairs(piece.object.children) do
          xml.append(p, xml.copy(el))
        end
      elseif piece.object then
        xml.append(p, piece.object)
      elseif piece.line_break then
        xml.append(p, xml.elem("text:line-break"))
        after_space = true
      elseif piece.footnote then
        xml.append(p, footnote_node(ctx, piece.footnote))
      elseif piece.image then
        xml.append(p, image_node(ctx, piece.image))
      elseif c then
        c.name = c.name or ("__Annotation__neo%d_%d"):format(os.time(), math.random(1e6))
        xml.append(p, annotation_el(c))
        for _, reply in ipairs(c.replies) do
          xml.append(p, annotation_el(reply, c.name))
        end
      elseif piece.comment_end then
        xml.append(p, xml.elem("office:annotation-end", { { "office:name", piece.comment_end.name } }))
      end
    end
  end
end

-- Namespaces of the elements we may write; files from other tools don't always declare them.
local NAMESPACES = {
  { "dc", "http://purl.org/dc/elements/1.1/" },
  { "draw", "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" },
  { "fo", "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" },
  { "loext", "urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0" },
  { "style", "urn:oasis:names:tc:opendocument:xmlns:style:1.0" },
  { "svg", "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" },
  { "table", "urn:oasis:names:tc:opendocument:xmlns:table:1.0" },
  { "xlink", "http://www.w3.org/1999/xlink" },
}

function M.save(ctx, live)
  -- tracked changes accepted or rejected (no longer in the buffer) lose their region
  for _, ch in pairs(ctx.changes) do
    if ch.region and ch.region.parent and not live.changes[ch] and not (ch.parent and live.changes[ch.parent]) then
      xml.remove(ch.region)
      ctx.dirty[M.MAIN] = true
    end
  end
  local entries, trees = {}, {}
  if ctx.dirty[M.MAIN] and ctx.auto then -- (neoffice's own automatic styles nothing uses: gone)
    local used = {}
    local function walk(el, skip)
      for _, kv in ipairs(el.attrs or {}) do
        if kv[1] ~= "style:name" or not skip then
          used[kv[2]] = true
        end
      end
      for _, c in ipairs(el.children or {}) do
        if c.tag then
          walk(c, false)
        end
      end
    end
    walk(xml.find(ctx.root, "office:body") or {}, false)
    for _, style in ipairs(ctx.auto.children) do
      if style.tag then
        walk(style, true) -- (what a style refers to: its parent, a list style...; not its own name)
      end
    end
    for _, style in ipairs(vim.list_slice(ctx.auto.children)) do
      local name = style.tag == "style:style" and xml.attr(style, "style:name")
      if name and name:match("^Neo") and not used[name] then
        xml.remove(style)
        ctx.auto_paras[name] = nil
        if ctx.named then
          ctx.named[(xml.attr(style, "style:family") or "") .. "|" .. name] = nil
        end
      end
    end
  end
  if ctx.dirty[M.MAIN] then
    local root = xml.find(ctx.root, "office:document-content")
    for _, ns in ipairs(NAMESPACES) do
      xml.set_attr(root, "xmlns:" .. ns[1], xml.attr(root, "xmlns:" .. ns[1]) or ns[2])
    end
    entries[M.MAIN], trees[M.MAIN] = xml.serialize(ctx.root), ctx.root
  end
  if ctx.dirty[MANIFEST] then
    entries[MANIFEST], trees[MANIFEST] = xml.serialize(ctx.manifest), ctx.manifest
  end
  if ctx.dirty[STYLES] then
    entries[STYLES], trees[STYLES] = xml.serialize(ctx.styles_doc), ctx.styles_doc
  end
  for name, bytes in pairs(ctx.files) do
    entries[name] = bytes
  end
  return entries, trees -- (trees: the XML each part was written from, to check: see init.lua)
end

return M
