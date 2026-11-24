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

-- The formatting { b, i, u, sup, sub } of a text style, following its parent styles, starting
-- from `inherited` (that of the enclosing text).
local function flags(ctx, name, inherited)
  local f = vim.deepcopy(inherited or {})
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
    f.h = background ~= "transparent"
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

-- An image: its file in the package and its alt text (description, or else title).
local function image_of(frame)
  if xml.child(frame, "draw:object") then -- a formula or chart (its draw:image is only a preview)
    return nil
  end
  local image = xml.child(frame, "draw:image")
  local path = image and xml.attr(image, "xlink:href")
  if path and not path:match("^%a+:") and not path:match("^%.%./") then
    local alt = xml.child(frame, "svg:desc") or xml.child(frame, "svg:title")
    local cm = markup.cm(xml.attr(frame, "svg:width"))
    return { path = path, alt = alt and vim.trim(xml.text(alt):gsub("%s+", " ")) or "", width = cm and markup.width(cm) or nil }
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

-- The list a paragraph is an item of: { kind = "bullet" | "number", level = 1.., id = its outermost list }.
local function list_of(ctx, p)
  local level, style, top, node = 0, nil, nil, p.parent
  while node and LISTS[node.tag] do
    if node.tag == "text:list-header" then -- (an unnumbered heading of the list)
      return nil
    elseif node.tag == "text:list" then
      level, top = level + 1, node
      style = style or xml.attr(node, "text:style-name")
    end
    node = node.parent
  end
  return top and { kind = list_kind(ctx, style, level), level = level, id = top }
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

-- A paragraph's own alignment and whether it starts on a new page (from its automatic style).
local function layout(ctx, p)
  local style = ctx.auto_paras[xml.attr(p, "text:style-name") or ""]
  local props = style and xml.child(style, "style:paragraph-properties") or { attrs = {} }
  return ALIGN[xml.attr(props, "fo:text-align") or ""], xml.attr(props, "fo:break-before") == "page"
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
  block.align = layout(ctx, p)
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
        add(vim.tbl_extend("force", flags(ctx, nil, f), { text = s, props = f.style, link = f.link }))
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
      elseif el.text then -- ODF collapses whitespace, also across elements; real spaces are <text:s/>
        local s = el.text:gsub("%s+", " ")
        text(after_space and s:gsub("^ ", "") or s)
        after_space = s:sub(-1) == " " or (after_space and s == "")
      elseif tag == "text:span" then
        local style = xml.attr(el, "text:style-name")
        local outer = type(f.style) == "table" and f.style or { f.style }
        local chain = style and vim.list_extend(vim.list_slice(outer), { style }) or outer
        walk(el, vim.tbl_extend("force", flags(ctx, style, f), { style = #chain > 1 and chain or chain[1], link = f.link }))
      elseif tag == "text:a" then
        local url = xml.attr(el, "xlink:href")
        ctx.link_attrs[url] = ctx.link_attrs[url] or el.attrs
        walk(el, vim.tbl_extend("force", f, { link = url }))
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
      elseif tag then
        local image = tag == "draw:frame" and image_of(el) or nil
        local text = not tag:match("^draw:") and plain(el) or ""
        add({ object = el, visible = VISIBLE[tag] or xml.text(el) ~= "", image = image, field = text ~= "" and { text = text } or nil })
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
    paras[#paras + 1] = { node = p, block = block, spans = spans }
  end
  for _, c in pairs(ctx.comments) do
    c.was_resolved = c.resolved
  end
  -- (or the headings are numbered by being in a list)
  ctx.numbered = ctx.numbered or vim.iter(paras):any(function(para) return para.node.tag == "text:h" and list_of(ctx, para.node) end)
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

local function automatic(ctx) -- <office:automatic-styles> (created if needed)
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
local function set_list(ctx, p, want)
  local have = list_of(ctx, p)
  local own_item = not have or element_children(p.parent)[1] == p -- (a new paragraph may share one)
  if own_item and (have and have.kind) == (want and want.kind) and (have and have.level) == (want and want.level) then
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
    local container = prev and prev.tag == "text:list" and list_kind(ctx, xml.attr(prev, "text:style-name"), 1) == want.kind and prev
    container = container or xml.append(p.parent, xml.elem("text:list", { { "text:style-name", list_style(ctx, want.kind) } }), xml.index(p))
    local top = container
    for _ = 2, want.level do -- down to its level, in the last item
      local items = element_children(container)
      local item = items[#items] and items[#items].tag == "text:list-item" and items[#items] or xml.append(container, xml.elem("text:list-item"))
      local inside = element_children(item)
      container = inside[#inside] and inside[#inside].tag == "text:list" and inside[#inside] or xml.append(item, xml.elem("text:list"))
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
local function style_for(ctx, piece)
  local chain = type(piece.props) == "table" and piece.props or { piece.props }
  local outer, inner, inherited = vim.list_slice(chain, 1, #chain - 1), chain[#chain], {}
  for _, name in ipairs(outer) do
    inherited = flags(ctx, name, inherited)
  end
  local have = flags(ctx, inner, inherited)
  if vim.iter(FLAGS):all(function(k) return have[k] == piece[k] end) then
    return chain
  end
  local key = table.concat(vim.tbl_map(function(k) return tostring(piece[k]) end, FLAGS), "|") .. "|" .. (inner or "")
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
    xml.set_attr(props, "fo:background-color", piece.h and "#ffff00" or "transparent")
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
  img.node = frame
  return frame
end

-- Give a paragraph this alignment and page break (when they changed), through an automatic style.
local ALIGN_VALUE = { center = "center", right = "end", justify = "justify" }
local function set_layout(ctx, p, align, page)
  local have_align, have_page = layout(ctx, p)
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
    if have_align ~= align then
      xml.set_attr(props, "fo:text-align", ALIGN_VALUE[align])
    end
    xml.set_attr(props, "fo:break-before", page and "page" or nil)
    xml.append(automatic(ctx), style)
    ctx.made[key], ctx.auto_paras[new] = new, style
    ctx.parent_style[new] = xml.attr(style, "style:parent-style-name") or false
  end
  xml.set_attr(p, "text:style-name", ctx.made[key])
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

function M.new_table(ctx) -- one cell, with borders (rows and cells are added by copying)
  if not ctx.styles_made_cell then
    xml.append(automatic(ctx), xml.elem("style:style", { { "style:name", "NeoCell" }, { "style:family", "table-cell" } }, {
      xml.elem("style:table-cell-properties", { { "fo:padding", "0.04in" }, { "fo:border", "0.5pt solid #000000" } }),
    }))
    ctx.styles_made_cell = true
  end
  ctx.tables_made = (ctx.tables_made or 0) + 1
  return xml.elem("table:table", { { "table:name", ("NeoTable%d_%d"):format(os.time(), ctx.tables_made) } }, {
    xml.elem("table:table-column"),
    xml.elem("table:table-row", {}, {
      xml.elem("table:table-cell", { { "table:style-name", "NeoCell" }, { "office:value-type", "string" } }, { xml.elem("text:p") }),
    }),
  })
end

-- As many columns as the widest row has (after rows or cells were added or removed).
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
  local page = vim.iter(pieces):any(function(piece) return piece.page_break end) -- (.odt: before it)
  set_layout(ctx, p, block.align, page)
  set_list(ctx, p, block.list)
  p.children = {}

  local cur, after_space = { el = p }, true -- cur: the open <text:a>, if any
  for _, piece in ipairs(pieces) do
    if piece.text then
      if piece.link ~= cur.link then
        local attrs = ctx.link_attrs[piece.link] or { { "xlink:type", "simple" }, { "xlink:href", piece.link } }
        local a = piece.link and xml.elem("text:a", vim.deepcopy(attrs))
        cur = { link = piece.link, el = a and xml.append(p, a) or p }
      end
      local target = cur.el
      for _, style in ipairs(style_for(ctx, piece)) do -- (nested spans, outermost first)
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
  local entries = {}
  if ctx.dirty[M.MAIN] then
    local root = xml.find(ctx.root, "office:document-content")
    for _, ns in ipairs(NAMESPACES) do
      xml.set_attr(root, "xmlns:" .. ns[1], xml.attr(root, "xmlns:" .. ns[1]) or ns[2])
    end
    entries[M.MAIN] = xml.serialize(ctx.root)
  end
  if ctx.dirty[MANIFEST] then
    entries[MANIFEST] = xml.serialize(ctx.manifest)
  end
  if ctx.dirty[STYLES] then
    entries[STYLES] = xml.serialize(ctx.styles_doc)
  end
  for name, bytes in pairs(ctx.files) do
    entries[name] = bytes
  end
  return entries
end

return M
