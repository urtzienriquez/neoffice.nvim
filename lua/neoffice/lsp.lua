-- neoffice_ls: a language server for documents, running in Neovim itself (vim.lsp.start with a
-- function as its command): nothing to install. It reads the buffer as it is, and answers:
--   textDocument/documentSymbol  the title and the headings, nested (gO, pickers, outlines)
--   textDocument/foldingRange    each heading's section
--   textDocument/hover           the comment, or the tracked change, at the cursor
--   diagnostics                  attributes in braces (a table's options, an image's width): an
--                                error if one isn't any, a hint if not written "name = value, ..."
--   textDocument/codeAction      for those hints: write them so
--   textDocument/formatting      tables aligned (with their line of dashes), attributes tidy
--   (and rangeFormatting)        (see format.lua)
--   textDocument/completion      a table's options and their values, an image's width, a
--                                highlight's colour
local api = vim.api
local M = { NAME = "neoffice_ls" }

-- The fonts installed (their families' names, from fontconfig: once a session, asked for in the
-- background when a document opens; none without it).
local installed, asking
local function families(out)
  local list, seen = {}, {}
  for line in (out or ""):gmatch("[^\n]+") do
    for family in line:gmatch("[^,]+") do
      family = vim.trim(family)
      if family ~= "" and not seen[family] and not family:find("\\") then
        seen[family], list[#list + 1] = true, family
      end
    end
  end
  table.sort(list)
  return list
end
function M.prefetch_fonts()
  if installed or asking then
    return
  end
  local ok, job = pcall(vim.system, { "fc-list", ":", "family" }, { text = true }, function(r)
    vim.schedule(function()
      installed = installed or (r.code == 0 and families(r.stdout) or {})
    end)
  end)
  asking = ok and job or nil
end
function M.fonts()
  if not installed then
    if asking then -- (asked already: its answer)
      local r = asking:wait(2000)
      installed = r and r.code == 0 and families(r.stdout) or {}
    else
      local ok, r = pcall(function() return vim.system({ "fc-list", ":", "family" }, { text = true }):wait() end)
      installed = ok and r.code == 0 and families(r.stdout) or {}
    end
  end
  return installed
end

local function neo()
  return require("neoffice")
end

local function plain(text) -- (a paragraph's text, without its marks)
  local _, chars = require("neoffice.markup").parse(text)
  return vim.trim(table.concat(vim.tbl_map(function(c) return c.ch end, chars)):gsub("%s+", " "))
end

-- The title and the headings: { name, level (0: the title), first, last (0-based rows: its section) };
-- with `captions`, the captions too ("Table 1: ...", level: math.huge, in the section they are in).
local function headings(buf, captions)
  local doc, list, count = neo().docs[buf], {}, {}
  local blocks = neo().blocks(buf)
  for _, para in ipairs(blocks) do
    local block = require("neoffice.markup").block(para.text)
    if captions and block.caption and not para.cell then
      local key = "caption." .. block.caption
      count[key] = (count[key] or 0) + 1
      local label = require("neoffice.markup").captions()[block.caption]
      list[#list + 1] = { name = label .. " " .. count[key] .. ": " .. plain(para.text), level = math.huge, first = para.first - 1, line_end = para.last - 1, caption = true }
    end
    if (block.heading or block.title) and not para.cell then
      local name = plain(para.text)
      if block.heading and doc and doc.ctx.numbered then -- (with its number, as shown)
        count[block.heading] = (count[block.heading] or 0) + 1
        for deeper = block.heading + 1, 9 do
          count[deeper] = nil
        end
        local number = {}
        for level = 1, block.heading do
          number[level] = count[level] or 0
        end
        name = table.concat(number, ".") .. " " .. name
      end
      list[#list + 1] = { name = name, level = block.heading or 0, first = para.first - 1, line_end = para.last - 1 }
    end
  end
  local last_row = api.nvim_buf_line_count(buf) - 1
  for k, h in ipairs(list) do -- (a section goes on until the next heading of its level or above)
    h.last = h.caption and h.line_end or last_row
    for j = k + 1, h.caption and 0 or #list do -- (a caption: only its own lines)
      if list[j].level <= h.level then
        h.last = list[j].first - 1
        break
      end
    end
  end
  return list
end

local function range(first, last, buf)
  local line = api.nvim_buf_get_lines(buf, last, last + 1, false)[1] or ""
  return { start = { line = first, character = 0 }, ["end"] = { line = last, character = #line } }
end

local function symbols(buf)
  local top, stack = {}, {}
  for _, h in ipairs(headings(buf, true)) do
    local symbol = {
      name = h.name ~= "" and h.name or "(untitled)",
      kind = h.caption and vim.lsp.protocol.SymbolKind.Object or h.level == 0 and vim.lsp.protocol.SymbolKind.File or vim.lsp.protocol.SymbolKind.String,
      range = range(h.first, math.max(h.last, h.line_end), buf),
      selectionRange = range(h.first, h.line_end, buf),
      children = {},
    }
    while #stack > 0 and stack[#stack].level >= h.level do
      table.remove(stack)
    end
    table.insert(#stack > 0 and stack[#stack].symbol.children or top, symbol)
    stack[#stack + 1] = { level = h.level, symbol = symbol }
  end
  return top
end

local function folds(buf)
  local out = {}
  local header = require("neoffice.markup").header(api.nvim_buf_get_lines(buf, 0, 200, false))
  if header then -- (the styles at the top: folded out of the way, if you like)
    out[#out + 1] = { startLine = 0, endLine = header.last - 1, kind = "region" }
  end
  for _, h in ipairs(headings(buf)) do
    local last = h.last
    while last > h.first and (api.nvim_buf_get_lines(buf, last, last + 1, false)[1] or "") == "" do
      last = last - 1 -- (not the blank lines before the next heading)
    end
    if last > h.first then
      out[#out + 1] = { startLine = h.first, endLine = last, kind = "region" }
    end
  end
  return out
end

local function covers(m, row, col)
  local d = m[4]
  return (m[2] < row or m[2] == row and m[3] <= col) and (d.end_row > row or d.end_row == row and d.end_col >= col)
end

local function hover(buf, pos)
  local doc, out = neo().docs[buf], {}
  if not doc then
    return nil
  end
  local row, col = pos.line, pos.character
  for _, m in ipairs(api.nvim_buf_get_extmarks(buf, neo().ns.comment, { row, 0 }, { row, -1 }, { details = true, overlap = true })) do
    local c = doc.comment[m[1]]
    if c and covers(m, row, col) then
      out[#out + 1] = ("**%s** · %s%s\n\n%s"):format(c.author, c.date:sub(1, 10), c.resolved and " · resolved" or "", c.text)
      for _, r in ipairs(c.replies) do
        out[#out + 1] = ("↳ **%s** · %s\n\n%s"):format(r.author, r.date:sub(1, 10), r.text)
      end
    end
  end
  for _, m in ipairs(api.nvim_buf_get_extmarks(buf, neo().ns.change, { row, 0 }, { row, -1 }, { details = true, overlap = true })) do
    local ch = doc.change[m[1]]
    if ch and covers(m, row, col) then
      local kind = ({ insert = "inserted", delete = "deleted", format = "formatting changed" })[ch.type]
      kind = ch.moved and (ch.type == "insert" and "moved here" or "moved away") or kind
      out[#out + 1] = ("*%s* by **%s** · %s"):format(kind, ch.author, (ch.date or ""):sub(1, 10))
    end
  end
  return #out > 0 and { contents = { kind = "markdown", value = table.concat(out, "\n\n---\n\n") } } or nil
end

-- Attributes in braces: a table's options line ("{width = full, height = tall}") and an image's
-- width ("{width = 12cm}"). An error if a table's option isn't one; a hint (with the text to write
-- instead: data.new) if they aren't written "name = value", separated by commas.
local function attribute_diagnostics(buf)
  local markup, tables, out = require("neoffice.markup"), require("neoffice.tables"), {}
  local lines = api.nvim_buf_get_lines(buf, 0, -1, false)
  local function add(r, c1, c2, severity, message, new)
    out[#out + 1] = {
      range = { start = { line = r - 1, character = c1 }, ["end"] = { line = r - 1, character = c2 } },
      severity = severity, message = message, source = "neoffice_ls", data = new and { new = new } or nil,
    }
  end
  -- the styles at the top: each line right, and a style the document has
  local header = markup.header(lines)
  local doc = neo().docs[buf]
  if header and header.page and header.page.bad then -- (the page's settings, and each section's)
    add(header.page.row, 0, #lines[header.page.row], vim.lsp.protocol.DiagnosticSeverity.Error, header.page.bad)
  end
  for r, line in ipairs(lines) do
    local sec = line:find("^\\newpage%s*{") and markup.section_line(line)
    if sec and sec.bad then
      add(r, 0, #line, vim.lsp.protocol.DiagnosticSeverity.Error, sec.bad)
    end
  end
  for _, st in ipairs(header and header.styles or {}) do
    local line = lines[st.row]
    if st.bad then
      add(st.row, 0, #line, vim.lsp.protocol.DiagnosticSeverity.Error, st.bad)
    elseif doc and doc.fmt.style_id and not (doc.header and doc.header[st.name:lower()]) and not doc.fmt.style_id(doc.ctx, st.name) then
      add(st.row, 0, #line, vim.lsp.protocol.DiagnosticSeverity.Error, ("not a paragraph style of this document: %s"):format(st.name))
    end
  end
  -- marks that close nothing: written as characters, or (^{, _{) formatting all that follows
  for _, block in ipairs(doc and neo().blocks(buf) or {}) do
    for _, o in ipairs(block.text ~= "" and markup.orphan_marks(block.text) or {}) do
      local r, c = neo().position_in(buf, block, o.col)
      add(r, c, c + #o.mark, vim.lsp.protocol.DiagnosticSeverity.Warning, o.spreads
        and ("this %s isn't closed: all that follows it is a %s (close it with })"):format(o.mark, o.flag == "sup" and "superscript" or "subscript")
        or ("this %s closes nothing: it is written as characters (write %s for them, or put back its other one to format)"):format(o.mark, (o.mark:gsub(".", "\\%0"))))
    end
  end
  for r, line in ipairs(lines) do
    if line:find("]{", 1, true) then -- a span of text's attributes ("[words]{font = Arial}")
      for _, c in ipairs(select(2, markup.parse(line))) do
        if c.span_bad then
          add(r, 0, #line, vim.lsp.protocol.DiagnosticSeverity.Error, c.span_bad)
          break
        end
      end
    end
    -- a format's line ("{align = justify, first = 0.5cm}"), or a paragraph's own attributes
    local above_table = (lines[r + 1] or ""):match("^|")
    local format = line:find("^[<%-{]") and not above_table and not (header and r <= header.last) and markup.para_line(line)
    local own = not format and line:find("^[<%-{]") and not (header and r <= header.last) and markup.block(line)
    local orphan = format and format.bad and line:sub(1, 1) == "{" and pcall(tables.parse, markup.table_options(line) or "")
    if orphan then
      add(r, 0, #line, vim.lsp.protocol.DiagnosticSeverity.Hint, "a table's options, with no table under it: nothing (delete the line)")
    elseif format and format.bad or own and own.para_bad then
      add(r, 0, #line, vim.lsp.protocol.DiagnosticSeverity.Error, format and format.bad or own.para_bad)
    elseif format and format.text then
      local tidy = markup.render_para_line(format)
      if tidy ~= line then
        add(r, 0, #line, vim.lsp.protocol.DiagnosticSeverity.Hint, "write it as " .. tidy, tidy)
      end
    end
    local fence, fence_attrs = markup.box_line(line) -- a box's line: its attributes (as an image's)
    if fence and fence_attrs then
      local attrs, bad = require("neoffice.floats").parse_box(fence_attrs)
      local tidy = attrs and "::: {" .. require("neoffice.floats").render(attrs) .. "}"
      if not attrs then
        add(r, 0, #line, vim.lsp.protocol.DiagnosticSeverity.Error, bad)
      elseif tidy ~= line and tidy ~= "::: {}" then
        add(r, 0, #line, vim.lsp.protocol.DiagnosticSeverity.Hint, "write it as " .. tidy, tidy)
      end
    end
    local inside = line:sub(1, 1) == "{" and markup.table_options(line)
    if inside and (lines[r + 1] or ""):match("^|") then
      local ok, err = pcall(tables.parse, inside)
      local _, tidy = markup.attributes(inside, tables.name_of)
      if not ok then
        add(r, 0, #line, vim.lsp.protocol.DiagnosticSeverity.Error, (tostring(err):gsub("^.-:%d+: ", "")))
      elseif "{" .. tidy .. "}" ~= line then
        add(r, 0, #line, vim.lsp.protocol.DiagnosticSeverity.Hint, "write it as {" .. tidy .. "}", "{" .. tidy .. "}")
      end
    end
    for s, inside, e in line:gmatch("==(){([^{}]*)}()") do -- a highlight's colour
      local name = inside:match("^%s*colou?r%s*=%s*([#%w]+)%s*$") or inside:match("^%s*([#%w]+)%s*$")
      local color = name and markup.color(name)
      if not color then
        add(r, s - 1, e - 1, vim.lsp.protocol.DiagnosticSeverity.Hint,
          "not a highlight colour: " .. inside .. " (they are " .. table.concat(vim.fn.sort(vim.tbl_keys(markup.COLORS)), ", ") .. ", or #rrggbb)")
      elseif color ~= inside then
        add(r, s - 1, e - 1, vim.lsp.protocol.DiagnosticSeverity.Hint, "write it as {" .. color .. "}", "{" .. color .. "}")
      end
    end
    local p = line:find("![", 1, true)
    while p do
      local _, _, after, _, attrs = markup.image_at(line, p)
      local tidy = attrs and attrs.attrs and "{" .. require("neoffice.floats").render(attrs.attrs) .. "}"
      if attrs and attrs.bad then -- (braces after an image that aren't its attributes)
        add(r, attrs.s - 1, attrs.e, vim.lsp.protocol.DiagnosticSeverity.Error, attrs.bad)
      elseif tidy and tidy ~= "{}" and line:sub(attrs.s, attrs.e) ~= tidy then
        add(r, attrs.s - 1, attrs.e, vim.lsp.protocol.DiagnosticSeverity.Hint, "write it as " .. tidy, tidy)
      end
      p = line:find("![", math.max(after or 0, p + 2), true)
    end
  end
  return out
end

-- Snippets (LSP snippet syntax: $1, ${1:default}, $0 where it ends), in neoffice's markup: those
-- of blocks at a line's start (the line holding only the word typed), those of text after a space.
M.SNIPPETS = {
  block = {
    { "head", "the page and the paragraph styles, at the top", "---\npage: {size = ${1:A4}, margins = ${2:2.5cm}}\nstyles:\n  ${3:Normal}: {font = ${4:Times New Roman}, size = ${5:12pt}}\n---\n$0" },
    { "table", "a table, its caption above it", "Table: ${1:Caption.}\n\n{rules = ${2:booktabs}}\n| ${3:Column} | ${4:Column} |\n|---|---|\n| $5 | $6 |\n$0" },
    { "figure", "an image, its caption under it", "![${1:what it shows}](${2:figures/file.png}){width = ${3:12cm}}\n\nFigure: ${4:Caption.}\n$0" },
    { "floatfigure", "an image and its caption in a box the text goes around", "::: {width = ${1:6cm}, float = ${2:right}}\n![](${3:figures/file.png})\n\nFigure: ${4:Caption.}\n:::\n\n$0" },
    { "box", "a box (text box) the text goes around", "::: {width = ${1:6cm}, float = ${2:right}}\n${3:Text in the box.}\n:::\n\n$0" },
    { "landscape", "landscape pages (a section), and portrait again after them", "\\\\newpage {orientation = landscape}\n\n${1:A wide table or figure.}\n\n\\\\newpage {orientation = portrait}\n\n$0" },
    { "section", "a section: a new page with its own settings", "\\\\newpage {${1:orientation = landscape}}\n\n$0" },
    { "newpage", "a page break", "\\\\newpage$0" },
    { "vspace", "vertical space between paragraphs: an empty paragraph of exactly that height", "\\\\vspace{${1:1cm}}$0" },
    { "toc", "a table of contents: its entries the headings", "\\\\tableofcontents\n\n$0" },
    { "equation", "an equation of its own (LaTeX)", "\\$\\$${1:E = mc^2}\\$\\$\n$0" },
    { "numbered", "a numbered list, with an item under another (the document counts)", "#. ${1:First}\n#. ${2:Second}\n  #. ${3:Under it}\n$0" },
    { "bullets", "a bulleted list", "- ${1:One}\n- ${2:Two}\n$0" },
    { "quote", "a quotation", "> ${1:Quoted text.}\n$0" },
    { "title", "the title", "% ${1:Title}\n$0" },
    { "format", "the format of the paragraphs after it", "{align = ${1:justify}, first = ${2:0.5cm}}+\n$0" },
  },
  inline = {
    { "footnote", "a footnote (right after the word before it)", "^[${1:Note.}]$0", attach = true },
    { "link", "a link", "[${1:text}](${2:https://})$0" },
    { "span", "words in another font or size", "[${1:words}]{font = ${2:Arial}, size = ${3:10pt}}$0" },
    { "highlight", "highlighted words", "==${1:words}==$0" },
    { "math", "an equation in the text (LaTeX)", "\\$${1:x^2}\\$$0" },
  },
}

-- Completion: inside a table's options line ("{" above a table), its options ("rules = grid") or,
-- after "name =", that option's values; an image's width after "){width ="; a highlight's colour
-- after "=={".
local KIND = vim.lsp.protocol.CompletionItemKind
local function completion(buf, pos)
  local markup, tables = require("neoffice.markup"), require("neoffice.tables")
  local lines = api.nvim_buf_get_lines(buf, pos.line, pos.line + 2, false)
  local line = lines[1] or ""
  local before = line:sub(1, pos.character)
  local items = {}
  local function add(label, start, kind, doc, insert)
    items[#items + 1] = {
      label = label, kind = kind, documentation = doc, filterText = label,
      textEdit = { newText = insert or label, range = { start = { line = pos.line, character = start }, ["end"] = pos } },
    }
  end
  local color = before:match("==%{([#%w]*)$")
  if color then
    local start = #before - #color
    for name, hex in pairs(markup.COLORS) do
      add(name, start, KIND.Color, "#" .. hex, name .. "}")
    end
    return items
  end
  -- a paragraph format's attributes ("<=> {..."): those not given yet, or a value
  -- a paragraph format's attributes ("<=> {..."), or a span of text's ("[words]{..."): those not
  -- given yet, or a value (a font: one of those installed)
  -- in the styles at the top: a style's name (those the document has), or its attributes
  local all = api.nvim_buf_get_lines(buf, 0, 200, false)
  local header = markup.header(all)
  local in_header = header and pos.line + 1 > 2 and pos.line + 1 < header.last
  local doc = neo().docs[buf]
  if in_header and before:match("^%s+[^:{}]*$") and doc and doc.fmt.paragraph_styles then
    local typed = before:match("^%s+(.*)$")
    local listed = {}
    for _, st in ipairs(header.styles) do
      listed[(st.name or ""):lower()] = true
    end
    for _, name in ipairs(doc.fmt.paragraph_styles(doc.ctx)) do
      if not listed[name:lower()] then
        add(name, #before - #typed, KIND.Class, "a paragraph style of this document", name .. ": {")
      end
    end
    return items
  end
  -- the page's settings ("page: {..." at the top) or a section's ("\\newpage {..."): its attributes,
  -- those not given yet, or a value
  local page_braces = header and pos.line + 1 < header.last and before:match("^page:%s*{([^}]*)$") or before:match("^\\newpage%s*{([^}]*)$")
  if page_braces then
    local pages = require("neoffice.pages")
    local used = {}
    for _, a in ipairs((markup.attributes(page_braces))) do
      used[a.name or ""] = true
    end
    local item = page_braces:match("([^,;]*)$")
    local name, value = item:match("^%s*([%a_%-]+)%s*=%s*(.*)$")
    local VALUES = {
      size = pages.VALUES.size, orientation = pages.VALUES.orientation, start = pages.VALUES.start,
      margins = { "2.5cm", "2cm", "2.54cm", "2.5cm 3cm" },
      ["margin-top"] = { "2.5cm", "2cm" }, ["margin-bottom"] = { "2.5cm", "2cm" }, ["margin-left"] = { "2.5cm", "3cm" }, ["margin-right"] = { "2.5cm", "3cm" },
    }
    if name and VALUES[name] then
      for _, v in ipairs(VALUES[name]) do
        add(v, #before - #value, KIND.EnumMember, pages.DESCRIPTIONS[name])
      end
    elseif not name then
      local typed = item:match("^%s*(.-)$")
      for _, key in ipairs({ "size", "orientation", "margins", "margin-top", "margin-bottom", "margin-left", "margin-right", "start" }) do
        if not used[key] and not (key == "start" and before:match("^page:")) then
          for _, v in ipairs(VALUES[key]) do
            add(key .. " = " .. v, #before - #typed, KIND.Property, pages.DESCRIPTIONS[key])
          end
        end
      end
    end
    return items
  end
  local hbraces = in_header and before:match("^%s+[^:{}]-%s*:%s*{([^}]*)$")
  local mark, pbraces = before:match("^(%S%S%S)%s*{([^}]*)$") -- (after any of the marks: <=> <|> ->> <<-)
  local marks = {}
  for _, m in pairs(markup.ALIGN) do
    marks[vim.trim(m)] = true
  end
  pbraces = not in_header and marks[mark or ""] and pbraces or nil
  if not pbraces and not in_header and not (lines[2] or ""):match("^|") then -- (a format's line, or a paragraph's own: "{...")
    pbraces = before:match("^{([^}]*)$")
  end
  pbraces = pbraces or hbraces
  -- (a span's attributes, "[words]{...", or a table cell's, "| {..." / "| <|> {...": its text's font and size)
  local sbraces = not pbraces and (before:match("%]{([^}]*)$") or before:match("^|.*|%s*[<>|=%-]*%s*{([^}]*)$") or before:match("^|%s*[<>|=%-]*%s*{([^}]*)$"))
  if pbraces or sbraces then
    local pf = require("neoffice.paras")
    local used = {}
    for _, a in ipairs((markup.attributes(pbraces or sbraces))) do
      used[a.name or ""] = true
    end
    local item = (pbraces or sbraces):match("([^,;]*)$")
    local name, value = item:match("^%s*([%a_%-]+)%s*=%s*(.*)$")
    local VALUES = {
      ["indent-left"] = { "0.5cm", "1cm", "1.27cm", "none" }, ["indent-right"] = { "0.5cm", "1cm", "none" },
      first = { "0.5cm", "1.27cm", "-0.5cm", "none" }, before = { "0pt", "6pt", "12pt", "none" },
      after = { "0pt", "6pt", "12pt", "none" }, line = { "single", "1.15", "1.5", "double", "none" },
      size = { "9pt", "10pt", "11pt", "12pt", "14pt", "16pt" },
      color = { "black", "red", "blue", "darkblue", "darkgreen", "darkred", "darkgray", "#1f4e79", "none" },
      align = { "left", "center", "right", "justify", "none" },
    }
    local keys = hbraces and pf.STYLE_KEYS or pbraces and pf.KEYS or markup.SPAN_KEYS
    if hbraces then -- (a style says what it gives: no none; "#..." in quotes, as YAML would read a comment)
      for key, list in pairs(VALUES) do
        VALUES[key] = vim.tbl_map(function(v) return v:find("^#") and '"' .. v .. '"' or v end, vim.tbl_filter(function(v) return v ~= "none" end, list))
      end
      VALUES.bold, VALUES.italic = { "" }, { "" }
    end
    if name == "font" then
      -- every word typed anywhere in its name ("new rom": Times New Roman); the document's fonts
      -- first (those its lines name, and its own), then those installed, those starting so first
      local typed = vim.trim(value:gsub('"', "")):lower()
      local words = vim.split(typed, "%s+", { trimempty = true })
      local function matches(family)
        local low = family:lower()
        for _, w in ipairs(words) do
          if not low:find(w, 1, true) then
            return false
          end
        end
        return true
      end
      local installed, seen, ours = {}, {}, {}
      for _, family in ipairs(M.fonts()) do
        installed[family:lower()] = true
      end
      local function mine(family)
        family = vim.trim(family or "")
        if family ~= "" and family ~= "none" and not seen[family:lower()] then
          seen[family:lower()], ours[#ours + 1] = true, family
        end
      end
      local doc = neo().docs[buf]
      if doc and doc.fmt.font_default then
        mine(doc.fmt.font_default(doc.ctx).font)
      end
      if doc and doc.fmt.run_font then -- (those its text has: most used first; as read or last written)
        if not doc.text_fonts or doc.text_fonts.paras ~= doc.paras then
          local count, known = {}, {}
          for _, props in pairs(doc.run) do
            if known[props] == nil then
              known[props] = doc.fmt.run_font(doc.ctx, props).font or false
            end
            local f = known[props]
            if f then
              count[f] = (count[f] or 0) + 1
            end
          end
          local used = vim.tbl_keys(count)
          table.sort(used, function(x, y) return count[x] > count[y] or count[x] == count[y] and x < y end)
          doc.text_fonts = { paras = doc.paras, list = used }
        end
        for _, f in ipairs(doc.text_fonts.list) do
          mine(f)
        end
      end
      for r, l in ipairs(api.nvim_buf_get_lines(buf, 0, -1, false)) do
        if r ~= pos.line + 1 and l:find("font", 1, true) then
          for v in l:gmatch('font%s*=%s*"([^"]+)"') do
            mine(v)
          end
          for v in l:gmatch("font%s*=%s*([^,;}\"]+)") do
            mine(v)
          end
        end
      end
      local function offer(family, sort, what)
        add(markup.attr_value(family), #before - #value, KIND.EnumMember, what)
        items[#items].sortText = sort .. family:lower()
      end
      for i, family in ipairs(ours) do -- (in their order: its default, then the most used)
        if matches(family) then
          offer(family, ("0%04d"):format(i), installed[family:lower()] and "in this document"
            or "in this document (not installed here: shown in a similar font; Word and LibreOffice elsewhere may have it)")
        end
      end
      for _, family in ipairs(M.fonts()) do
        if not seen[family:lower()] and matches(family) then
          offer(family, family:lower():sub(1, #typed) == typed and "1" or "2", "installed")
        end
      end
    elseif name == "style" and pbraces and not hbraces then -- (the document's paragraph styles)
      local doc = require("neoffice").docs[buf]
      local typed = vim.trim(value):lower()
      local names = doc and doc.fmt.paragraph_styles and doc.fmt.paragraph_styles(doc.ctx) or {}
      table.insert(names, "none")
      for _, v in ipairs(names) do
        if v:lower():find(typed, 1, true) then
          add(v, #before - #value, KIND.EnumMember, v == "none" and "the style most plain paragraphs have" or "a paragraph style of this document")
        end
      end
    elseif name and VALUES[name] then
      for _, v in ipairs(VALUES[name]) do
        add(v, #before - #value, KIND.EnumMember, pf.DESCRIBE[name])
      end
    elseif not name then
      local typed = item:match("^%s*(.-)$")
      for _, key in ipairs(keys) do
        if not used[key] then
          if (key == "bold" or key == "italic") and hbraces then -- (said alone)
            add(key, #before - #typed, KIND.Property, pf.DESCRIBE[key])
          end
          for _, v in ipairs((key == "font" or key == "style") and { "" } or (key ~= "bold" and key ~= "italic") and VALUES[key] or {}) do
            if v ~= "none" then
              add(key .. " = " .. v, #before - #typed, KIND.Property, pf.DESCRIBE[key])
            end
          end
        end
      end
    end
    return items
  end
  -- an image's attributes ("){...") or a box's ("::: {..."): those not given yet, or a value
  local braces = before:match("%)%{([^}]*)$") or before:match("^:::%s*{([^}]*)$")
  if braces then
    local floats = require("neoffice.floats")
    local used = {}
    for _, a in ipairs((markup.attributes(braces))) do
      used[a.name or floats.name_of(a.value) or ""] = true
    end
    local item = braces:match("([^,;]*)$")
    local name, value = item:match("^%s*([%a_]+)%s*=%s*(%S*)$")
    local box = before:match("^:::") ~= nil
    local function values(key)
      return floats.VALUES[key] or key == "width" and { "100%", "50%", "6cm", "8cm", "12cm" }
        or key:match("^crop%-") and { "5%", "10%", "25%" } or { "0.3cm", "0.5cm", "1cm" }
    end
    local function describe(key, v)
      return floats.DESCRIBE[key .. "." .. v] or key:match("^crop%-") and floats.DESCRIBE.crop or nil
    end
    if name then
      for _, v in ipairs(values(name:lower())) do
        add(v, #before - #value, KIND.EnumMember, describe(name:lower(), v))
      end
    else
      local typed = item:match("^%s*(.-)$")
      for _, key in ipairs(floats.KEYS) do
        if not used[key] and key ~= "x" and key ~= "y" and not (box and key:match("^crop%-")) then
          for _, v in ipairs(values(key)) do
            add(key .. " = " .. v, #before - #typed, KIND.Property, describe(key, v))
          end
        end
      end
    end
    return items
  end
  local inside = line:sub(1, 1) == "{" and not before:find("}", 1, true) and (lines[2] or ""):match("^|") and before:sub(2)
  if not inside then -- snippets
    local word = before:match("^(%a*)$")
    local list = word and M.SNIPPETS.block or {}
    if not word then
      word = before:match("^(%a+)$") or before:match("%s(%a+)$")
      list = word and M.SNIPPETS.inline or {}
    end
    for _, sn in ipairs(word and vim.list_extend(vim.list_slice(list), before:match("^%a*$") and M.SNIPPETS.inline or {}) or {}) do
      if sn[1]:sub(1, #word) == word:lower() or word == "" then
        items[#items + 1] = {
          label = sn[1], kind = KIND.Snippet, detail = sn[2], filterText = sn[1], insertTextFormat = 2,
          documentation = { kind = "markdown", value = "```\n" .. sn[3]:gsub("%${%d+:([^}]*)}", "%1"):gsub("%$%d", ""):gsub("\\(.)", "%1") .. "\n```" },
          textEdit = { newText = sn[3], range = { start = { line = pos.line, character = #before - #word - (sn.attach and before:sub(-#word - 1, -#word - 1) == " " and 1 or 0) }, ["end"] = pos } },
        }
      end
    end
    return items
  end
  local used = {}
  for _, a in ipairs((markup.attributes(inside))) do
    used[a.name or tables.name_of(a.value) or ""] = true
  end
  local item = inside:match("([^,;]*)$") -- (the one being written)
  local name, value = item:match("^%s*([%a_]+)%s*=%s*(%S*)$")
  local DESCRIBE = tables.DESCRIBE
  if name then -- its values
    local key = name:lower()
    key = key == "borders" and "rules" or key == "align" and "position" or key
    for _, v in ipairs(tables.WORDS[key] or {}) do
      add(v, #before - #value, KIND.EnumMember, DESCRIBE[key .. "." .. v])
    end
    if key == "width" then
      for _, v in ipairs({ "80%", "60%", "12cm" }) do
        add(v, #before - #value, KIND.Value, DESCRIBE["width.length"])
      end
    end
  else -- the options not given yet, each with its values
    local typed = item:match("^%s*(.-)$")
    for _, key in ipairs(tables.ORDER) do
      if not used[key] or typed ~= "" and tables.name_of(typed) == key then
        for _, v in ipairs(tables.WORDS[key]) do
          add(key .. " = " .. v, #before - #typed, KIND.Property, DESCRIBE[key .. "." .. v])
        end
      end
    end
  end
  return items
end

-- A floating image followed by its caption: both in a box, so the caption goes where the image
-- goes (the image's place becomes the box's).
local function box_figure(buf, params, out)
  local markup, floats = require("neoffice.markup"), require("neoffice.floats")
  local lines = api.nvim_buf_get_lines(buf, 0, -1, false)
  for r = params.range.start.line + 1, math.min(params.range["end"].line + 1, #lines) do
    local line = lines[r]
    local p = line:find("![", 1, true)
    local attrs
    if p then
      attrs = select(5, markup.image_at(line, p))
    end
    local caption = r + 2 <= #lines and lines[r + 1] == "" and markup.block(lines[r + 2]).caption == "figure"
    if attrs and attrs.attrs and floats.floating(attrs.attrs) and caption and not markup.box_line(lines[r - 1] or "") then
      local box, image = {}, { width = attrs.attrs.width }
      for key, value in pairs(attrs.attrs) do
        box[key] = value
      end
      local inner = floats.render(image)
      local last = r + 2 -- (the caption's last line)
      while lines[last + 1] and lines[last + 1] ~= "" and not markup.box_line(lines[last + 1]) do
        last = last + 1
      end
      out[#out + 1] = {
        title = "Box this figure with its caption (::: {...})", kind = "refactor.rewrite",
        edit = { changes = { [params.textDocument.uri] = {
          { range = { start = { line = r - 1, character = 0 }, ["end"] = { line = r - 1, character = 0 } }, newText = "::: {" .. floats.render(box) .. "}\n" },
          { range = { start = { line = r - 1, character = attrs.s }, ["end"] = { line = r - 1, character = attrs.e - 1 } }, newText = inner },
          { range = { start = { line = last - 1, character = #lines[last] }, ["end"] = { line = last - 1, character = #lines[last] } }, newText = "\n:::" },
        } } },
      }
    end
  end
end

-- On a format's line: its attributes moved into the style of the paragraphs under it (the line at
-- the top says so; the format's line says what is left of them, or goes). Nothing looks any
-- different: paragraphs of that style elsewhere that had its values say them now. Offered only
-- when that holds: the paragraphs under it all of one style (known: as the document has them,
-- saved), listed at the top; none of that style that would change and can't say so (in a cell,
-- a box).
local function into_style(buf, params, out)
  local markup, pf = require("neoffice.markup"), require("neoffice.paras")
  local doc = neo().docs[buf]
  local lines = api.nvim_buf_get_lines(buf, 0, -1, false)
  local header = markup.header(lines)
  local r = params.range.start.line + 1
  local line = lines[r] or ""
  if not (doc and doc.header and header and doc.fmt.style_name) or r <= header.last or (lines[r + 1] or ""):match("^|") then
    return
  end
  local f = line:find("^[<%-{]") and markup.para_line(line)
  if not f or f.bad then
    return
  end
  local moved = {}
  for k, v in pairs(f.attrs) do
    moved[k] = v ~= "none" and v or nil
  end
  moved.align = f.align ~= "none" and f.align or nil
  if not next(moved) then
    return
  end
  local blocks, styles = neo().block_styles(buf)
  local before = { attrs = {} } -- (the format before it: what the paragraphs under it would go back to)
  for _, b in ipairs(blocks) do
    if b.first < r and (b.run or b.format) then
      before = b.run or b.format -- (its run's: a paragraph's own line is for it only)
    end
  end
  local run, style -- (its paragraphs' format, and their style: a "+" line's run, else the one right below it)
  local key = f.sticky and "run" or "format"
  for k, b in ipairs(blocks) do
    if b.first > r and b[key] and not run then
      run = b[key]
    end
    local _, _, kind_at = markup.block(b.text)
    if run and rawequal(b[key], run) and b.text ~= "" and not b.cell and not b.box and not kind_at then
      if not styles[k] or (style and styles[k] ~= style) then
        return
      end
      style = styles[k]
    end
  end
  if not style then
    return
  end
  local name = doc.fmt.style_name(doc.ctx, style)
  local entry = vim.iter(header.styles):find(function(st) return st.name and st.name:lower() == name:lower() and st.attrs end)
  if not entry then
    return
  end
  -- (its line: what it says itself, now with these; what it gives as it comes out, to say elsewhere)
  local old, new = doc.fmt.style_view(doc.ctx, style), vim.deepcopy(entry.attrs)
  for k, v in pairs(moved) do
    new[k] = v
  end
  local DEFAULT = { align = "left", line = "single" }
  local edits = {}
  local function replace(row, c1, c2, text)
    edits[#edits + 1] = { range = { start = { line = row - 1, character = c1 }, ["end"] = { line = row - 1, character = c2 } }, newText = text }
  end
  -- the others of that style that had its values: said now
  for k, b in ipairs(blocks) do
    local own, start, kind_at, _, brace_at = markup.block(b.text)
    local plain_para = not kind_at and not own.caption
    if styles[k] == style and b.text ~= "" and not rawequal(b[key], run) and plain_para then
      local eff = pf.apply(b.format or { attrs = {} }, own.align, own.para)
      local adds = {}
      for key in pairs(moved) do
        local have = key == "align" and eff.align or eff.attrs[key]
        if have == nil or have == "none" then
          adds[key] = old[key] or DEFAULT[key] or (vim.tbl_contains(pf.LAYOUT, key) and "0cm") or nil
        end
      end
      if next(adds) then
        if b.cell or b.box then
          return -- (it can't say so)
        end
        local row = b.first + (b.lead and 1 or 0)
        local text = lines[row]
        local mine = vim.tbl_extend("force", {}, own.para or {}, adds)
        local align = own.align or mine.align
        mine.align = nil
        local braces = "{" .. pf.render(vim.tbl_extend("force", {}, mine, { align = not own.align and align or nil })) .. "} "
        if brace_at then -- (its braces: with these too)
          local close = text:find("} ", brace_at, true)
          replace(row, brace_at - 1, close + 1, braces)
        else
          replace(row, start - 1, start - 1, braces)
        end
      end
    elseif styles[k] == style and b.text ~= "" and not rawequal(b[key], run) and (b.cell or b.box) then
      return
    end
  end
  -- the style's line, and the format's
  local hline = lines[entry.row]
  local brace = hline:find("{", 1, true)
  replace(entry.row, brace - 1, #hline, "{" .. pf.render_style(new) .. "}")
  -- (what it says of what moved: the style's now, "none" if the format before it would go on)
  local inherited = function(key)
    return not f.reset and (key == "align" and before.align or before.attrs[key]) ~= nil
  end
  local left = { reset = f.reset, attrs = {}, sticky = f.sticky }
  for k, v in pairs(f.attrs) do
    if not moved[k] then
      left.attrs[k] = v
    elseif inherited(k) then
      left.attrs[k] = "none"
    end
  end
  if not moved.align then
    left.align = f.align
  elseif inherited("align") then
    left.align = "none"
  end
  if not left.reset and not left.align and not next(left.attrs) then
    edits[#edits + 1] = { range = { start = { line = r - 1, character = 0 }, ["end"] = { line = r, character = 0 } }, newText = "" }
  else
    replace(r, 0, #line, markup.render_para_line(left))
  end
  out[#out + 1] = {
    title = ("Move this format into the style %s (nothing looks different)"):format(name), kind = "refactor.rewrite",
    edit = { changes = { [params.textDocument.uri] = edits } },
  }
end

local function code_actions(buf, params)
  local uri, out = params.textDocument.uri, {}
  box_figure(buf, params, out)
  into_style(buf, params, out)
  for _, d in ipairs(attribute_diagnostics(buf)) do
    if d.data and d.range.start.line >= params.range.start.line and d.range.start.line <= params.range["end"].line then
      out[#out + 1] = {
        title = d.message:gsub("^write", "Write"), kind = "quickfix", diagnostics = { d }, isPreferred = true,
        edit = { changes = { [uri] = { { range = d.range, newText = d.data.new } } } },
      }
    end
  end
  return out
end

-- The server: requests are answered at once, from the buffer.
function M.server(dispatchers)
  local closing, id, pending = false, 0, {}
  local handlers = {
    initialize = function()
      return {
        capabilities = {
          positionEncoding = "utf-8", -- (columns are bytes, as in Neovim)
          -- (it reads the buffer: changes only say when to look at it again, for diagnostics)
          textDocumentSync = { openClose = true, change = 2 },
          codeActionProvider = true,
          documentFormattingProvider = true,
          documentRangeFormattingProvider = true,
          completionProvider = { triggerCharacters = { "{", "=", ",", " " } },
          documentSymbolProvider = true,
          foldingRangeProvider = true,
          hoverProvider = true,
        },
        serverInfo = { name = "neoffice_ls" },
      }
    end,
    ["textDocument/documentSymbol"] = function(params) return symbols(vim.uri_to_bufnr(params.textDocument.uri)) end,
    ["textDocument/foldingRange"] = function(params) return folds(vim.uri_to_bufnr(params.textDocument.uri)) end,
    ["textDocument/hover"] = function(params) return hover(vim.uri_to_bufnr(params.textDocument.uri), params.position) end,
    ["textDocument/codeAction"] = function(params) return code_actions(vim.uri_to_bufnr(params.textDocument.uri), params) end,
    ["textDocument/formatting"] = function(params)
      return require("neoffice.format").edits(vim.uri_to_bufnr(params.textDocument.uri))
    end,
    ["textDocument/rangeFormatting"] = function(params)
      return require("neoffice.format").edits(vim.uri_to_bufnr(params.textDocument.uri), params.range.start.line, params.range["end"].line)
    end,
    ["textDocument/completion"] = function(params)
      return { isIncomplete = false, items = completion(vim.uri_to_bufnr(params.textDocument.uri), params.position) }
    end,
    shutdown = function() return vim.NIL end,
  }
  return {
    request = function(method, params, callback)
      id = id + 1
      local handler = handlers[method]
      local ok, result = pcall(handler or error, params)
      vim.schedule(function()
        if handler and ok then
          callback(nil, result)
        else
          callback({ code = -32601, message = handler and tostring(result) or ("not supported: " .. method) }, nil)
        end
      end)
      return true, id
    end,
    notify = function(method, params)
      if method == "exit" then
        dispatchers.on_exit(0, 15)
      elseif method == "textDocument/didOpen" or method == "textDocument/didChange" then
        local uri = params.textDocument.uri
        -- (once typing pauses: each change puts it off again, so nothing is checked as you type)
        local timer = pending[uri] or vim.uv.new_timer()
        pending[uri] = timer
        timer:stop()
        timer:start(method == "textDocument/didOpen" and 0 or 300, 0, vim.schedule_wrap(function()
          if pending[uri] == timer then
            pending[uri] = nil
            timer:close()
          end
          local buf = vim.uri_to_bufnr(uri)
          if not closing and api.nvim_buf_is_loaded(buf) then
            local ok, diagnostics = pcall(attribute_diagnostics, buf)
            dispatchers.notification("textDocument/publishDiagnostics", { uri = uri, diagnostics = ok and diagnostics or {} })
          end
        end))
      end
      return true
    end,
    is_closing = function() return closing end,
    terminate = function() closing = true end,
  }
end

function M.start(buf)
  M.prefetch_fonts()
  if neo().config.lsp == false then
    return
  end
  vim.lsp.start({ name = M.NAME, cmd = M.server, root_dir = vim.fs.dirname(api.nvim_buf_get_name(buf)) }, { bufnr = buf })
end

return M
