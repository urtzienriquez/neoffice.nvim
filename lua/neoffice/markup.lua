-- One paragraph <-> one line of text with Markdown-style marks:
--   **bold**  *italic*  ++underline++  ~~strikethrough~~  ==highlight== (yellow; another colour:
--   ==highlight=={green}, see M.COLORS)
--   superscript^{2}  subscript_{food} (as in LaTeX)
--   [text](url)  ![alt](image path){width = 12cm}  $inline equation$  $$display equation$$  ^[a footnote]
--   ⟦a field⟧ (a citation, date, reference... : shown, not edited)
--   \ at the end of a line: a line break    \newpage: a page break
-- and, starting a line, what kind of paragraph it is (its "block"):
--   "# " heading (## level 2, ...)   "% " title   "> " quote   "- " bullet item   "1. " numbered item
--   "Table: " a table's caption   "Figure: " a figure's caption (numbered by the document: see
--   M.captions; the words are the option `captions`)
--   (list items indented by two spaces per level)
--   then its alignment, if not left:  "<|> centered"   "->> right"   "<=> justified"
-- Tables are Markdown pipe tables: "| cell | cell |", one line per row (a line of dashes after the
-- first); paragraphs in one cell are separated by "<br>", and "|" in text is written "\|". The line
-- of dashes says its columns' alignment ("|:---:|---:|": centered, right), and a line above it the
-- options of its look ("{width = full, height = tall}": see M.table_options and tables.lua).
-- Escapes: \* \$ \# ...
-- A span is { text = "...", b, i, u, s, sup, sub = bool, h = false|colour, link = url|nil,
-- raw = bool (no escaping) },
-- or one of: { image = { alt, path, width = "12cm" | nil } }, { math = { tex, display } }, { footnote = { text } },
-- { field = { text } }, { line_break = true }, { page_break = true }.
local M = {}

-- The formatting a span can have. To support another, add it here (and to the reader/writer of
-- each format).
-- h (highlight) is a colour: false, or one of M.COLORS (Word's highlight colours) or "#rrggbb".
M.FLAGS = { "b", "i", "u", "s", "h", "sup", "sub" }
local MARK = {
  b = { "**", "**" },
  i = { "*", "*" },
  u = { "++", "++" },
  s = { "~~", "~~" },
  h = { "==", "==" },
  sup = { "^{", "}" },
  sub = { "_{", "}" },
}

-- Highlight colours, as written after a highlight ("==text=={green}"; yellow, the default, isn't
-- written): Word's (its names, in lower case) with their RGB values (for .odt).
M.COLORS = {
  yellow = "ffff00", green = "00ff00", cyan = "00ffff", magenta = "ff00ff", blue = "0000ff", red = "ff0000",
  darkblue = "000080", darkcyan = "008080", darkgreen = "008000", darkmagenta = "800080", darkred = "800000",
  darkyellow = "808000", darkgray = "808080", lightgray = "c0c0c0", black = "000000", white = "ffffff",
}
-- A colour as written ("Green", "#A1B2C3"...) -> as stored ("green", "#a1b2c3"), or nil if it isn't one.
function M.color(s)
  s = (s or ""):lower()
  return (M.COLORS[s] or s:match("^#%x%x%x%x%x%x$")) and s or nil
end
-- The highlight colour of a hex value: its name if it has one.
function M.color_of_rgb(hex)
  hex = (hex or ""):lower():gsub("^#", "")
  for name, value in pairs(M.COLORS) do
    if value == hex then
      return name
    end
  end
  return hex:match("^%x%x%x%x%x%x$") and "#" .. hex or nil
end
-- The named colour nearest to a highlight colour (for .docx, whose highlights are named).
function M.named_color(c)
  if M.COLORS[c] then
    return c
  end
  local function rgb(hex) return tonumber(hex:sub(1, 2), 16), tonumber(hex:sub(3, 4), 16), tonumber(hex:sub(5, 6), 16) end
  local r, g, b = rgb(c:sub(2))
  local best, distance = "yellow", math.huge
  for name, value in pairs(M.COLORS) do
    local r2, g2, b2 = rgb(value)
    local d = (r - r2) ^ 2 + (g - g2) ^ 2 + (b - b2) ^ 2
    if d < distance or d == distance and name < best then
      best, distance = name, d
    end
  end
  return best
end
local function closing(k, value) -- (a highlight's colour after it)
  return MARK[k][2] .. (k == "h" and type(value) == "string" and value ~= "yellow" and "{" .. value .. "}" or "")
end

-- (braced: inside ^{...} or _{...}, where "}" closes; strict: every character that could start a
-- mark, for text next to marks: see M.render)
local function escape(s, braced, cell, strict)
  if strict then -- (a "\" too; otherwise one is written as it is, as it escapes nothing)
    s = s:gsub("[\\~=+{^_!]", "\\%0")
  end
  s = s:gsub("[*%[%]$]", "\\%0"):gsub("([%^_]){", "\\%1{")
  s = s:gsub("⟦", "\\⟦"):gsub("⟧", "\\⟧"):gsub("¶", "\\¶")
  if cell then -- (in a table cell, "|" and "<br>" separate)
    s = s:gsub("|", "\\|"):gsub("<br>", "\\<br>")
  end
  return braced and (s:gsub("}", "\\}")) or s
end

M.escape = escape

-- What starts a line of each kind of paragraph. block = { heading = level } |
-- { title = true } | { quote = true } | { list = { kind = "bullet" | "number", level = 1.., n = its
-- number } } | {} (a plain one); with align = "center" | "right" | "justify" (or nil: left).
local ALIGN = { center = "<|> ", right = "->> ", justify = "<=> ", left = "<<- " }
M.ALIGN = ALIGN

-- The words that start captions ("Table: ", "Figure: "): { table = "Table", figure = "Figure" }
-- (the option `captions`: each a word, or a list of them, the first written and the others also
-- read, as "Fig."), and the kind of caption a word is.
local function words(value, default)
  return type(value) == "table" and value or { value or default }
end
function M.captions()
  local config = require("neoffice").config.captions or {}
  return { table = words(config.table, "Table")[1], figure = words(config.figure, "Figure")[1] }
end
local function caption_kind(word)
  word = word:lower()
  local config = require("neoffice").config.captions or {}
  for kind, default in pairs({ table = "Table", figure = "Figure" }) do
    for _, label in ipairs(words(config[kind], default)) do
      if label:lower() == word then
        return kind
      end
    end
  end
end
M.caption_kind = caption_kind

function M.prefix(block)
  local list, kind = block.list, ""
  if block.caption then
    kind = M.captions()[block.caption] .. ": "
  elseif block.heading then
    kind = ("#"):rep(block.heading) .. " "
  elseif block.title then
    kind = "% "
  elseif block.quote then
    kind = "> "
  elseif list then
    kind = ("  "):rep(list.level - 1) .. (list.kind == "bullet" and "- " or (list.n or 1) .. ". ")
  end
  -- its own format's attributes, in braces after its mark (see paras.lua)
  local attrs = block.para and require("neoffice.paras").render(block.para) or ""
  if attrs ~= "" then
    return kind .. (ALIGN[block.align] or ALIGN.left) .. "{" .. attrs .. "} "
  end
  return kind .. (ALIGN[block.align] or "")
end

-- A line with only a mark and its attributes ("<=> {first = 0.5cm}"), which sets the format of
-- the paragraphs after it (see paras.lua): { align (nil: left), attrs }, and why if they aren't
-- right (bad); nil if the line isn't one.
function M.para_line(line)
  for align, mark in pairs(ALIGN) do
    local marker = vim.trim(mark)
    if line:sub(1, #marker) == marker then
      local rest = line:sub(#marker + 1)
      if rest:match("^%s*$") then
        return { align = align ~= "left" and align or nil, attrs = {} }
      end
      local inside = rest:match("^%s*{([^{}]*)}%s*$")
      if inside then
        local attrs, bad = require("neoffice.paras").parse(inside)
        return { align = align ~= "left" and align or nil, attrs = attrs or {}, bad = bad, text = inside }
      end
    end
  end
end

-- The block a line starts with; where its text begins; where its kind's and its alignment's
-- marker characters are (to escape them in a paragraph that isn't of that kind). Escaped markers
-- (\#, \- ...) don't count.
function M.block(line)
  local block, start, kind_at = {}, 1, nil
  local hashes = line:match("^(#+) ")
  local indent, after = line:match("^( *)%- ()")
  local number, dot, after_dot = line:match("^( *)%d+()%. ()")
  local word, after_word = line:match("^([^%s:\\]+): ()")
  local caption = word and caption_kind(word)
  if M.box_line(line) then -- (a box's line: text that reads as one is escaped, "\:::")
    block, start, kind_at = { fence = true }, #line + 1, 1
  elseif caption then
    block, start, kind_at = { caption = caption }, after_word, 1
  elseif hashes then
    block, start, kind_at = { heading = #hashes }, #hashes + 2, 1
  elseif line:match("^%% ") then
    block, start, kind_at = { title = true }, 3, 1
  elseif line:match("^> ") then
    block, start, kind_at = { quote = true }, 3, 1
  elseif line:match("^|") then -- (a table row: see M.cells)
    return { row = true }, 1, 1, nil
  elseif indent then
    block, start, kind_at = { list = { kind = "bullet", level = math.floor(#indent / 2) + 1 } }, after, #indent + 1
  elseif number then
    block, start, kind_at = { list = { kind = "number", level = math.floor(#number / 2) + 1 } }, after_dot, dot
  end
  local align_at
  for _, align in ipairs({ "center", "justify", "right", "left" }) do
    local marker = ALIGN[align]
    if line:sub(start, start + #marker - 1) == marker then
      block.align, align_at, start = align, start, start + #marker
      break
    end
  end
  -- (after its mark, its own format's attributes: "<=> {first = 0}")
  local inside, after
  if align_at and not kind_at then -- (plain paragraphs only)
    inside, after = line:match("^{([^{}]*)} ()", start)
  end
  if inside then
    local attrs, bad = require("neoffice.paras").parse(inside)
    if attrs then
      block.para, start = attrs, after
    else
      block.para_bad = bad
      start = after
    end
  end
  return block, start, kind_at, align_at
end

local function kind(block)
  return block.heading or block.title or block.quote or block.list or block.row or block.caption or block.fence
end

local math_at, enclosed, link_at -- (below)

-- Where the image, equation, footnote, field or link at `p` ends (its "|" don't separate cells).
local function skip(line, p)
  local c = line:sub(p, p)
  local after
  if c == "$" then
    after = select(3, math_at(line, p))
  elseif line:sub(p, p + 1) == "^[" then
    after = select(2, enclosed(line, p + 2, "]"))
  elseif line:sub(p, p + 2) == "⟦" then
    after = select(2, enclosed(line, p + 3, "⟧"))
  elseif c == "!" then
    after = select(3, M.image_at(line, p))
  elseif c == "[" then
    after = link_at(line, p)
  end
  return after
end

-- A table's line of dashes: { align = { [column] = "center" | "right" | "left" | nil } }, or nil
-- if the line isn't one. "|:---:|---:|:---|---|": centered, right, left, and left too (not said).
function M.table_format(line)
  local dashes = line
  if not dashes:match("^|[%s:|%-]*%-[%s:|%-]*$") then
    return nil
  end
  local format = { align = {} }
  local col = 0
  for cell in dashes:sub(2):gmatch("([^|]*)|?") do
    cell = vim.trim(cell)
    if cell ~= "" then
      col = col + 1
      local left, right = cell:sub(1, 1) == ":", cell:sub(-1) == ":"
      format.align[col] = left and right and "center" or right and "right" or left and "left" or nil
    end
  end
  return format
end

-- The line of dashes of a table of `columns` columns with this format (see M.table_format); with
-- `widths` (of the columns' text: an aligned table), as wide as each column.
function M.format_line(format, columns, widths)
  local cells = {}
  for col = 1, columns do
    local align = format and format.align[col]
    local left, right = (align == "left" or align == "center") and 1 or 0, (align == "right" or align == "center") and 1 or 0
    -- (aligned: as wide as the cells' text and the spaces around it; else three dashes)
    local n = widths and (widths[col] or 3) + 2 or 3 + left + right
    cells[col] = (":"):rep(left) .. ("-"):rep(n - left - right) .. (":"):rep(right)
  end
  return "|" .. table.concat(cells, "|") .. "|"
end

-- A line of dashes with these columns aligned so ({ [column] = "right" | ... }), as wide as it was.
function M.realign(line, aligns)
  local col = 0
  return (line:gsub("([^|]+)", function(seg)
    col = col + 1
    local align = aligns[col]
    if not align then
      return seg
    end
    local n = math.max(#vim.trim(seg), 3)
    local left, right = (align == "left" or align == "center") and 1 or 0, (align == "right" or align == "center") and 1 or 0
    return (":"):rep(left) .. ("-"):rep(n - left - right) .. (":"):rep(right)
  end))
end

-- The cells of a table row line, as written: for each, { s, e } (1-based bytes of its text without
-- the spaces around it; e = s - 1 for an empty one).
function M.cell_spans(line)
  local out = {}
  for _, paras in ipairs(M.cells(line) or {}) do
    local first, last = paras[1], paras[#paras]
    out[#out + 1] = { s = first.col + 1, e = last.col + #last.text }
  end
  return out
end

-- A box's line ("::: {width = 6cm, float = right}" opens a box, ":::" closes it: see floats.lua):
-- true, and the text in its braces (or nil), if the line is one.
function M.box_line(line)
  if line:sub(1, 3) ~= ":::" then
    return nil
  end
  local rest = line:match("^:::%s*(.-)%s*$")
  if rest == "" then
    return true, nil
  end
  local attrs = rest:match("^{([^{}]*)}$")
  if attrs then
    return true, attrs
  end
end

-- Whether a line holds only floating images (after its paragraph's marks, if any): such a line
-- is joined to the next with nothing between, as the floating images of the paragraph below it.
function M.floats_only(line)
  if not line:find("float", 1, true) then -- (quick: most lines)
    return false
  end
  local _, pos = M.block(line)
  local floats = require("neoffice.floats")
  local any = false
  while pos <= #line do
    if line:sub(pos, pos + 1) ~= "![" then
      return false
    end
    local _, _, after, _, attrs = M.image_at(line, pos)
    if not (after and attrs and attrs.attrs and floats.floating(attrs.attrs)) then
      return false
    end
    pos, any = after, true
  end
  return any
end

-- The options on the line above a table ("{width = full, height = tall}"): the text inside the
-- braces, or nil if the line isn't one.
function M.table_options(line)
  return line:match("^{([^{}]*)}%s*$")
end

-- Attributes, as written in braces: "name = value" (or "name=value"), separated by commas; a value
-- alone stands for itself ("full": its name is the one it belongs to). Returns a list of
-- { name (or nil), value, s, e (bytes of `text`, 1-based) }, and the way to write them: "name =
-- value, name = value" (`names(value)` gives a lone value's name).
function M.attributes(text, names)
  local list, p = {}, 1
  while p <= #text do
    local s, name, value, e = text:match("^%s*()([%a_][%w_%-]*)%s*=%s*([^%s,;=]+)()", p)
    if not s then
      s, value, e = text:match("^%s*()([^%s,;=]+)()", p)
    end
    if s then
      list[#list + 1] = { name = name and name:lower(), value = value, s = s, e = e - 1 }
      p = e
    else
      p = p + 1 -- (a separator, or an "=" alone)
    end
  end
  local tidy = {}
  for _, a in ipairs(list) do
    local name = a.name or names and names(a.value)
    tidy[#tidy + 1] = name and name .. " = " .. a.value or a.value
  end
  return list, table.concat(tidy, ", ")
end

-- The cells of a table row line: for each, { col = where its text starts (0-based), text }, with
-- the paragraphs of the cell, split at "<br>". nil for the line of dashes (`dashes`: this line may
-- be it, being the one under a table's first row).
function M.cells(line, dashes)
  if dashes and M.table_format(line) then
    return nil
  end
  local cells, start, p = {}, 2, 2
  local function cell(stop) -- (one cell: line:sub(start, stop - 1))
    local paras, from, q = {}, start, start
    local function para(to)
      local text = line:sub(from, to - 1)
      local lead = #text:match("^%s*")
      paras[#paras + 1] = { col = from - 1 + lead, text = text:sub(lead + 1):gsub("%s+$", "") }
    end
    while q < stop do
      local after = skip(line, q)
      if after then
        q = after
      elseif line:sub(q, q) == "\\" then
        q = q + 2
      elseif line:sub(q, q + 3) == "<br>" then
        para(q)
        q = q + 4
        from = q
      else
        q = q + 1
      end
    end
    para(stop)
    cells[#cells + 1] = paras
  end
  while p <= #line do
    local c = line:sub(p, p)
    local after = skip(line, p)
    if after then
      p = after
    elseif c == "\\" then
      p = p + 2
    elseif c == "|" and not (line:sub(p - 1, p - 1) == "<" and line:sub(p + 1, p + 1) == ">") then -- (not "<|>")
      cell(p)
      start = p + 1
      p = p + 1
    else
      p = p + 1
    end
  end
  if line:sub(start):match("%S") then -- (no "|" at the end)
    cell(#line + 1)
  end
  return cells
end

-- Whether a line could read back otherwise than it was written: only text with what could start
-- a mark that isn't always escaped ("==", "~~", "++", "\\", "{", "!") can. (Most lines can't: they
-- aren't read again.)
local function could_misread(line)
  return line:find("==", 1, true) or line:find("~~", 1, true) or line:find("++", 1, true)
    or line:find("[\\{!]") or line:find("[%^_]{")
end

-- Returns the line and, for every span, the { start, end } byte columns of its text.
-- A line without the backslashes before brackets that start nothing: a "[" not in a link's text,
-- not after ^ or ! and with no "](" before the next "[", and a "]" not in a link's text. (And the
-- columns of cols, moved as the backslashes go.)
local function unbracket(line, cols)
  local out, removed, inlink, i = {}, {}, false, 1
  while i <= #line do
    local c, nx = line:sub(i, i), line:sub(i + 1, i + 1)
    if c == "\\" and (nx == "[" or nx == "]") then
      local rest = line:sub(i + 2)
      local upto = rest:find("[", 1, true)
      local drop = not inlink and (nx == "]" or not line:sub(i - 1, i - 1):match("[%^!]")
        and not (upto and rest:sub(1, upto - 1) or rest):find("](", 1, true) and not (upto and rest:sub(1, upto - 1) or rest):find("\\](", 1, true))
      if drop then
        removed[#removed + 1] = i - 1
      else
        out[#out + 1] = c
      end
      out[#out + 1] = nx
      i = i + 2
    elseif c == "\\" then
      out[#out + 1] = line:sub(i, i + 1)
      i = i + 2
    else
      if c == "[" then
        inlink = true
      elseif c == "]" and nx == "(" then
        inlink = false
      end
      out[#out + 1] = c
      i = i + 1
    end
  end
  if #removed == 0 then
    return line, cols
  end
  local moved = {}
  for k, col in ipairs(cols or {}) do
    local pair = {}
    for j = 1, 2 do
      local n = 0
      for _, r in ipairs(removed) do
        n = n + (r < col[j] and 1 or 0)
      end
      pair[j] = col[j] - n
    end
    moved[k] = vim.tbl_extend("force", col, pair)
  end
  return table.concat(out), moved
end

function M.render(spans, block, strict)
  local parts, cols, len, cur = {}, {}, 0, {}
  local function put(s)
    parts[#parts + 1] = s
    len = len + #s
  end
  local function switch(to)
    if cur.link and cur.link ~= to.link then
      put("](" .. cur.link:gsub("%)", "%%29") .. ")")
    end
    for k = #M.FLAGS, 1, -1 do -- (closing in the reverse order of opening, so marks nest)
      local flag = M.FLAGS[k]
      if cur[flag] and cur[flag] ~= to[flag] then -- (off, or a highlight of another colour)
        put(closing(flag, cur[flag]))
      end
    end
    for _, k in ipairs(M.FLAGS) do
      if to[k] and to[k] ~= cur[k] then
        put(MARK[k][1])
      end
    end
    if to.link and to.link ~= cur.link then
      if (parts[#parts] or ""):sub(-1) == "!" then -- so "![" doesn't read as an image
        parts[#parts], len = parts[#parts]:sub(1, -2) .. "\\!", len + 1
      end
      put("[")
    end
    cur = to
  end

  block = block or {}
  put(M.prefix(block))
  for k, s in ipairs(spans) do
    local token -- (images, equations, footnotes, fields and breaks: written as they are)
    if s.image then
      local attrs = require("neoffice.floats").render(s.image)
      token = "![" .. escape(s.image.alt, false, block.cell) .. "](" .. s.image.path:gsub("%)", "%%29") .. ")"
        .. (attrs ~= "" and "{" .. attrs .. "}" or "")
    elseif s.math then -- (on one line)
      local tex = s.math.tex:gsub("%s*\n%s*", " ")
      token = (s.math.display and "$$" or "$") .. tex .. (s.math.display and "$$" or "$")
    elseif s.footnote then
      token = "^[" .. escape(s.footnote.text:gsub("\n", " "), false, block.cell) .. "]"
    elseif s.field then
      token = "⟦" .. escape(s.field.text:gsub("\n", " "), false, block.cell) .. "⟧"
    elseif s.line_break then
      token = "\\\n"
    elseif s.page_break then
      token = "\\newpage"
    end
    if token then
      s.raw, s.text = true, token
    end
    if s.text and s.text ~= "" then
      switch(token and {} or s)
      local text = s.raw and s.text or escape(s.text, s.sup or s.sub, block.cell, strict)
      cols[k] = { len, len + #text }
      put(text)
    end
    cols[k] = cols[k] or { len, len }
  end
  switch({})
  local line = table.concat(parts)
  if line:sub(-2) == "\\\n" then -- (a line break ending the paragraph: a "\" ending its last line)
    line = line:sub(1, -2)
  end
  -- a paragraph whose text would read as another kind or alignment (or as a blank line) gets a
  -- "\"; so does each line after a line break that would read as the start of a paragraph
  local start = 1
  while start do
    for _ = 1, 3 do
      local text = line:sub(start, (line:find("\n", start, true) or #line + 1) - 1)
      local parsed, _, kind_at, align_at = M.block(text)
      local at = text:match("^%s+$") and 1
      if not at and (start > 1 or not kind(block)) and kind(parsed) and not (block.cell and parsed.row) then
        at = kind_at
      elseif not at and (start > 1 or not block.align) and parsed.align then
        at = align_at
      elseif not at and start == 1 and block.align and not block.para and (parsed.para or parsed.para_bad) then
        at = align_at + #ALIGN[parsed.align] -- (text in braces after its mark: "\{")
      end
      if not at then
        break
      end
      at = start + at - 1
      line = line:sub(1, at - 1) .. "\\" .. line:sub(at)
      for _, c in ipairs(cols) do
        c[1], c[2] = c[1] + (c[1] >= at - 1 and 1 or 0), c[2] + (c[2] >= at - 1 and 1 or 0)
      end
    end
    start = line:find("\n", start, true)
    start = start and start + 1
  end
  -- brackets that start no link, image or footnote are written as they are ("[B]", "[1, 2]": no
  -- "\[", which reads as LaTeX's display math), if the line still reads back the same
  if not strict and line:find("\\[%[%]]") then
    local plain, pcols = unbracket(line, cols)
    if M.same(plain, spans) then
      line, cols = plain, pcols
    end
  end
  -- what is written must read back as the same text: if not (marks next to text that looks like
  -- marks), it is written again escaping every character that could start a mark
  -- (only text with characters that could be read as marks can read back otherwise)
  if not strict and could_misread(line) and not M.same(line, spans) then
    return M.render(spans, block, true)
  end
  return line, cols
end

-- Whether a line reads back as these spans: the same characters, formatting and objects.
local function signature(entry, token)
  if token then
    return "\0" .. token
  end
  local flags = {}
  for _, k in ipairs(M.FLAGS) do
    flags[#flags + 1] = entry[k] and k .. (entry[k] == true and k == "h" and "yellow" or tostring(entry[k])) or ""
  end
  return table.concat(flags, ",") .. "|" .. (entry.link or "") .. "|" .. entry.ch
end
function M.same(line, spans)
  local want = {}
  for _, s in ipairs(spans) do
    local token = s.image and "image" or s.math and "math" or s.footnote and "footnote" or s.field and "field"
      or s.page_break and "page_break"
    if token then
      want[#want + 1] = signature(nil, token)
    elseif s.line_break then
      want[#want + 1] = signature({ ch = "\n" })
    elseif s.text then
      local fmt = s.raw and {} or s
      for ch in s.text:gmatch("[^\128-\191][\128-\191]*") do
        want[#want + 1] = signature({ ch = ch, link = fmt.link, b = fmt.b, i = fmt.i, u = fmt.u, s = fmt.s, h = fmt.h, sup = fmt.sup, sub = fmt.sub })
      end
    end
  end
  local _, chars = M.parse(line)
  if #chars ~= #want then
    return false
  end
  for k, c in ipairs(chars) do
    local token = c.image and "image" or c.math and "math" or c.footnote and "footnote" or c.field and "field"
      or c.page_break and "page_break"
    if signature(c, token) ~= want[k] then
      return false
    end
  end
  return true
end

-- An image "![alt](path)" or "![alt](path){width = 12cm, float = right}" at `pos`: alt, path,
-- position after it, width (or nil), and its attributes ({ s, e, text: inside the braces, attrs
-- = { name = value } (see floats.lua) }, or nil). Braces after it that aren't right have `bad`
-- (why) instead of attrs: writing the document then fails, saying it.
function M.image_at(line, pos)
  local alt, path, after = line:match("^!%[(.-)%]%(([^)]*)%)()", pos)
  if alt then
    alt, path = alt:gsub("\\(.)", "%1"), path:gsub("%%29", ")")
    local s, inside, after_attrs = line:match("^(){([^{}]*)}()", after)
    local attrs, bad
    if inside then
      attrs, bad = require("neoffice.floats").parse(inside)
    end
    if attrs then
      return alt, path, after_attrs, attrs.width, { s = s, e = after_attrs - 1, text = inside, attrs = attrs }
    end
    if inside then
      return alt, path, after_attrs, nil, { s = s, e = after_attrs - 1, text = inside, bad = bad }
    end
    return alt, path, after
  end
end

-- Image widths: "12cm", "120mm", "4.5in", or "50%" of `text_width` (in cm) <-> centimetres.
local CM = { cm = 1, mm = 0.1, ["in"] = 2.54, pt = 2.54 / 72 }
function M.cm(width, text_width)
  local n, unit = (width or ""):match("^%s*([%d.]+)%s*(%a+)%s*$")
  if not n then
    n = (width or ""):match("^%s*([%d.]+)%s*%%%s*$")
    return tonumber(n) and text_width and tonumber(n) / 100 * text_width or nil
  end
  return tonumber(n) and CM[unit] and tonumber(n) * CM[unit] or nil
end
function M.width(cm) -- (as shown: to a millimetre)
  return (("%.1f"):format(cm):gsub("%.0$", "")) .. "cm"
end

-- An equation "$tex$" or "$$tex$$" at `pos`: tex, display, position after it. As in pandoc, an
-- inline equation's "$" must touch its content, and the closing one can't be followed by a digit
-- (so "$5 and $10" stays text).
function math_at(line, pos)
  if line:sub(pos, pos + 1) == "$$" then
    local stop = line:find("$$", pos + 2, true)
    return stop and stop > pos + 2 and line:sub(pos + 2, stop - 1), true, stop and stop + 2
  end
  if line:sub(pos + 1, pos + 1):match("%s") then
    return nil
  end
  local stop = pos
  repeat
    stop = line:find("$", stop + 1, true)
  until not stop or (stop > pos + 1 and not line:sub(stop - 1, stop - 1):match("[%s\\]") and not line:sub(stop + 1, stop + 1):match("%d"))
  if stop then
    return line:sub(pos + 1, stop - 1), false, stop + 1
  end
end

-- Text up to an unescaped `close` from `pos` (after an opening mark): its content (unescaped)
-- and the position after `close`.
function enclosed(line, pos, close)
  local p = pos
  while p <= #line do
    if line:sub(p, p) == "\\" then
      p = p + 2
    elseif line:sub(p, p + #close - 1) == close then
      return (line:sub(pos, p - 1):gsub("\\(.)", "%1")), p + #close
    else
      p = p + 1
    end
  end
end

-- A link "[text](url)" at `pos`: the position after it.
function link_at(line, pos)
  local p = pos + 1
  while p <= #line do
    local c = line:sub(p, p)
    if c == "\\" then
      p = p + 2
    elseif c == "]" then
      local after = line:match("^%([^)]*%)()", p + 1)
      return after
    else
      p = p + 1
    end
  end
end

-- What a "\\" escapes (the characters escape() and render() escape); before anything else, it is
-- just a backslash. (A "\\" ending the text is a line break; "\\ " starting a line, a space.)
local ESCAPED = {}
for c in ("\\*[]$<>^_~=+|#%-.!{}:\n" .. "⟦⟧¶"):gmatch("[^\128-\191][\128-\191]*") do
  ESCAPED[c] = true
end

-- The first bytes of everything that isn't plain text (marks, escapes, images, equations...).
local MARKS = {}
for c in ("\\!$^_*~=+}[" .. "\226"):gmatch(".") do -- (\226: the first byte of ⟦)
  MARKS[c:byte()] = true
end

-- Returns the block (see M.prefix: list numbers aren't read) and a list of characters:
-- { col = byte column in the line (0-based), ch = "x", b, i, u, s, sup, sub = bool, h = false|colour,
-- link = url|nil }
-- A line break is the character "\n" (from "\" at the end of a line). An image, equation,
-- footnote, field or page break is one entry { col, ch = "", len, image | math | footnote | field |
-- page_break = ... }.
function M.parse(line)
  local block, pos = M.block(line)
  local text_start = pos

  local chars, cur, link_close, braces = {}, {}, nil, {} -- braces: the open ^{ and _{
  local h_from -- (where the open highlight started, in chars)
  local function emit(at)
    local byte = line:byte(at) -- (the character's bytes, from its first: UTF-8)
    local ch = line:sub(at, at + (byte < 0xC0 and 0 or byte < 0xE0 and 1 or byte < 0xF0 and 2 or 3))
    chars[#chars + 1] = { -- (the flags of M.FLAGS)
      col = at - 1, ch = ch, link = cur.link, b = cur.b == true, i = cur.i == true, u = cur.u == true,
      s = cur.s == true, h = cur.h or false, sup = cur.sup == true, sub = cur.sub == true,
    }
    return at + #ch
  end
  local function token(key, value, after)
    chars[#chars + 1] = { col = pos - 1, ch = "", len = after - pos, [key] = value }
    return after
  end
  local function later(mark, from) -- (an unescaped `mark` after `from`: a mark only opens if it closes)
    local p = from
    while p <= #line do
      if line:sub(p, p) == "\\" then
        p = p + 2
      elseif line:sub(p, p + #mark - 1) == mark then
        return true
      else
        p = p + 1
      end
    end
    return false
  end
  -- "[" starts a link only if a matching "](url)" follows
  local function find_link(at)
    local p = at + 1
    while p <= #line do
      local c = line:sub(p, p)
      if c == "\\" then
        p = p + 2
      elseif c == "]" then
        local url = line:match("^%(([^)]*)%)", p + 1)
        return url and p, url and url:gsub("%%29", ")")
      else
        p = p + 1
      end
    end
  end

  while pos <= #line do
    if not MARKS[line:byte(pos)] and pos ~= link_close then -- (most characters: just text)
      pos = emit(pos)
      goto continue
    end
    local c = line:sub(pos, pos)
    local two = line:sub(pos, pos + 1)
    local note, after_note
    if two == "^[" then
      note, after_note = enclosed(line, pos + 2, "]")
    end
    local field, after_field
    if line:sub(pos, pos + 2) == "⟦" then
      field, after_field = enclosed(line, pos + 3, "⟧")
    end
    if pos == link_close then
      cur.link, link_close, pos = nil, nil, line:find(")", pos, true) + 1
    elseif line:sub(pos, pos + 7) == "\\newpage" then
      pos = token("page_break", true, pos + 8)
    elseif c == "\\" and pos == #line and pos > text_start then -- (a line break ending the text)
      pos = emit(pos)
      chars[#chars].ch = "\n"
    elseif c == "\\" and (ESCAPED[line:match("^.[\128-\191]*", pos + 1) or ""]
      or line:sub(pos + 1, pos + 1):match("[ \t]") and (pos == text_start or line:sub(pos - 1, pos - 1) == "\n")
      or pos == text_start and M.block(line:sub(pos + 1)).caption) then -- (\Table: is text)
      pos = emit(pos + 1)
      chars[#chars].esc = true -- (escaped: a "¶" is just a ¶)
    elseif c == "!" and M.image_at(line, pos) then
      local alt, path, after, _, attrs = M.image_at(line, pos)
      pos = token("image", vim.tbl_extend("force", attrs and attrs.attrs or {}, { alt = alt, path = path, bad = attrs and attrs.bad }), after)
    elseif c == "$" and (math_at(line, pos)) then
      local tex, display, after = math_at(line, pos)
      pos = token("math", { tex = tex, display = display }, after)
    elseif note then
      pos = token("footnote", { text = note }, after_note)
    elseif field then
      pos = token("field", { text = field }, after_field)
    elseif two == "**" and (cur.b or later("**", pos + 2)) then
      cur.b, pos = not cur.b, pos + 2
    elseif c == "*" and two ~= "**" and (cur.i or later("*", pos + 1)) then
      cur.i, pos = not cur.i, pos + 1
    elseif two == "~~" and (cur.s or later("~~", pos + 2)) then
      cur.s, pos = not cur.s, pos + 2
    elseif two == "==" and cur.h then -- (its colour, if not yellow, after it: "=={green}")
      local name, after = line:match("^{%s*([#%w]+)%s*}()", pos + 2)
      if not name then -- ("=={color = green}" too)
        name, after = line:match("^{%s*colou?r%s*=%s*([#%w]+)%s*}()", pos + 2)
      end
      local color = M.color(name)
      if color and color ~= "yellow" then
        for k = h_from, #chars do
          chars[k].h = chars[k].h and color
        end
      end
      cur.h, pos = false, color and after or pos + 2
    elseif two == "==" and later("==", pos + 2) then
      cur.h, h_from, pos = "yellow", #chars + 1, pos + 2
    elseif two == "++" and (cur.u or later("++", pos + 2)) then
      cur.u, pos = not cur.u, pos + 2
    elseif two == "^{" or two == "_{" then
      local flag = two == "^{" and "sup" or "sub"
      table.insert(braces, flag)
      cur[flag], pos = true, pos + 2
    elseif c == "}" and #braces > 0 then
      cur[table.remove(braces)], pos = false, pos + 1
    elseif c == "[" and not cur.link and find_link(pos) then
      link_close, cur.link = find_link(pos)
      pos = pos + 1
    else
      pos = emit(pos)
    end
    ::continue::
  end
  return block, chars
end

return M
