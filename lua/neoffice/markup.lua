-- One paragraph <-> one line of text with Markdown-style marks:
--   **bold**  *italic*  <u>underline</u>  ~~strikethrough~~  ==highlight==
--   superscript^{2}  subscript_{food} (as in LaTeX)
--   [text](url)  ![alt](image path){width=12cm}  $inline equation$  $$display equation$$  ^[a footnote]
--   ⟦a field⟧ (a citation, date, reference... : shown, not edited)
--   \ at the end of a line: a line break    \newpage: a page break
-- and, starting a line, what kind of paragraph it is (its "block"):
--   "# " heading (## level 2, ...)   "% " title   "> " quote   "- " bullet item   "1. " numbered item
--   (list items indented by two spaces per level)
--   then its alignment, if not left:  "<|> centered"   "->> right"   "<=> justified"
-- Tables are Markdown pipe tables: "| cell | cell |", one line per row (a line of dashes after the
-- first); paragraphs in one cell are separated by "<br>", and "|" in text is written "\|".
-- Escapes: \* \$ \# ...
-- A span is { text = "...", b, i, u, s, h, sup, sub = bool, link = url|nil, raw = bool (no escaping) },
-- or one of: { image = { alt, path, width = "12cm" | nil } }, { math = { tex, display } }, { footnote = { text } },
-- { field = { text } }, { line_break = true }, { page_break = true }.
local M = {}

-- The formatting a span can have. To support another, add it here (and to the reader/writer of
-- each format).
M.FLAGS = { "b", "i", "u", "s", "h", "sup", "sub" }
local MARK = {
  b = { "**", "**" },
  i = { "*", "*" },
  u = { "<u>", "</u>" },
  s = { "~~", "~~" },
  h = { "==", "==" },
  sup = { "^{", "}" },
  sub = { "_{", "}" },
}

local function escape(s, braced, cell) -- (braced: inside ^{...} or _{...}, where "}" closes)
  s = s:gsub("[\\*%[%]$]", "\\%0"):gsub("<(/?u>)", "\\<%1"):gsub("([%^_]){", "\\%1{")
  s = s:gsub("~~+", function(r) return (r:gsub("~", "\\~")) end):gsub("==+", function(r) return (r:gsub("=", "\\=")) end)
  s = s:gsub("⟦", "\\⟦"):gsub("⟧", "\\⟧")
  if cell then -- (in a table cell, "|" and "<br>" separate)
    s = s:gsub("|", "\\|"):gsub("<br>", "\\<br>")
  end
  return braced and (s:gsub("}", "\\}")) or s
end

M.escape = escape

-- What starts a line of each kind of paragraph. block = { heading = level } |
-- { title = true } | { quote = true } | { list = { kind = "bullet" | "number", level = 1.., n = its
-- number } } | {} (a plain one); with align = "center" | "right" | "justify" (or nil: left).
local ALIGN = { center = "<|> ", right = "->> ", justify = "<=> " }

function M.prefix(block)
  local list, kind = block.list, ""
  if block.heading then
    kind = ("#"):rep(block.heading) .. " "
  elseif block.title then
    kind = "% "
  elseif block.quote then
    kind = "> "
  elseif list then
    kind = ("  "):rep(list.level - 1) .. (list.kind == "bullet" and "- " or (list.n or 1) .. ". ")
  end
  return kind .. (ALIGN[block.align] or "")
end

-- The block a line starts with; where its text begins; where its kind's and its alignment's
-- marker characters are (to escape them in a paragraph that isn't of that kind). Escaped markers
-- (\#, \- ...) don't count.
function M.block(line)
  local block, start, kind_at = {}, 1, nil
  local hashes = line:match("^(#+) ")
  local indent, after = line:match("^( *)%- ()")
  local number, dot, after_dot = line:match("^( *)%d+()%. ()")
  if hashes then
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
  for _, align in ipairs({ "center", "justify", "right" }) do
    local marker = ALIGN[align]
    if line:sub(start, start + #marker - 1) == marker then
      block.align, align_at, start = align, start, start + #marker
      break
    end
  end
  return block, start, kind_at, align_at
end

local function kind(block)
  return block.heading or block.title or block.quote or block.list or block.row
end

-- The cells of a table row line: for each, { col = where its text starts (0-based), text }, with
-- the paragraphs of the cell, split at "<br>". nil for the line of dashes under the first row.
function M.cells(line)
  if line:match("^|[%s:|%-]*%-[%s:|%-]*$") then
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
      if line:sub(q, q) == "\\" then
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
    if c == "\\" then
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

-- Returns the line and, for every span, the { start, end } byte columns of its text.
function M.render(spans, block)
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
      if cur[M.FLAGS[k]] and not to[M.FLAGS[k]] then
        put(MARK[M.FLAGS[k]][2])
      end
    end
    for _, k in ipairs(M.FLAGS) do
      if to[k] and not cur[k] then
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
      token = "![" .. escape(s.image.alt) .. "](" .. s.image.path:gsub("%)", "%%29") .. ")"
        .. (s.image.width and "{width=" .. s.image.width .. "}" or "")
    elseif s.math then
      token = (s.math.display and "$$" or "$") .. s.math.tex .. (s.math.display and "$$" or "$")
    elseif s.footnote then
      token = "^[" .. escape(s.footnote.text) .. "]"
    elseif s.field then
      token = "⟦" .. escape(s.field.text) .. "⟧"
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
      local text = s.raw and s.text or escape(s.text, s.sup or s.sub, block.cell)
      cols[k] = { len, len + #text }
      put(text)
    end
    cols[k] = cols[k] or { len, len }
  end
  switch({})
  local line = table.concat(parts)
  -- a paragraph whose text would read as another kind or alignment (or as a blank line) gets a "\"
  for _ = 1, 3 do
    local parsed, _, kind_at, align_at = M.block(line)
    local at = line:match("^%s+$") and 1
    if not at and not kind(block) and kind(parsed) and not (block.cell and parsed.row) then
      at = kind_at
    elseif not at and not block.align and parsed.align then
      at = align_at
    end
    if not at then
      break
    end
    line = line:sub(1, at - 1) .. "\\" .. line:sub(at)
    for _, c in ipairs(cols) do
      c[1], c[2] = c[1] + (c[1] >= at - 1 and 1 or 0), c[2] + (c[2] >= at - 1 and 1 or 0)
    end
  end
  return line, cols
end

-- An image "![alt](path)" or "![alt](path){width=12cm}" at `pos`: alt, path, position after it,
-- width (or nil).
function M.image_at(line, pos)
  local alt, path, after = line:match("^!%[(.-)%]%(([^)]*)%)()", pos)
  if alt then
    local width, after_width = line:match("^{width=([^}]*)}()", after)
    return (alt:gsub("\\(.)", "%1")), (path:gsub("%%29", ")")), after_width or after, width
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
local function math_at(line, pos)
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
local function enclosed(line, pos, close)
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

-- Returns the block (see M.prefix: list numbers aren't read) and a list of characters:
-- { col = byte column in the line (0-based), ch = "x", b, i, u, s, h, sup, sub = bool, link = url|nil }
-- A line break is the character "\n" (from "\" at the end of a line). An image, equation,
-- footnote, field or page break is one entry { col, ch = "", len, image | math | footnote | field |
-- page_break = ... }.
function M.parse(line)
  local block, pos = M.block(line)

  local chars, cur, link_close, braces = {}, {}, nil, {} -- braces: the open ^{ and _{
  local function emit(at)
    local ch = line:match("^.[\128-\191]*", at)
    local entry = { col = at - 1, ch = ch, link = cur.link }
    for _, k in ipairs(M.FLAGS) do
      entry[k] = not not cur[k]
    end
    chars[#chars + 1] = entry
    return at + #ch
  end
  local function token(key, value, after)
    chars[#chars + 1] = { col = pos - 1, ch = "", len = after - pos, [key] = value }
    return after
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
    elseif c == "\\" and pos < #line then
      pos = emit(pos + 1)
    elseif c == "!" and M.image_at(line, pos) then
      local alt, path, after, width = M.image_at(line, pos)
      pos = token("image", { alt = alt, path = path, width = width }, after)
    elseif c == "$" and (math_at(line, pos)) then
      local tex, display, after = math_at(line, pos)
      pos = token("math", { tex = tex, display = display }, after)
    elseif note then
      pos = token("footnote", { text = note }, after_note)
    elseif field then
      pos = token("field", { text = field }, after_field)
    elseif two == "**" then
      cur.b, pos = not cur.b, pos + 2
    elseif c == "*" then
      cur.i, pos = not cur.i, pos + 1
    elseif two == "~~" then
      cur.s, pos = not cur.s, pos + 2
    elseif two == "==" then
      cur.h, pos = not cur.h, pos + 2
    elseif line:find("^<u>", pos) then
      cur.u, pos = true, pos + 3
    elseif line:find("^</u>", pos) then
      cur.u, pos = false, pos + 4
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
  end
  return block, chars
end

return M
