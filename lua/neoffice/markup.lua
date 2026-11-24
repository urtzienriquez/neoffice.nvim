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

-- (braced: inside ^{...} or _{...}, where "}" closes; strict: every character that could start a
-- mark, for text next to marks: see M.render)
local function escape(s, braced, cell, strict)
  if strict then -- (a "\" too; otherwise one is written as it is, as it escapes nothing)
    s = s:gsub("[\\~=<^_!]", "\\%0")
  end
  s = s:gsub("[*%[%]$]", "\\%0"):gsub("<(/?u>)", "\\<%1"):gsub("([%^_]){", "\\%1{")
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

-- The cells of a table row line: for each, { col = where its text starts (0-based), text }, with
-- the paragraphs of the cell, split at "<br>". nil for the line of dashes (`dashes`: this line may
-- be it, being the one under a table's first row).
function M.cells(line, dashes)
  if dashes and line:match("^|[%s:|%-]*%-[%s:|%-]*$") then
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

-- Returns the line and, for every span, the { start, end } byte columns of its text.
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
      token = "![" .. escape(s.image.alt, false, block.cell) .. "](" .. s.image.path:gsub("%)", "%%29") .. ")"
        .. (s.image.width and "{width=" .. s.image.width .. "}" or "")
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
  -- what is written must read back as the same text: if not (marks next to text that looks like
  -- marks), it is written again escaping every character that could start a mark
  -- (only text with characters that could be read as marks can read back otherwise)
  if not strict and line:find("[=~_^<\\!]") and not M.same(line, spans) then
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
    flags[#flags + 1] = entry[k] and k or ""
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
for c in ("\\*[]$<>^_~=|#%-.!{}\n" .. "⟦⟧¶"):gmatch("[^\128-\191][\128-\191]*") do
  ESCAPED[c] = true
end

-- The first bytes of everything that isn't plain text (marks, escapes, images, equations...).
local MARKS = {}
for c in ("\\!$^_*~=<}[" .. "\226"):gmatch(".") do -- (\226: the first byte of ⟦)
  MARKS[c:byte()] = true
end

-- Returns the block (see M.prefix: list numbers aren't read) and a list of characters:
-- { col = byte column in the line (0-based), ch = "x", b, i, u, s, h, sup, sub = bool, link = url|nil }
-- A line break is the character "\n" (from "\" at the end of a line). An image, equation,
-- footnote, field or page break is one entry { col, ch = "", len, image | math | footnote | field |
-- page_break = ... }.
function M.parse(line)
  local block, pos = M.block(line)
  local text_start = pos

  local chars, cur, link_close, braces = {}, {}, nil, {} -- braces: the open ^{ and _{
  local function emit(at)
    local ch = line:match("^.[\128-\191]*", at)
    chars[#chars + 1] = { -- (the flags of M.FLAGS)
      col = at - 1, ch = ch, link = cur.link, b = cur.b == true, i = cur.i == true, u = cur.u == true,
      s = cur.s == true, h = cur.h == true, sup = cur.sup == true, sub = cur.sub == true,
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
      or line:sub(pos + 1, pos + 1) == " " and (pos == text_start or line:sub(pos - 1, pos - 1) == "\n")) then
      pos = emit(pos + 1)
      chars[#chars].esc = true -- (escaped: a "¶" is just a ¶)
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
    elseif two == "**" and (cur.b or later("**", pos + 2)) then
      cur.b, pos = not cur.b, pos + 2
    elseif c == "*" and two ~= "**" and (cur.i or later("*", pos + 1)) then
      cur.i, pos = not cur.i, pos + 1
    elseif two == "~~" and (cur.s or later("~~", pos + 2)) then
      cur.s, pos = not cur.s, pos + 2
    elseif two == "==" and (cur.h or later("==", pos + 2)) then
      cur.h, pos = not cur.h, pos + 2
    elseif line:find("^<u>", pos) and not cur.u and later("</u>", pos + 3) then
      cur.u, pos = true, pos + 3
    elseif line:find("^</u>", pos) and cur.u then
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
    ::continue::
  end
  return block, chars
end

return M
