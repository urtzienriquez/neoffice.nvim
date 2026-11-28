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
-- A text's colour as shown: its name if it has one ("red"), else "#rrggbb"; nil if it isn't one.
-- From a name, "#rrggbb", or six hex digits (as Word writes them); "auto" (Word's) is none.
function M.text_color(v)
  v = vim.trim(v or ""):lower()
  if M.COLORS[v] then
    return v
  end
  local hex = v:match("^#?(%x%x%x%x%x%x)$")
  return hex and M.color_of_rgb(hex) or nil
end
-- Its six hex digits ("ff0000").
function M.color_hex(v)
  local c = M.text_color(v)
  return c and (M.COLORS[c] or c:sub(2)) or nil
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
-- A paragraph break in deleted text, in a span's text (see init.lua): written as a "¶" that isn't
-- escaped, which reads back as the break (a "¶" in the text itself is escaped, "\¶").
M.DELETED_BREAK = "\u{E000}"

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

-- A table of contents, a list of figures or of tables: a line of its own, as in LaTeX (see init.lua).
M.INDEX_LINES = { contents = "\\tableofcontents", figures = "\\listoffigures", tables = "\\listoftables" }
function M.prefix(block)
  if block.vspace then -- (vertical space: an empty paragraph of that height, its line all of it)
    return "\\vspace{" .. block.vspace .. "}"
  elseif block.toc then
    return M.INDEX_LINES[block.toc]
  end
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
    kind = ("  "):rep(list.level - 1) .. (list.kind == "bullet" and "- " or "#. ")
  end
  -- its own format's attributes, in braces (after its mark, if it has one: see paras.lua)
  local attrs = block.para and require("neoffice.paras").render(block.para) or ""
  if block.align == "none" then -- (its style's alignment, though the paragraphs around it have another)
    attrs = "align = none" .. (attrs ~= "" and ", " .. attrs or "")
  end
  if attrs ~= "" then
    return kind .. (ALIGN[block.align] or "") .. "{" .. attrs .. "} "
  end
  return kind .. (ALIGN[block.align] or "")
end

-- A vertical space's length as \vspace shows it, from its length in cm: in cm when it is a whole
-- number of millimetres (1cm, 0.5cm), else in points (12pt, 7.5pt).
function M.vspace_shown(cm)
  local mm = cm * 10
  if math.abs(mm - math.floor(mm + 0.5)) < 0.02 then
    return (("%.1f"):format(math.floor(mm + 0.5) / 10):gsub("%.0$", "")) .. "cm"
  end
  local pt = cm / 2.54 * 72
  return (("%.1f"):format(math.floor(pt * 2 + 0.5) / 2):gsub("%.0$", "")) .. "pt"
end

-- A line with only a mark and its attributes ("<=> {first = 0.5cm}"), which sets the format of
-- the paragraphs after it (see paras.lua): { align (nil: left), attrs }, and why if they aren't
-- right (bad); nil if the line isn't one.
-- Also a line with only the attributes, "{align = justify, first = 0.5cm}". It says what changes
-- from the format before it (what it doesn't say goes on; "name = none": the style's again), and
-- "none" alone goes back to the style for all ("{none}", "{none, size = 10pt}"; "{}" too).
-- { reset, align (or "none"), attrs (with "none"s), bad, text }; nil if the line isn't one.
-- (A table's options are a line in braces too, but just above the table: the caller tells.)
function M.para_line(line)
  local sticky -- (a "+" at its end: for the paragraphs from here on; without: for the next one only)
  line = line:gsub("%s*%+%s*$", function() sticky = true return "" end)
  local function read(align, inside)
    local reset = inside:match("^%s*$") ~= nil
    local rest = inside:gsub("^%s*none%s*$", function() reset = true return "" end)
    rest = rest:gsub("^%s*none%s*[,;]%s*", function() reset = true return "" end)
    local attrs, bad = require("neoffice.paras").parse(rest)
    if attrs and attrs.align then
      if align and align ~= attrs.align then
        bad = "its alignment is said twice (its mark, and align =)"
      end
      align, attrs.align = attrs.align, nil
    end
    return { reset = reset or nil, align = align, attrs = attrs or {}, bad = bad, text = inside, sticky = sticky }
  end
  local inside = line:match("^{([^{}]*)}%s*$")
  if inside and (inside:match("^%s*$") or inside:find("=", 1, true) or inside:match("^%s*none%s*$") or inside:match("^%s*none%s*[,;]")) then
    return read(nil, inside)
  end
  for align, mark in pairs(ALIGN) do
    local marker = vim.trim(mark)
    if line:sub(1, #marker) == marker then
      local rest = line:sub(#marker + 1)
      if rest:match("^%s*$") then
        return { align = align, attrs = {}, sticky = sticky }
      end
      inside = rest:match("^%s*{([^{}]*)}%s*$")
      if inside then
        return read(align, inside)
      end
    end
  end
end

-- A format's line as written: "{none, align = justify, first = 0.5cm}" (see M.para_line), and "+"
-- after it for one that goes on (f.sticky).
-- Its alignment as a mark ("<|> {size = 10pt}", "<=>"): shorter, as a paragraph's.
function M.render_para_line(f)
  local mark = ALIGN[f.align or ""]
  local attrs = require("neoffice.paras").render(vim.tbl_extend("force", {}, f.attrs or {}, { align = not mark and f.align or nil }))
  local braces = (f.reset or attrs ~= "" or not mark) and "{" .. (f.reset and "none" .. (attrs ~= "" and ", " or "") or "") .. attrs .. "}" or ""
  return vim.trim((mark or "") .. braces) .. (f.sticky and "+" or "")
end
-- A format (the whole of it: { align, attrs }) after this format's line, on `before`'s.
function M.apply_para_line(before, f)
  local out = { align = before.align, attrs = {} }
  if f.reset then
    out.align = nil
  else
    for k, v in pairs(before.attrs or {}) do
      out.attrs[k] = v
    end
  end
  if f.align then
    out.align = f.align ~= "none" and f.align or nil
  end
  for k, v in pairs(f.attrs or {}) do
    out.attrs[k] = v ~= "none" and v or nil
  end
  return out
end
-- The line that takes `before` to `after` (both whole formats): what changes, or from the style
-- when that says it in much less.
function M.para_line_for(before, after)
  local pf = require("neoffice.paras")
  local change = { attrs = {} }
  if not pf.same({ align = before.align }, { align = after.align }) then
    change.align = after.align or "none"
  end
  for _, key in ipairs(pf.KEYS) do
    local was, now = (before.attrs or {})[key], (after.attrs or {})[key]
    if key ~= "align" and pf.measure(key, was) ~= pf.measure(key, now) then
      change.attrs[key] = now or "none"
    end
  end
  local delta = M.render_para_line(change)
  local fresh = M.render_para_line({ reset = true, align = after.align, attrs = after.attrs })
  return #fresh < #delta * 0.75 and fresh or delta
end

-- The document's header, at its top: its paragraph styles, and what each gives (see init.lua):
--   ---
--   styles:
--     Normal:    {font = Times New Roman, size = 11pt, align = justify}
--     Heading 1: {font = Arial, size = 14pt, bold}
--   ---
-- { last = its last row, styles = { { row, name, text (in its braces), attrs, bad } } }, or nil.
-- (and the document's page settings: "page: {size = A4, margins = 2.5cm}", see pages.lua)
function M.header(lines)
  local second, at = nil, 2 -- (its first line but comments: "# ...", as in YAML)
  while lines[at] and lines[at]:match("^%s*#") do
    at = at + 1
  end
  second = lines[at] or ""
  if lines[1] ~= "---" or not (second:match("^styles:%s*$") or second:match("^page:") or second == "---" and at > 2) then
    return nil
  end
  local out = { styles = {} }
  for r = 2, #lines do
    local l = lines[r]
    if l == "---" then
      out.last = r
      return out
    end
    local name, inside = l:match("^%s+([^:{}]-)%s*:%s*{(.*)}%s*$")
    local page = l:match("^page:%s*{(.*)}%s*$")
    if l:match("^%s*#") then -- (a comment, as in YAML: inactive, as if deleted)
      name = nil
    elseif l:match("^styles:%s*$") then
      name = nil
    elseif page then
      local attrs, bad = require("neoffice.pages").parse(page)
      out.page = { row = r, text = page, attrs = attrs, bad = bad }
    elseif l:match("^page:") then
      out.page = { row = r, bad = "not the page's line: write it as   page: {size = A4, margins = 2.5cm}" }
    elseif name then
      local attrs, bad = require("neoffice.paras").parse_style(inside)
      out.styles[#out.styles + 1] = { row = r, name = vim.trim(name), text = inside, attrs = attrs, bad = bad }
    elseif not l:match("^%s*$") then
      out.styles[#out.styles + 1] = { row = r, bad = "not a style's line: write it as   Name: {font = ..., size = ...}" }
    end
  end
end
-- A style's line at the top: indented two spaces more for each style it is based on (depth), its
-- braces lined up with the others' (width: the widest's indentation and name). (Not YAML as YAML
-- reads it: a key with braces has none under it; neoffice highlights the lines itself, see init.lua)
function M.render_style_line(s, width)
  local lead = ("  "):rep((s.depth or 0) + 1) .. s.name .. ":"
  return lead .. (" "):rep(width + 1 - vim.fn.strdisplaywidth(lead)) .. "{" .. require("neoffice.paras").render_style(s.attrs) .. "}"
end
function M.style_width(styles)
  local width = 0
  for _, s in ipairs(styles or {}) do
    width = math.max(width, 2 + 2 * (s.depth or 0) + vim.fn.strdisplaywidth(s.name) + 1)
  end
  return width
end
function M.render_header(styles, page) -- (styles: { { name, attrs, depth } }, in order, or nil; page: its attributes, written)
  local width = M.style_width(styles)
  local lines = { "---" }
  if page then
    lines[#lines + 1] = "page: {" .. page .. "}"
  end
  if styles then
    lines[#lines + 1] = "styles:"
  end
  styles = styles or {}
  for _, s in ipairs(styles) do
    lines[#lines + 1] = M.render_style_line(s, width)
  end
  lines[#lines + 1] = "---"
  return lines
end

-- A section's line ("\newpage {orientation = landscape}", alone on its line: see pages.lua): {
-- attrs (what it changes), bad (why they aren't right) }, or nil if the line isn't one.
function M.section_line(line)
  local inside = line:match("^\\newpage%s*{(.*)}%s*$")
  if not inside then
    return nil
  end
  local attrs, bad = require("neoffice.pages").parse(inside)
  return { attrs = attrs or {}, bad = bad, text = inside }
end

-- The block a line starts with; where its text begins; where its kind's and its alignment's
-- marker characters are (to escape them in a paragraph that isn't of that kind). Escaped markers
-- (\#, \- ...) don't count.
-- A numbered item's marker at the start of a line: not its number but its numbering's pattern, one
-- part for each level shown, each ending in "." or ")": "1." "1.1." "1.a." "1)" "a)" "i)" (1:
-- numbers, a or A: letters, i or I: roman numerals; letters and numerals alone only with ")", as
-- "A. Smith" or "e.g. this" are text). Returns its pattern (as 1, a, A, i, I: "2.b." is "1.a."),
-- its level (its parts' number, or with one part its indentation: two spaces a level), where its
-- first separator is, and where its text starts; nil if the line doesn't start with one.
local function part_class(part)
  if part:match("^%d+$") then
    return "1"
  elseif part:match("^[ivxlcdm]+$") and (#part > 1 or part == "i") then
    return "i"
  elseif part:match("^[IVXLCDM]+$") and (#part > 1 or part == "I") then
    return "I"
  elseif part:match("^%l$") then
    return "a"
  elseif part:match("^%u$") then
    return "A"
  end
end
function M.list_marker(line)
  local indent, rest_at = line:match("^( *)()")
  local p, parts, first_sep = rest_at, {}, nil
  while true do
    local part, sep, after = line:match("^(%w+)([.)])()", p)
    local class = part and part_class(part)
    if not class then
      return nil
    end
    parts[#parts + 1] = class .. sep
    first_sep = first_sep or (after - 1)
    p = after
    if line:sub(p, p) == " " then
      local first = line:match("^(%w+)", rest_at)
      if not first:match("^%d+$") and (#parts > 1 or sep ~= ")") then
        return nil -- (letters and numerals alone: only with ")")
      end
      local level = #parts > 1 and #parts or math.floor(#indent / 2) + 1
      return table.concat(parts), level, first_sep, p + 1
    end
  end
end

function M.block(line)
  local block, start, kind_at = {}, 1, nil
  local hashes = line:match("^(#+) ")
  local indent, after = line:match("^( *)%- ()")
  -- (a numbered item: "#. ", as pandoc's; its level by its indentation, two spaces a level; its
  -- numbers as its list's in the document. "1. ", "2.1. " are text: never escaped)
  local nindent, nafter = line:match("^( *)#%. ()")
  local word, after_word = line:match("^([^%s:\\]+): ()")
  local caption = word and caption_kind(word)
  local vspace = line:match("^\\vspace%s*{%s*([^{}]-)%s*}%s*$") or line:match("^\\vspace%s*$") and "1cm" -- (alone: 1cm)
  for kind, text in pairs(M.INDEX_LINES) do -- (a table of contents..., its line all of it)
    if line:match("^" .. vim.pesc(text) .. "%s*$") then
      return { toc = kind }, #line + 1, 1, nil
    end
  end
  if vspace then -- (vertical space, its line all of it: "\vspace{1cm}", see M.vspace_shown)
    local cm = tonumber(require("neoffice.paras").measure("before", vspace) or "")
    return { vspace = vspace, vspace_cm = cm and cm > 0 and cm or nil,
      vspace_bad = not (cm and cm > 0) and ("\\vspace{%s}: not a length above 0 (1cm, 12pt, 5mm, 0.5in)"):format(vspace) or nil }, #line + 1, 1, nil
  end
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
  elseif nindent then
    block, start, kind_at = { list = { kind = "number", level = math.floor(#nindent / 2) + 1 } }, nafter, #nindent + 1
  end
  local align_at
  for _, align in ipairs({ "center", "justify", "right", "left" }) do
    local marker = ALIGN[align]
    if line:sub(start, start + #marker - 1) == marker then
      block.align, align_at, start = align, start, start + #marker
      break
    elseif line:sub(start) == vim.trim(marker) then -- (alone, ending the text: an empty paragraph's, as a
      -- table cell's, whose spaces at its end don't show)
      block.align, align_at, start = align, start, #line + 1
      break
    end
  end
  -- (its own format's attributes, after its mark or alone: "<=> {first = 0} ...", "{before = 6pt} ...";
  -- alone, only "name = value"s)
  local inside, after, brace_at
  if not kind_at then -- (plain paragraphs only)
    inside, after = line:match("^{([^{}]*)} ()", start)
    if inside and not align_at and not inside:find("=", 1, true) then
      inside = nil
    end
    brace_at = inside and start
  end
  if inside then
    local attrs, bad = require("neoffice.paras").parse(inside)
    if attrs and attrs.align then
      if block.align and block.align ~= attrs.align then
        attrs, bad = nil, "its alignment is said twice (its mark, and align =)"
      else
        block.align, attrs.align = attrs.align, nil
      end
    end
    if attrs then
      block.para, start = attrs, after
    else
      block.para_bad = bad
      start = after
    end
  end
  return block, start, kind_at, align_at, brace_at
end

local function kind(block)
  return block.heading or block.title or block.quote or block.list or block.row or block.caption or block.fence or block.vspace or block.toc
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

-- Whether line `row` of a table whose first line is `first` (rows of `lines`) is its line of
-- dashes (its format), not a row: right under its first row, any such line; else (its first row
-- deleted, a row put between them, a row deleted with changes tracked, shown above it...) one with
-- three dashes at least in each cell, if the table has none above it (`seen`) and the next line
-- isn't one too. (A row of "-", "--" for values missing stays a row.)
local function three_dashes(line) -- (a line of dashes with three at least in each cell)
  if not M.table_format(line) then
    return false
  end
  for cell in line:sub(2):gmatch("([^|]*)|?") do
    cell = vim.trim(cell)
    if cell ~= "" and not cell:match("^:?%-%-%-+:?$") then
      return false
    end
  end
  return true
end
function M.is_dashes(lines, first, row, seen)
  local line = lines[row] or ""
  if seen or not M.table_format(line) then -- (one line of dashes: after it, rows)
    return false
  elseif row == first + 1 then
    return true
  end
  return three_dashes(line) and not three_dashes(lines[row + 1] or "")
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
-- (`lead`: whether it starts with them, whatever comes after: text typed after a paragraph's floating
-- images, on their line)
function M.floats_only(line, lead)
  if not line:find("float", 1, true) then -- (quick: most lines)
    return false
  end
  local _, pos = M.block(line)
  local floats = require("neoffice.floats")
  local any = false
  while pos <= #line do
    if line:sub(pos, pos + 1) ~= "![" then
      return lead and any or false
    end
    local _, _, after, _, attrs = M.image_at(line, pos)
    if not (after and attrs and attrs.attrs and floats.floating(attrs.attrs)) then
      return lead and any or false
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
    if s then -- (a value of several words, up to the next comma: "font = Times New Roman"; or quoted)
      local at = text:match("^%s*[%a_][%w_%-]*%s*=%s*()", p)
      local quoted, qe = text:match('^"([^"]*)"()', at)
      local words, we = text:match("^([^,;=\"]-)%s*()[,;]", at)
      if not words then
        words, we = text:match("^([^,;=\"]-)%s*()$", at)
      end
      if quoted then
        value, e = quoted, qe
      elseif words and words:find(" ", 1, true) and (name:lower() == "font" or name:lower() == "style" or name:lower() == "size" and words:match("^%S+%s*[x×]%s*%S+$") or name:lower() == "margins" and words:match("^%S+%s+%S+$")) then -- (only a font's name and a style's have spaces, a page's size: "21cm x 29.7cm", and its margins: "2.5cm 3cm")
        value, e = words, we
      end
    end
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

-- The attributes of a span of text ("[words]{font = Arial, size = 9pt, color = red}"): { font,
-- size, color }, or nil and why. (Its size as shown: in pt; its colour: by name, or "#rrggbb".)
M.SPAN_KEYS = { "font", "size", "color" }
function M.span_attrs(text)
  local out, any = {}, false
  for _, a in ipairs((M.attributes(text))) do
    local name = a.name == "colour" and "color" or a.name
    if name ~= "font" and name ~= "size" and name ~= "color" then
      return nil, ("not an attribute of text: %s (they are font, size, color)"):format(a.name or a.value)
    end
    a.name = name
    local value = vim.trim(a.value)
    if a.name == "color" then
      value = M.text_color(value)
      if not value then
        return nil, "not a colour: " .. a.value .. " (red, darkblue... or #rrggbb)"
      end
    elseif a.name == "size" then
      local cm = M.cm(value)
      if not cm or cm <= 0 then
        return nil, "not a size: " .. a.value .. " (11pt, 12pt...)"
      end
      local pt = cm * 72 / 2.54
      value = (("%.1f"):format(pt):gsub("%.?0+$", "")) .. "pt"
    elseif value == "" then
      return nil, "a font's name is missing"
    end
    out[a.name], any = value, true
  end
  if not any then
    return nil, "a font, a size or a colour is missing (font = Arial, size = 9pt, color = red)"
  end
  return out
end
-- (a value as written: in quotes if it holds what separates values, as "Times New Roman;serif")
function M.attr_value(v)
  return (v:find("[,;=\"{}]") or v ~= vim.trim(v)) and '"' .. v:gsub('"', "") .. '"' or v
end
function M.render_span_attrs(s)
  local parts = {}
  for _, key in ipairs(M.SPAN_KEYS) do
    if s[key] then
      parts[#parts + 1] = key .. " = " .. M.attr_value(s[key])
    end
  end
  return table.concat(parts, ", ")
end
-- (what tells spans of text apart: their font and size)
local function span_key(s)
  return (s.font or s.size or s.color) and ((s.font or "") .. "|" .. (s.size or "") .. "|" .. (s.color or "")) or nil
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
        and not (upto and rest:sub(1, upto - 1) or rest):find("%]%(") and not (upto and rest:sub(1, upto - 1) or rest):find("\\%]%(")
        and not (upto and rest:sub(1, upto - 1) or rest):find("%]{") and not (upto and rest:sub(1, upto - 1) or rest):find("\\%]{"))
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
    if span_key(cur) ~= span_key(to) then -- (a span of text in another font: everything closes, and opens again inside)
      if cur.link then
        put("](" .. cur.link:gsub("%)", "%%29") .. ")")
      end
      for k = #M.FLAGS, 1, -1 do
        if cur[M.FLAGS[k]] then
          put(closing(M.FLAGS[k], cur[M.FLAGS[k]]))
        end
      end
      if span_key(cur) then
        put("]{" .. M.render_span_attrs(cur) .. "}")
      end
      if span_key(to) then
        if (parts[#parts] or ""):sub(-1) == "!" then -- (so "![" doesn't read as an image)
          parts[#parts], len = parts[#parts]:sub(1, -2) .. "\\!", len + 1
        end
        put("[")
      end
      cur = { font = to.font, size = to.size }
    end
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
      local text = s.raw and s.text or escape(s.text, s.sup or s.sub, block.cell, strict):gsub(M.DELETED_BREAK, "¶")
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
      local parsed, _, kind_at, align_at, brace_at = M.block(text)
      local at = text:match("^%s+$") and 1
      if not at and (start > 1 or not kind(block)) and kind(parsed) and not (block.cell and parsed.row) then
        at = kind_at
      elseif not at and (start > 1 or not block.align) and parsed.align then
        at = align_at
      elseif not at and start == 1 and not block.para and block.align ~= "none" and brace_at and (parsed.para or parsed.para_bad) then
        at = brace_at -- (text in braces that would read as its attributes: "\{")
      elseif not at and start == 1 and not kind(block) and not block.align and (text == "---" or M.para_line(text)) then
        at = 1 -- (text that would read as a format's line, or the header's: "\{...}", "\---")
      elseif not at and text:match("^\\newpage%s*{.*}%s*$") then
        at = #text:match("^\\newpage%s*") + 1 -- (a page break and braces that would read as a section's line: "\newpage \{")
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
    if plain ~= line and M.same(plain, spans) then
      line, cols = plain, pcols
    end
  end
  -- what is written must read back as the same text: if not (marks next to text that looks like
  -- marks), it is written again escaping every character that could start a mark
  -- (only text with characters that could be read as marks can read back otherwise)
  if not strict and could_misread(line) and not M.same(line, spans) then
    return M.render(spans, block, true)
  end
  -- (nor as a format's line, which it would be on its own: a paragraph whose text looks like
  -- attributes, after its mark, "<=> \{first = 0.5cm}")
  if not strict and not block.cell and line:find("{", 1, true) and M.para_line(line)
    and vim.iter(spans):any(function(s) return s.text and s.text ~= "" end) then
    return M.render(spans, block, true)
  end
  -- (an empty table cell's paragraph with its alignment: its mark alone, as a cell shows no space
  -- at its end, and reads it so: see M.block)
  if block.cell and line:match("^%S+ $") and M.block(line:sub(1, -2)).align then
    line = line:sub(1, -2)
    for _, c in ipairs(cols) do
      c[1], c[2] = math.min(c[1], #line), math.min(c[2], #line)
    end
  end
  return line, cols
end

-- Whether a line reads back as these spans: the same characters, formatting and objects.
-- (compared character by character, by what each reads as: a flag it hasn't is false, a highlight
-- that is true is yellow; an object by its kind)
local NONE = {}
local function flag_value(v, k)
  if not v then
    return ""
  end
  return v == true and k == "h" and "yellow" or tostring(v)
end
local function kind_of(c)
  return c.image and "image" or c.math and "math" or c.footnote and "footnote" or c.field and "field" or c.page_break and "page_break" or nil
end
function M.same(line, spans)
  local _, chars = M.parse(line)
  local k = 0
  local function reads(ch, fmt) -- (the next character reads as ch, formatted so)
    k = k + 1
    local c = chars[k]
    if c and ch == M.DELETED_BREAK then -- (a deleted paragraph break: a "¶" not escaped; one in the text is)
      ch = c.ch == "¶" and not c.esc and "¶" or ch
    elseif c and ch == "¶" and not c.esc then
      return false
    end
    if not c or kind_of(c) or c.ch ~= ch or (c.link or "") ~= (fmt.link or "") or (c.font or "") ~= (fmt.font or "")
      or (c.size or "") ~= (fmt.size or "") or (c.color or "") ~= (fmt.color or "") then
      return false
    end
    if not (c.b or c.i or c.u or c.s or c.h or c.sup or c.sub) then -- (most: none, on both sides)
      return not (fmt.b or fmt.i or fmt.u or fmt.s or fmt.h or fmt.sup or fmt.sub)
    end
    return not c.b == not fmt.b and not c.i == not fmt.i and not c.u == not fmt.u and not c.s == not fmt.s
      and not c.sup == not fmt.sup and not c.sub == not fmt.sub and flag_value(c.h, "h") == flag_value(fmt.h, "h")
  end
  for _, s in ipairs(spans) do
    local token = kind_of(s)
    if token then
      k = k + 1
      if not chars[k] or kind_of(chars[k]) ~= token then
        return false
      end
    elseif s.line_break then
      if not reads("\n", NONE) then
        return false
      end
    elseif s.text then
      local fmt = s.raw and NONE or s
      for ch in s.text:gmatch("[^\128-\191][\128-\191]*") do
        if not reads(ch, fmt) then
          return false
        end
      end
    end
  end
  return k == #chars
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
for c in ("\\*[]$<>^_~=+|#%-.!{}:)\n" .. "⟦⟧¶"):gmatch("[^\128-\191][\128-\191]*") do
  ESCAPED[c] = true
end

-- The first bytes of everything that isn't plain text (marks, escapes, images, equations...).
local PLAIN_STOP = "[\\!%$%^_%*~=%+}%[\226]" -- (the same, as a pattern: where plain text stops)
local MARKS = {}
for c in ("\\!$^_*~=+}[" .. "\226"):gmatch(".") do -- (\226: the first byte of ⟦)
  MARKS[c:byte()] = true
end

-- Marks that close nothing, in a text: what would be written as characters, not formatting
-- (a "**" whose other one was deleted...). { col (0-based), mark, flag (what it would format) }, in
-- order. Escaped ones ("\*\*") are characters on purpose; so is a lone "*" between spaces or
-- letters ("2 * 3", "2*3"): only one where a word starts ("*word") counts. A "~~", "==" or "++" is
-- written as it is when it can't be a mark ("C++", "a == b"): one counts where a word starts too,
-- unless `anywhere`.
M.ORPHAN_FLAGS = { ["**"] = "b", ["*"] = "i", ["~~"] = "s", ["=="] = "h", ["++"] = "u", ["^{"] = "sup", ["_{"] = "sub" }
local SHAPED = { ["~~"] = true, ["=="] = true, ["++"] = true }
function M.orphan_marks(text, anywhere)
  local out = {}
  if not text:find("[*~=+^_]") then
    return out
  end
  local _, chars = M.parse_shared(text)
  local function plain(c) -- (a character as written, not escaped, not an image...)
    return c and not c.esc and not c.len
  end
  local k = 1
  while k <= #chars do
    local c, n, p = chars[k], chars[k + 1], chars[k - 1]
    local next_to = plain(n) and n.col == c.col + #c.ch
    local two = plain(c) and next_to and c.ch .. n.ch
    local after = two and chars[k + 2]
    local starts = not (p and p.col + #p.ch == c.col and not p.ch:find("^%s")) -- (nothing right before it but a space)
      and after and after.col == n.col + #n.ch and not after.ch:find("^%s")
    if two and M.ORPHAN_FLAGS[two] and #two == 2 and (anywhere or not SHAPED[two] or starts) then
      out[#out + 1] = { col = c.col, mark = two, flag = M.ORPHAN_FLAGS[two] }
      k = k + 2
    else
      if plain(c) and c.ch == "*" and next_to and not n.ch:find("^%s") and n.ch ~= "*"
        and not (p and p.col + #p.ch == c.col and not p.ch:find("^%s")) then -- (where a word starts)
        out[#out + 1] = { col = c.col, mark = "*", flag = "i" }
      end
      k = k + 1
    end
  end
  -- (and a "^{" or "_{" whose "}" was deleted: what follows it is a superscript, or subscript, to
  -- the end)
  local last = chars[#chars]
  for _, key in ipairs(last and { "sup", "sub" } or {}) do
    if last[key] and not text:sub(last.col + (last.len or #last.ch) + 1):find("}", 1, true) then
      local first = #chars
      while first > 1 and chars[first - 1][key] do
        first = first - 1
      end
      local mark = key == "sup" and "^{" or "_{"
      local at = text:sub(1, chars[first].col):match(".*()" .. vim.pesc(mark))
      out[#out + 1] = { col = at and at - 1 or chars[first].col, mark = mark, flag = key, spreads = true }
    end
  end
  table.sort(out, function(x, y) return x.col < y.col end)
  return out
end

-- M.parse's result for a text, the same tables each time for the last few texts read: only to be
-- read, never changed (typing reads the line it changes again and again: before the edit and after
-- it, and to decorate it).
local shared = {}
function M.parse_shared(text)
  for k, entry in ipairs(shared) do
    if entry[1] == text then
      if k > 1 then -- (the most recent first)
        table.remove(shared, k)
        table.insert(shared, 1, entry)
      end
      return entry[2], entry[3]
    end
  end
  local block, chars = M.parse(text)
  table.insert(shared, 1, { text, block, chars })
  shared[9] = nil
  return block, chars
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
  local span_close, span_after, spans_open = nil, nil, {} -- (the open span of text in another font)
  local h_from -- (where the open highlight started, in chars)
  local function emit(at)
    local byte = line:byte(at) -- (the character's bytes, from its first: UTF-8)
    local ch = line:sub(at, at + (byte < 0xC0 and 0 or byte < 0xE0 and 1 or byte < 0xF0 and 2 or 3))
    local c = { -- (the flags of M.FLAGS: only those it has, nil for the others)
      col = at - 1, ch = ch, link = cur.link, b = cur.b == true or nil, i = cur.i == true or nil, u = cur.u == true or nil,
      s = cur.s == true or nil, h = cur.h or nil, sup = cur.sup == true or nil, sub = cur.sub == true or nil,
    }
    if span_close then -- (in a span of text in another font)
      c.font, c.size, c.color, c.span_bad = cur.font, cur.size, cur.color, cur.span_bad
    end
    chars[#chars + 1] = c
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
  -- "[" starts a span of text in another font if a matching "]{...}" follows (its "]", its
  -- attributes, and the position after them)
  local function find_span(at)
    local p, depth = at + 1, 0
    while p <= #line do
      local c = line:sub(p, p)
      if c == "\\" then
        p = p + 2
      elseif c == "[" then
        depth, p = depth + 1, p + 1
      elseif c == "]" and depth > 0 then
        depth, p = depth - 1, p + 1
      elseif c == "]" then
        local inside, after = line:match("^{([^}]*)}()", p + 1)
        return inside and p, inside, after
      else
        p = p + 1
      end
    end
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

  local n = #line
  while pos <= n do
    if not MARKS[line:byte(pos)] and pos ~= link_close and pos ~= span_close then -- (most characters: just text)
      -- (up to the next character that could start a mark, or close a link or a span: all at once,
      -- with the formatting as it is)
      local stop = line:find(PLAIN_STOP, pos) or n + 1
      if link_close and link_close > pos and link_close < stop then
        stop = link_close
      end
      if span_close and span_close > pos and span_close < stop then
        stop = span_close
      end
      local link, b, i, u, s, h, sup, sub = cur.link, cur.b == true or nil, cur.i == true or nil, cur.u == true or nil,
        cur.s == true or nil, cur.h or nil, cur.sup == true or nil, cur.sub == true or nil
      local in_span, font, size, color, bad = span_close, cur.font, cur.size, cur.color, cur.span_bad
      local bare = not (link or b or i or u or s or h or sup or sub or in_span) -- (plain text: just where it is, and what)
      local count = #chars
      while pos < stop do
        local byte = line:byte(pos)
        local len = byte < 0xC0 and 1 or byte < 0xE0 and 2 or byte < 0xF0 and 3 or 4
        local c
        if bare then
          c = { col = pos - 1, ch = line:sub(pos, pos + len - 1) }
        else
          c = { col = pos - 1, ch = line:sub(pos, pos + len - 1), link = link, b = b, i = i, u = u, s = s, h = h, sup = sup, sub = sub }
          if in_span then
            c.font, c.size, c.color, c.span_bad = font, size, color, bad
          end
        end
        count = count + 1
        chars[count] = c
        pos = pos + len
      end
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
    if pos == span_close and pos ~= link_close then
      cur.font, cur.size, cur.color, cur.span_bad, pos = nil, nil, nil, nil, span_after
      span_close, span_after = unpack(table.remove(spans_open) or {})
    elseif pos == link_close then
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
    elseif c == "[" and not cur.link and not span_close and find_span(pos) then
      local close, inside, after = find_span(pos)
      local attrs, bad = M.span_attrs(inside)
      table.insert(spans_open, { span_close, span_after })
      span_close, span_after = close, after
      cur.font, cur.size, cur.color, cur.span_bad = attrs and attrs.font, attrs and attrs.size, attrs and attrs.color, bad
      pos = pos + 1
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
