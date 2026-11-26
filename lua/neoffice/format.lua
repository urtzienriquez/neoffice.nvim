-- Formatting a document's text (the language server's textDocument/formatting): as neoffice
-- itself writes it when it opens a document.
--   tables      a line of dashes under the first row if there is none, a cell for every column,
--               and the columns aligned (see tables.text_widths)
--   attributes  in braces (a table's options, an image's width): "name = value", separated by
--               commas (see markup.attributes); a highlight's colour alone ("{green}")
-- The edits only touch what is between the cells' text (spaces and "|") and inside braces, never
-- the text itself: its formatting, comments and tracked changes (marks on it) stay as they are.
local markup = require("neoffice.markup")
local tables = require("neoffice.tables")
local M = {}

-- An edit replacing `old` (bytes a..b of line `row`, 0-based, b excluded) by `new`, as small as it
-- can be (only what differs).
local function edit(out, row, a, old, new)
  if old == new then
    return
  end
  local p, s = 0, 0
  while p < math.min(#old, #new) and old:byte(p + 1) == new:byte(p + 1) do
    p = p + 1
  end
  while s < math.min(#old, #new) - p and old:byte(#old - s) == new:byte(#new - s) do
    s = s + 1
  end
  out[#out + 1] = {
    range = { start = { line = row, character = a + p }, ["end"] = { line = row, character = a + #old - s } },
    newText = new:sub(p + 1, #new - s),
  }
end

-- The edits for a table: its lines `first`..`last` (1-based rows of `lines`); `numbers`: its
-- columns of numbers to align right (see init.lua: number_columns).
local function format_table(out, lines, first, last, numbers)
  local format = markup.table_format(lines[first + 1] or "")
  local has_dashes = format ~= nil
  if numbers then -- (as saving it would)
    format = format or { align = {} }
    for col in pairs(numbers) do
      format.align[col] = format.align[col] or "right"
    end
  end
  local rows = {} -- (the rows of cells: line numbers)
  for r = first, last do
    if not (r == first + 1 and has_dashes) then
      rows[#rows + 1] = r
    end
  end
  local spans, texts, columns = {}, {}, 0
  for k, r in ipairs(rows) do
    spans[k] = markup.cell_spans(lines[r])
    texts[k] = vim.tbl_map(function(sp) return lines[r]:sub(sp.s, sp.e) end, spans[k])
    columns = math.max(columns, #spans[k])
  end
  local widths = tables.text_widths(texts)
  local aligns = format and format.align or {}
  local dashes = markup.format_line(format or { align = {} }, columns, widths)
  for k, r in ipairs(rows) do
    local line, cells = lines[r], spans[k]
    -- the text between the cells, as it is and as it should be
    local gaps, want = {}, {}
    for c = 0, #cells do
      local a = c == 0 and 1 or cells[c].e + 1
      local b = c == #cells and #line or cells[c + 1].s - 1
      gaps[c] = { a = a, text = line:sub(a, b) }
      local after, before = "", ""
      if c > 0 then -- (after the text of cell c, before that of the next)
        after = select(2, tables.padding(texts[k][c], widths and widths[c], aligns[c]))
      end
      if c < #cells then
        before = (tables.padding(texts[k][c + 1], widths and widths[c + 1], aligns[c + 1]))
      end
      want[c] = #cells == 0 and "|" or c == 0 and "| " .. before or c == #cells and after .. " |" or after .. " | " .. before
    end
    for c = #cells + 1, columns do -- (the cells a row lacks: empty)
      want[#cells] = want[#cells] .. " " .. (" "):rep(widths and widths[c] or 0) .. " |"
    end
    if k == 1 and not has_dashes then -- (the line of dashes it lacks: after the first row)
      want[#cells] = want[#cells] .. "\n" .. dashes
    end
    for c = 0, #cells do
      edit(out, r - 1, gaps[c].a - 1, gaps[c].text, want[c])
    end
  end
  if has_dashes then
    edit(out, first, 0, lines[first + 1], dashes)
  end
  -- its options, above it
  local options = first > 1 and markup.table_options(lines[first - 1])
  if options and pcall(tables.parse, options) then
    local _, tidy = markup.attributes(options, tables.name_of)
    edit(out, first - 2, 0, lines[first - 1], "{" .. tidy .. "}")
  end
end

-- The edits (LSP TextEdits) that format a buffer's lines, or those of rows first..last (0-based).
function M.edits(buf, first, last)
  local lines = vim.api.nvim_buf_get_lines(buf, 0, -1, false)
  first, last = first or 0, last or #lines - 1
  local numbers = require("neoffice").number_columns(buf)
  local out, r = {}, 1
  while r <= #lines do
    if lines[r]:match("^|") and #(markup.cells(lines[r]) or {}) > 0 then -- (a table: up to a line that isn't one)
      local stop = r
      while lines[stop + 1] and lines[stop + 1]:match("^|") do
        stop = stop + 1
      end
      if stop - 1 >= first and r - 1 <= last then
        format_table(out, lines, r, stop, numbers[r])
      end
      r = stop + 1
    else
      r = r + 1
    end
  end
  for row = first + 1, math.min(last + 1, #lines) do -- boxes' lines (no marks on them: the whole line)
    local fence, attrs = markup.box_line(lines[row])
    local parsed = fence and attrs and require("neoffice.floats").parse(attrs)
    if parsed then
      local tidy = require("neoffice.floats").render(parsed)
      edit(out, row - 1, 0, lines[row], tidy ~= "" and "::: {" .. tidy .. "}" or ":::")
    end
  end
  for row = first + 1, math.min(last + 1, #lines) do -- images' widths and highlights' colours
    for s, inside in lines[row]:gmatch("=={()([^{}]*)}") do
      local name = inside:match("^%s*colou?r%s*=%s*([#%w]+)%s*$") or inside:match("^%s*([#%w]+)%s*$")
      local color = name and markup.color(name)
      if color then -- ("{Green}", "{color = green}": "{green}")
        edit(out, row - 1, s - 1, inside, color)
      end
    end
    local line, p = lines[row], lines[row]:find("![", 1, true)
    while p do
      local _, _, after, _, attrs = markup.image_at(line, p)
      local tidy = attrs and attrs.attrs and require("neoffice.floats").render(attrs.attrs)
      if tidy and tidy ~= "" then -- (only inside the braces: the image's mark ends with them)
        edit(out, row - 1, attrs.s, line:sub(attrs.s + 1, attrs.e - 1), tidy)
      end
      p = line:find("![", math.max(after or 0, p + 2), true)
    end
  end
  table.sort(out, function(a, b)
    return a.range.start.line < b.range.start.line or a.range.start.line == b.range.start.line and a.range.start.character < b.range.start.character
  end)
  return out
end

return M
