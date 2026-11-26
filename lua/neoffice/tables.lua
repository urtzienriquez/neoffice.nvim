-- Tables on save. The tables, rows and cells of the buffer are matched with the document's
-- through the paragraphs in them (a row is the document's row its paragraphs came from). Rows and
-- cells typed in the buffer are added (copies of their neighbours, emptied); those deleted from the
-- buffer are removed. Each paragraph in a cell is then written as any other paragraph.
--
-- Tables also typeset themselves (see M.WORDS): a table's look (its rules, column widths, row
-- height, position) comes from a few options, on a line above it ("{width = full, height = tall}");
-- those it doesn't say are the defaults (the option `tables`).
local xml = require("neoffice.xml")
local M = {}

local PARAGRAPHS = { ["w:p"] = true, ["text:p"] = true, ["text:h"] = true }
local CONTENT = { ["w:p"] = true, ["w:tbl"] = true, ["w:sdt"] = true, ["text:p"] = true, ["text:h"] = true, ["text:list"] = true, ["table:table"] = true }
local CELLS = { ["w:tc"] = true, ["table:table-cell"] = true }
local ROWS = { ["w:tr"] = true, ["table:table-row"] = true }

-- ── Table formats ─────────────────────────────────────────────────────────

-- The options of a table's look, and their values (on the line above it: "{width = full}").
-- width is also a percentage of the text's width ("80%") or a length ("12cm").
M.ORDER = { "rules", "width", "height", "position" }
M.WORDS = {
  rules = { "booktabs", "header", "rows", "grid", "none" },
  width = { "auto", "full" },
  height = { "compact", "normal", "tall" },
  position = { "center", "left", "right" },
}
-- What each value does (for completion).
M.DESCRIBE = {
  ["rules.booktabs"] = "a thick line above and below, a thin one under the header (the default)",
  ["rules.header"] = "only a line under the header",
  ["rules.rows"] = "lines above, below and between the rows",
  ["rules.grid"] = "every line",
  ["rules.none"] = "no lines",
  ["width.auto"] = "as narrow as the content (the default)",
  ["width.full"] = "the width of the text",
  ["width.length"] = "a percentage of the text's width, or a length (cm, mm, in)",
  ["height.compact"] = "0.05 cm above and below the text of each cell (the default)",
  ["height.normal"] = "0.1 cm above and below the text of each cell",
  ["height.tall"] = "0.2 cm above and below the text of each cell",
  ["position.center"] = "in the middle of the page (the default)",
  ["position.left"] = "at the left margin",
  ["position.right"] = "at the right margin",
}
local ALIASES = { all = "grid", borders = "rules", align = "position", centre = "center", ["100%"] = "full" }

-- Rules: which lines a table has. Each line is 0 (none), 1 (thin) or 2 (thick); `header` is the
-- line under the first row.
M.BORDERS = {
  grid = { top = 1, bottom = 1, left = 1, right = 1, insideH = 1, insideV = 1 },
  none = {},
  booktabs = { top = 2, bottom = 2, header = 1 }, -- (as in scientific papers)
  header = { header = 1 },
  rows = { top = 1, bottom = 1, insideH = 1 },
}

function M.defaults()
  local config = require("neoffice").config.tables or {}
  local out = {}
  for _, key in ipairs(M.ORDER) do
    out[key] = config[key] or M.WORDS[key][1]
  end
  return out
end

local function is_width(word)
  local percent = tonumber(word:match("^([%d.]+)%%$") or "")
  return word == "auto" or word == "full" or percent and percent > 0 and percent <= 100 or require("neoffice.markup").cm(word) ~= nil
end

-- The option a value is of ("full": width), or nil.
function M.name_of(value)
  value = ALIASES[value:lower()] or value:lower()
  for _, k in ipairs(M.ORDER) do
    if vim.tbl_contains(M.WORDS[k], value) or k == "width" and is_width(value) then
      return k
    end
  end
end

-- What a table's options line says ("width = full, height = tall"; "width=full", or "full" alone,
-- also do): { option = value }. An error for one that isn't, saying which are.
function M.parse(text)
  local out = {}
  for _, a in ipairs((require("neoffice.markup").attributes(text or ""))) do
    local value = a.value:lower()
    value = ALIASES[value] or value
    local key = a.name and (ALIASES[a.name] or a.name)
    if key and not M.WORDS[key] then
      error(("a table's options are %s (not %s)"):format(table.concat(M.ORDER, ", "), a.name))
    end
    local found = M.name_of(value)
    if not found or key and key ~= found and not (key == "width" and is_width(value)) then
      local all = {}
      for _, k in ipairs(M.ORDER) do
        all[#all + 1] = k .. " = " .. table.concat(M.WORDS[k], " | ") .. (k == "width" and " | 80% | 12cm" or "")
      end
      error(("not a table option: %s (they are: %s)"):format(key and key .. " = " .. a.value or a.value, table.concat(all, "; ")))
    end
    out[key or found] = value
  end
  return out
end

-- What a table's options line shows: its options that aren't the defaults ("width = full, height
-- = tall"; unknown ones, nil, aren't shown).
function M.line(options)
  local defaults, words = M.defaults(), {}
  for _, key in ipairs(M.ORDER) do
    if options[key] and options[key] ~= defaults[key] then
      words[#words + 1] = key .. " = " .. options[key]
    end
  end
  return table.concat(words, ", ")
end

-- The lines around a cell of a table with these rules: { top, bottom, left, right } (0, 1 or 2),
-- for the cell in row r of n rows, column c of the m cells of its row.
function M.cell_borders(preset, r, n, c, m)
  local b = M.BORDERS[preset]
  local function line(k) return b[k] or 0 end
  local function inner_h(upper) -- (between row `upper` and the next)
    return math.max(line("insideH"), upper == 1 and line("header") or 0)
  end
  return {
    top = r == 1 and line("top") or inner_h(r - 1),
    bottom = r == n and line("bottom") or inner_h(r),
    left = c == 1 and line("left") or line("insideV"),
    right = c == m and line("right") or line("insideV"),
  }
end

-- The rules a table's cells have: `grid` is its rows, each a list of { top, bottom, left, right }
-- (0 none, 1 thin, 2 thick; nil if a cell doesn't say). nil if none fits. (A line between two cells
-- may be either's: the bottom of one or the top of the one below it.)
function M.borders_of(grid)
  if #grid == 0 then
    return nil
  end
  -- whether each line is there: true, false, or nil (it isn't the same all along, or unknown)
  local function all(list, get) -- (every one says the line is there, or every one says it isn't)
    local on, off, level = true, true, nil
    for _, x in ipairs(list) do
      local w = get(x)
      on, off = on and w ~= nil and w > 0, off and w == 0
      level = w and w > 0 and (level == nil and w or level == w and w or -1) or level
    end
    if on then
      return true, level
    elseif off then
      return false
    end
  end
  local function either(a, b) -- (a line between cells: either's)
    if a == true or b == true then
      return true
    end
    if a == false and b == false then
      return false
    end
  end
  local _, top_level = all(grid[1], function(s) return s.top end)
  local _, bottom_level = all(grid[#grid], function(s) return s.bottom end)
  local lines = {
    top = (all(grid[1], function(s) return s.top end)),
    bottom = (all(grid[#grid], function(s) return s.bottom end)),
    left = all(grid, function(row) return row[1] and row[1].left end),
    right = all(grid, function(row) return row[#row] and row[#row].right end),
    between = {}, -- (under each row but the last)
  }
  for r = 1, #grid - 1 do
    lines.between[r] = either(all(grid[r], function(s) return s.bottom end), all(grid[r + 1], function(s) return s.top end))
  end
  local inside_v, seen -- (the lines between cells in a row: all there, or none)
  for _, row in ipairs(grid) do
    for c = 1, #row - 1 do
      local function on(w)
        if w ~= nil then
          return w > 0
        end
      end
      local line = either(on(row[c].right), on(row[c + 1].left))
      if line == nil or seen and line ~= inside_v then
        return nil
      end
      inside_v, seen = line, true
    end
  end
  local names = vim.tbl_keys(M.BORDERS)
  table.sort(names)
  for _, name in ipairs(names) do
    local b = M.BORDERS[name]
    local function is(k, line) return line == ((b[k] or 0) > 0) end
    local fits = is("top", lines.top) and is("bottom", lines.bottom) and is("left", lines.left) and is("right", lines.right)
      and (not lines.top or top_level == b.top) and (not lines.bottom or bottom_level == b.bottom) -- (thin or thick)
    for r, line in pairs(lines.between) do
      fits = fits and line == ((b.insideH or 0) > 0 or r == 1 and (b.header or 0) > 0)
    end
    fits = fits and (not seen or is("insideV", inside_v))
    if fits then
      return name
    end
  end
end

-- Cell padding (cm): left and right, and above and below for each row height.
M.PADDING = 0.19
M.HEIGHT = { compact = 0.05, normal = 0.1, tall = 0.2 }
function M.height_of(cm) -- (the row height of a padding, if it is one of them)
  for name, value in pairs(M.HEIGHT) do
    if cm and math.abs(cm - value) < 0.006 then
      return name
    end
  end
end

-- The widths (cm) of a table's columns, from their content, as LaTeX would: `columns` is, for each
-- column, the texts of its cells (paragraphs: lists of strings); `em` the size of the text (cm);
-- `text` the text's width (cm). A column is as wide as its longest line; if they don't fit,
-- short columns keep their width and long ones share the rest and wrap (never narrower than
-- their longest word). width "auto": as narrow as that; "full", "80%", "12cm": that wide, the
-- columns widened in proportion.
local function ems(s) -- how wide a text is, in em (roughly, in a proportional font)
  local w = 0
  for ch in s:gmatch("[^\128-\191][\128-\191]*") do
    if ch == " " then
      w = w + 0.28
    elseif ch:match("^[ilIjtfr.,;:'!|()%[%]]$") then
      w = w + 0.32
    elseif ch:match("^[A-Zmw]$") then
      w = w + 0.72
    elseif #ch > 1 then -- (accented, Greek...)
      w = w + 0.66
    else
      w = w + 0.53
    end
  end
  return w * 1.06 -- (and a little to spare)
end
function M.widths(columns, width, em, text)
  local natural, least = {}, {}
  for c, cells in ipairs(columns) do
    local longest, word = 0, 0
    for _, paras in ipairs(cells) do
      for _, para in ipairs(paras) do
        longest = math.max(longest, ems(para))
        for w in para:gmatch("%S+") do
          word = math.max(word, ems(w))
        end
      end
    end
    natural[c] = math.max(longest * em + 2 * M.PADDING, 0.8)
    least[c] = math.min(math.max(word * em + 2 * M.PADDING, 0.8), natural[c])
  end
  if #natural == 0 then
    return {}
  end
  local total = 0
  for _, w in ipairs(natural) do
    total = total + w
  end
  local percent = tonumber((width or ""):match("^([%d.]+)%%$") or "")
  local target = width == "full" and text or percent and text * percent / 100 or require("neoffice.markup").cm(width or "")
  if not target then -- auto
    if total <= text then
      return natural
    end
    target = text
  end
  if total <= target then -- (wider: each in proportion)
    return vim.tbl_map(function(w) return w * target / total end, natural)
  end
  -- narrower: the widest columns give way (wrapping their text), all to the same width, so the
  -- others keep theirs (as tabularx's X columns); never narrower than their longest word
  local function total_at(cap)
    local sum = 0
    for c, w in ipairs(natural) do
      sum = sum + math.max(least[c], math.min(w, cap))
    end
    return sum
  end
  local low, high = 0, math.max(unpack(natural))
  for _ = 1, 50 do
    local mid = (low + high) / 2
    if total_at(mid) > target then
      high = mid
    else
      low = mid
    end
  end
  local out, sum = {}, 0
  for c, w in ipairs(natural) do
    out[c] = math.max(least[c], math.min(w, low))
    sum = sum + out[c]
  end
  if sum > target * 1.001 then -- (not even their longest words fit: all narrower)
    out = vim.tbl_map(function(w) return w * target / sum end, out)
  end
  return out
end

-- ── Tables as text: aligned columns ───────────────────────────────────────

-- A table's columns are aligned (padded with spaces to the same width) when its lines fit in
-- this many columns; otherwise one space on each side of every cell.
M.ALIGN_WIDTH = 120

-- The widths of a table's columns, for aligning them: `rows` are lists of the cells' texts (as
-- written); nil when the table is too wide to align.
function M.text_widths(rows)
  local widths, columns = {}, 0
  for _, row in ipairs(rows) do
    columns = math.max(columns, #row)
    for c, text in ipairs(row) do
      widths[c] = math.max(widths[c] or 3, vim.fn.strdisplaywidth(text))
    end
  end
  local total = 1
  for c = 1, columns do
    widths[c] = widths[c] or 3
    total = total + widths[c] + 3
  end
  return total <= M.ALIGN_WIDTH and widths or nil
end

-- The spaces before and after a cell's text to fill its column (`width`: nil, not aligned).
function M.padding(text, width, align)
  local extra = width and math.max(width - vim.fn.strdisplaywidth(text), 0) or 0
  if align == "right" then
    return (" "):rep(extra), ""
  elseif align == "center" then
    local before = math.floor(extra / 2)
    return (" "):rep(before), (" "):rep(extra - before)
  end
  return "", (" "):rep(extra)
end

-- Tables of up to this many rows are kept on one page (as LaTeX's floats); longer ones break
-- across pages, their header repeated.
M.KEEP_ROWS = 20

-- Whether a cell's text is a number ("12", "-3.5", "4.3 ± 0.2", "< 0.001", "12%", "1,234"), or
-- what stands for none ("", "–", "NA"), for aligning number columns right.
local NONE = { [""] = true, ["-"] = true, ["–"] = true, ["—"] = true, NA = true, ["n/a"] = true, ["."] = true }
function M.is_number(text)
  text = vim.trim((text:gsub("[*\\]", "")))
  if NONE[text] then
    return nil -- (neither)
  end
  local num = "[%+%-−]?%d[%d.,]*%%?"
  return text:match("^[<>≤≥~]?%s*" .. num .. "$") ~= nil
    or text:match("^" .. num .. "%s*±%s*" .. num .. "$") ~= nil
    or text:match("^" .. num .. "%s*%(%s*" .. num .. "%s*%)$") ~= nil
    or text:match("^" .. num .. "%s*%[%s*" .. num .. "%s*[,;–-]%s*" .. num .. "%s*%]$") ~= nil
end

local function next_of(list, el) -- the element after el in list (the first if el is nil)
  for k, x in ipairs(list) do
    if x == el then
      return list[k + 1]
    end
  end
  return not el and list[1] or nil
end

local function insert_after(el, after, parent, before)
  if after then
    return xml.append(after.parent, el, xml.index(after) + 1)
  end
  return xml.append(before and before.parent or parent, el, before and xml.index(before) or nil)
end

-- start(fmt, ctx, plans, cell_of): plans are { cell = { table, row, col } (where in the buffer) or
-- nil, node = the paragraph's node in the document or nil }, in order. Then, for each plan in order,
-- :node(plan, prev) gives the node of a paragraph in a cell (nil for others) and :finish() ends.
function M.start(fmt, ctx, plans, cell_of)
  -- the document's table, row and cell for each of the buffer's (the first paragraph found wins)
  local of, taken = {}, {}
  local function claim(key, el)
    if not of[key] and not taken[el] then
      of[key], taken[el] = el, true
    end
  end
  for _, step in ipairs({ "table", "row", "cell" }) do
    for _, plan in ipairs(plans) do
      local c, at = plan.cell and plan.node and cell_of(plan.node), plan.cell
      if c and step == "table" then
        claim("t" .. at.table, c.tbl)
      elseif c and step == "row" and of["t" .. at.table] == c.tbl then
        claim("r" .. at.row, c.tr)
      elseif c and step == "cell" and of["r" .. at.row] == c.tr then
        claim("c" .. at.row .. ":" .. at.col, c.tc)
      end
    end
  end

  local self, state, fresh, changed_tables = {}, {}, {}, {} -- fresh: the empty paragraph of a cell made here
  local function changed()
    ctx.dirty[fmt.MAIN], state.changed = true, true
  end
  -- A copy of a row or cell with one empty paragraph in each cell, formatted like its first one.
  local function emptied(el)
    local copy = xml.copy(el)
    for _, cell in ipairs(CELLS[copy.tag] and { copy } or xml.collect(copy, CELLS)) do
      local first = xml.collect(cell, PARAGRAPHS)[1] or { tag = "text:p", attrs = {}, children = {} }
      cell.children = vim.tbl_filter(function(c) return not CONTENT[c.tag] end, cell.children)
      for _, merge in ipairs(xml.collect(cell, { ["w:vMerge"] = true })) do -- (not merged with the row above)
        xml.remove(merge)
      end
      fresh[cell] = xml.append(cell, fmt.new_para(first, ctx))
    end
    return copy
  end
  local function end_row()
    if state.tr then
      for _, tc in ipairs(xml.collect(state.tr, CELLS)) do
        if not state.cells[tc] then -- (deleted from the buffer)
          xml.remove(tc)
          changed()
        end
      end
    end
    -- (cells added to a row: not merged as the one they were copied from; empty ones beyond the
    -- table's columns, padding, go; one with text takes a column from a merged cell before it)
    if state.tr and state.new and next(state.new) then
      if fmt.fit_row then
        fmt.fit_row(ctx, state.tbl, state.tr, state.new)
      end
      for tc in pairs(state.new) do -- (the row changed if one is still in it: padding alone, it didn't)
        if tc.parent then
          changed()
          break
        end
      end
    end
    state.tr, state.tc, state.col, state.cells, state.new = nil, nil, nil, {}, {}
  end
  local function end_table()
    end_row()
    if state.tbl then
      for _, tr in ipairs(xml.collect(state.tbl, ROWS)) do
        if not state.rows[tr] then
          xml.remove(tr)
          changed()
        end
      end
      if state.changed then -- (the columns the table declares)
        fmt.fix_grid(ctx, state.tbl)
        changed_tables[#changed_tables + 1] = state.tbl
      end
      -- typeset: a new table, one whose line changed, and one typeset before that changed (its
      -- rows, cells or text); a table that wasn't only gets its rules drawn again when rows or
      -- cells were added or removed (they are drawn cell by cell)
      if fmt.set_table_format then
        local have = state.have
        if state.made or state.retyped or have.typeset and (state.changed or state.edited) then
          local options = vim.tbl_extend("force", M.defaults(), state.want)
          options.typeset = true
          fmt.set_table_format(ctx, state.tbl, options)
          ctx.dirty[fmt.MAIN] = true
        elseif state.changed and have.rules then
          fmt.set_table_format(ctx, state.tbl, { rules = have.rules })
        end
      end
    end
    state = {}
  end

  function self:node(plan, prev)
    local at = plan.cell
    if not at then
      end_table()
      return nil
    end
    if state.table ~= at.table then -- a table starts: the document's, or a new one
      end_table()
      local tbl = of["t" .. at.table]
      if not tbl then
        local place = prev and cell_of(prev)
        while place and cell_of(place.tbl) do
          place = cell_of(place.tbl)
        end
        tbl = fmt.new_table(ctx)
        for _, tc in ipairs(xml.collect(tbl, CELLS)) do
          fresh[tc] = xml.collect(tc, PARAGRAPHS)[1]
        end
        if prev then
          insert_after(tbl, place and place.tbl or prev)
        else -- (at the start)
          local body = xml.find(ctx.root, "w:body") or xml.find(ctx.root, "office:text")
          local first = vim.iter(body.children):find(function(c) return CONTENT[c.tag] or c.tag == "w:sectPr" end)
          insert_after(tbl, nil, body, first)
        end
      end
      local made = not of["t" .. at.table]
      local want = M.parse(at.format) -- (an error if its line is wrong: nothing is written)
      local have = not made and fmt.table_format and fmt.table_format(ctx, tbl) or {} -- (as it was)
      state = {
        table = at.table, tbl = tbl, rows = {}, cells = {}, made = made, want = want, have = have,
        -- its line changed: what it says isn't what it showed
        retyped = not made and M.line(vim.tbl_extend("force", M.defaults(), want)) ~= M.line(have),
      }
      if made then
        changed()
      end
    end
    state.edited = state.edited or plan.edited -- (a paragraph of it is written)
    if state.row ~= at.row then -- a row starts: the document's, the next one if it lost its own, or a new one
      local last = state.tr
      end_row()
      local tr = of["r" .. at.row]
      if not tr then
        local rows = xml.collect(state.tbl, ROWS)
        local nxt = next_of(rows, last)
        local lost = nxt and not taken[nxt] and not state.rows[nxt] -- (its paragraphs all went)
        tr = lost and nxt or insert_after(emptied(last or rows[1]), last, state.tbl, not last and rows[1] or nil)
        changed()
      end
      state.row, state.tr, state.rows[tr] = at.row, tr, true
    end
    if state.col ~= at.col then -- a cell starts (same rules)
      local last = state.tc
      local tc = of["c" .. at.row .. ":" .. at.col]
      if not tc then
        local cells = xml.collect(state.tr, CELLS)
        local nxt = next_of(cells, last)
        local lost = nxt and not taken[nxt] and not state.cells[nxt]
        if not lost then -- (the row changed: said when it ends, see end_row)
          tc = insert_after(emptied(last or cells[1] or xml.collect(state.tbl, CELLS)[1]), last, state.tr, not last and cells[1] or nil)
          state.new = state.new or {}
          state.new[tc] = { filled = false }
        else
          tc = nxt
          changed()
        end
      end
      state.col, state.tc, state.cells[tc], state.prev = at.col, tc, true, nil
    end
    if state.new and state.new[state.tc] and not plan.empty then
      state.new[state.tc].filled = true
    end
    local node = plan.node
    local c = node and cell_of(node)
    if not (c and c.tc == state.tc) then -- a new paragraph here
      node = fresh[state.tc]
      fresh[state.tc] = nil
      if not node then
        local first = xml.collect(state.tc, PARAGRAPHS)[1]
        node = insert_after(fmt.new_para(state.prev or first, ctx), state.prev, state.tc, not state.prev and first or nil)
      end
    elseif fresh[state.tc] then -- (not needed after all)
      xml.remove(fresh[state.tc])
      fresh[state.tc] = nil
    end
    state.prev = node
    return node
  end

  function self:finish()
    end_table()
    for _, tbl in ipairs(changed_tables) do
      if fmt.after_table then
        fmt.after_table(tbl)
      end
    end
  end
  return self
end

return M
