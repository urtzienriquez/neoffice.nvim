-- Tables on save. The tables, rows and cells of the buffer are matched with the document's
-- through the paragraphs in them (a row is the document's row its paragraphs came from). Rows and
-- cells typed in the buffer are added (copies of their neighbours, emptied); those deleted from the
-- buffer are removed. Each paragraph in a cell is then written as any other paragraph.
local xml = require("neoffice.xml")
local M = {}

local PARAGRAPHS = { ["w:p"] = true, ["text:p"] = true, ["text:h"] = true }
local CONTENT = { ["w:p"] = true, ["w:tbl"] = true, ["w:sdt"] = true, ["text:p"] = true, ["text:h"] = true, ["text:list"] = true, ["table:table"] = true }
local CELLS = { ["w:tc"] = true, ["table:table-cell"] = true }
local ROWS = { ["w:tr"] = true, ["table:table-row"] = true }

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

  local self, state, fresh = {}, {}, {} -- fresh: the empty paragraph of a cell made here
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
      fresh[cell] = xml.append(cell, fmt.new_para(first))
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
    state.tr, state.tc, state.cells = nil, nil, {}
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
      state = { table = at.table, tbl = tbl, rows = {}, cells = {} }
      if not of["t" .. at.table] then
        changed()
      end
    end
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
        tc = lost and nxt or insert_after(emptied(last or cells[1] or xml.collect(state.tbl, CELLS)[1]), last, state.tr, not last and cells[1] or nil)
        changed()
      end
      state.col, state.tc, state.cells[tc], state.prev = at.col, tc, true, nil
    end
    local node = plan.node
    local c = node and cell_of(node)
    if not (c and c.tc == state.tc) then -- a new paragraph here
      node = fresh[state.tc]
      fresh[state.tc] = nil
      if not node then
        local first = xml.collect(state.tc, PARAGRAPHS)[1]
        node = insert_after(fmt.new_para(state.prev or first), state.prev, state.tc, not state.prev and first or nil)
      end
    elseif fresh[state.tc] then -- (not needed after all)
      xml.remove(fresh[state.tc])
      fresh[state.tc] = nil
    end
    state.prev = node
    return node
  end

  self.finish = end_table
  return self
end

return M
