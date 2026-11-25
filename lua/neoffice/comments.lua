-- Comments: highlighted ranges in the document plus a side panel to read and edit threads.
-- A comment is { author, date, text, resolved, replies = { { author, date, text }, ... } }.
local api = vim.api
local M = {}
local panel = { rows = {} } -- buf, win, doc (the document buffer), rows[line] = { mark, reply index }

local function neo()
  return require("neoffice")
end

local function now()
  return os.date("%Y-%m-%dT%H:%M:%S")
end

function M.style(c)
  local more = #c.replies > 0 and (" (+%d)"):format(#c.replies) or ""
  return {
    hl_group = c.resolved and "NeofficeCommentResolved" or "NeofficeComment",
    sign_text = "",
    sign_hl_group = "NeofficeCommentSign",
    virt_text = neo().config.comment_text -- (its text at the end of the line: if asked for)
        and { { ("  %s: %s%s"):format(c.author, vim.fn.strcharpart((c.text:gsub("\n", " ")), 0, 50), more), "NeofficeCommentVirt" } }
      or nil,
  }
end

-- :Neoffice comments text [on|off]: each comment's text at the end of its line, or not (toggles).
function M.text(arg)
  local on = arg == "on" or arg ~= "off" and not neo().config.comment_text
  neo().config.comment_text = on
  for buf, doc in pairs(neo().docs) do
    for _, m in ipairs(api.nvim_buf_get_extmarks(buf, neo().ns.comment, 0, -1, { details = true })) do
      local c = doc.comment[m[1]]
      if c then
        local opts = vim.tbl_extend("force", M.style(c), { id = m[1], end_row = m[4].end_row, end_col = m[4].end_col, right_gravity = false })
        api.nvim_buf_set_extmark(buf, neo().ns.comment, m[2], m[3], opts)
      end
    end
  end
  vim.notify("[neoffice] Comment text " .. (on and "shown" or "hidden"))
end

-- All comment marks of a document, in document order: { buf, id, row, col, end_row, end_col, c }
local function marks(buf)
  return vim.tbl_map(function(m)
    local d = m[4]
    return { buf = buf, id = m[1], row = m[2], col = m[3], end_row = d.end_row, end_col = d.end_col, c = neo().docs[buf].comment[m[1]] }
  end, api.nvim_buf_get_extmarks(buf, neo().ns.comment, 0, -1, { details = true }))
end

-- The comments changed: the buffer has something to save, and so do the comment parts.
function M.modified(buf)
  vim.bo[buf].modified = true
  neo().docs[buf].ctx.comments_dirty = true
end

-- ── Panel ───────────────────────────────────────────────────────────────────

local function render()
  if not (panel.buf and api.nvim_buf_is_valid(panel.buf) and neo().docs[panel.doc]) then
    return
  end
  local lines, rows, headers = {}, {}, {}
  local function add(text, mark, reply, header)
    lines[#lines + 1], rows[#lines + 1] = text, { mark, reply }
    headers[#lines] = header
  end
  panel.threads = {} -- (the first line of each comment)
  for _, m in ipairs(marks(panel.doc)) do
    local c = m.c
    panel.threads[#panel.threads + 1] = #lines + 1
    add(("%s%s · %s"):format(c.resolved and "✓ " or "", c.author, c.date:sub(1, 10)), m, nil, true)
    for _, l in ipairs(vim.split(c.text, "\n")) do
      add("  " .. l, m)
    end
    for k, r in ipairs(c.replies) do
      add(("  ↳ %s · %s"):format(r.author, r.date:sub(1, 10)), m, k, true)
      for _, l in ipairs(vim.split(r.text, "\n")) do
        add("    " .. l, m, k)
      end
    end
    add("", m)
  end
  if #lines == 0 then
    lines = { "No comments yet." }
  end
  local help = {}
  for _, action in ipairs({ "reply", "resolve", "delete", "jump", "next", "prev", "close" }) do
    local key = neo().config.panel_keys[action]
    help[#help + 1] = key and key .. " " .. action or nil
  end
  table.insert(lines, table.concat(help, " · "))
  vim.bo[panel.buf].modifiable = true
  api.nvim_buf_set_lines(panel.buf, 0, -1, false, lines)
  vim.bo[panel.buf].modifiable = false
  api.nvim_buf_clear_namespace(panel.buf, neo().ns.comment, 0, -1)
  for row in pairs(headers) do
    api.nvim_buf_set_extmark(panel.buf, neo().ns.comment, row - 1, 0, { end_col = #lines[row], hl_group = "NeofficeCommentAuthor" })
  end
  api.nvim_buf_set_extmark(panel.buf, neo().ns.comment, #lines - 1, 0, { end_col = #lines[#lines], hl_group = "Comment" })
  panel.rows = rows
end

-- The comment (and reply index) under the cursor: in the panel, or in the document (the comment
-- whose range holds the cursor, else the first one on the cursor line).
local function at_cursor()
  local buf, cursor = api.nvim_get_current_buf(), api.nvim_win_get_cursor(0)
  if buf == panel.buf then
    local row = panel.rows[cursor[1]]
    return row and row[1], row and row[2]
  elseif not neo().docs[buf] then
    return vim.notify("[neoffice] Not a document buffer", vim.log.levels.WARN)
  end
  local pos, on_line = { cursor[1] - 1, cursor[2] }, nil
  for _, m in ipairs(marks(buf)) do
    local inside = (m.row < pos[1] or m.row == pos[1] and m.col <= pos[2]) and (m.end_row > pos[1] or m.end_row == pos[1] and m.end_col >= pos[2])
    if inside then
      return m
    end
    on_line = on_line or (m.row == pos[1] and m)
  end
  if not on_line then
    vim.notify("[neoffice] No comment here")
  end
  return on_line or nil
end

-- A comment changed: update its mark, flag the document as modified, redraw the panel.
local function changed(m)
  local mark = api.nvim_buf_get_extmark_by_id(m.buf, neo().ns.comment, m.id, { details = true })
  local opts = vim.tbl_extend("force", M.style(m.c), { id = m.id, end_row = mark[3].end_row, end_col = mark[3].end_col, right_gravity = false })
  api.nvim_buf_set_extmark(m.buf, neo().ns.comment, mark[1], mark[2], opts)
  M.modified(m.buf)
  render()
end

function M.reply()
  local m = at_cursor()
  if not m then
    return
  end
  vim.ui.input({ prompt = "Reply to " .. m.c.author .. ": " }, function(text)
    if text and text ~= "" then
      table.insert(m.c.replies, { author = neo().config.author, date = now(), text = text, replies = {} })
      changed(m)
    end
  end)
end

function M.resolve()
  local m = at_cursor()
  if m then
    m.c.resolved = not m.c.resolved
    changed(m)
  end
end

function M.delete()
  local m, reply = at_cursor()
  if not m or vim.fn.confirm(reply and "Delete reply?" or "Delete comment and its replies?", "&Yes\n&No") ~= 1 then
    return
  end
  if reply then
    table.remove(m.c.replies, reply)
    return changed(m)
  end
  api.nvim_buf_del_extmark(m.buf, neo().ns.comment, m.id)
  neo().docs[m.buf].comment[m.id] = nil
  M.modified(m.buf)
  render()
end

function M.jump()
  local m = at_cursor()
  local win = m and vim.fn.bufwinid(m.buf)
  if win and win ~= -1 then
    api.nvim_set_current_win(win)
    local pos = api.nvim_buf_get_extmark_by_id(m.buf, neo().ns.comment, m.id, {}) -- (where it is now)
    api.nvim_win_set_cursor(win, { (pos[1] or m.row) + 1, pos[2] or m.col })
  end
end

-- To the next (dir 1) or previous (-1) comment in the panel.
local function to_comment(dir)
  local row = api.nvim_win_get_cursor(0)[1]
  local threads = panel.threads or {}
  for k = dir > 0 and 1 or #threads, dir > 0 and #threads or 1, dir do
    if dir * (threads[k] - row) > 0 then
      return api.nvim_win_set_cursor(0, { threads[k], 0 })
    end
  end
end

local function shown() -- (the panel is open: its window shows it)
  return panel.win and api.nvim_win_is_valid(panel.win) and api.nvim_win_get_buf(panel.win) == panel.buf
end

-- The text of the comment the panel is on, highlighted in the document (NeofficeCommentCurrent).
local NS_CURRENT = api.nvim_create_namespace("neoffice_comment_current")
local function show_current()
  if panel.doc and api.nvim_buf_is_valid(panel.doc) then
    api.nvim_buf_clear_namespace(panel.doc, NS_CURRENT, 0, -1)
  end
  local entry = shown() and panel.rows and panel.rows[api.nvim_win_get_cursor(panel.win)[1]]
  local pos = entry and entry[1] and api.nvim_buf_get_extmark_by_id(panel.doc, neo().ns.comment, entry[1].id, { details = true })
  if not (pos and pos[1]) then
    return
  end
  local row, col, end_row, end_col = pos[1], pos[2], pos[3].end_row, pos[3].end_col
  if row == end_row and col == end_col then -- (a comment on a point: the character there)
    local line = api.nvim_buf_get_lines(panel.doc, row, row + 1, false)[1] or ""
    col = math.max(math.min(col, #line - 1), 0)
    end_col = math.min(col + #(line:match("^.[\128-\191]*", col + 1) or " "), #line)
  end
  pcall(api.nvim_buf_set_extmark, panel.doc, NS_CURRENT, row, col, {
    end_row = end_row, end_col = end_col, hl_group = "NeofficeCommentCurrent", priority = 300,
  })
  return row, col
end

-- Moving in the panel: the document scrolls to the comment's text, if it isn't in view.
local function reveal()
  local row, col = show_current()
  local win = row and vim.fn.bufwinid(panel.doc)
  if win and win ~= -1 then
    api.nvim_win_call(win, function()
      if row + 1 < vim.fn.line("w0") or row + 1 > vim.fn.line("w$") then
        api.nvim_win_set_cursor(win, { row + 1, col })
        vim.cmd("normal! zz")
      end
    end)
  end
end

local function close()
  if shown() and #api.nvim_list_wins() > 1 then
    api.nvim_win_close(panel.win, true)
  end
  if panel.doc and api.nvim_buf_is_valid(panel.doc) then
    api.nvim_buf_clear_namespace(panel.doc, NS_CURRENT, 0, -1)
  end
  panel.win = nil
  pcall(api.nvim_del_augroup_by_name, "neoffice_panel")
end

-- The comment at a position of the document, or else the closest to it (in lines, then columns):
-- its number, in document order (as in the panel).
local function nearest(buf, row, col)
  local best, best_distance
  for k, m in ipairs(marks(buf)) do
    local before = m.end_row < row or m.end_row == row and m.end_col < col -- (it ends before the position)
    local after = m.row > row or m.row == row and m.col > col
    local r, c = before and m.end_row or after and m.row or row, before and m.end_col or after and m.col or col
    local distance = math.abs(r - row) * 100000 + math.abs(c - col)
    if not best_distance or distance < best_distance then
      best, best_distance = k, distance
    end
  end
  return best
end

-- The panel shows the comment nearest to the cursor in the document.
local function follow()
  local win = vim.fn.bufwinid(panel.doc or -1)
  if not (shown() and win ~= -1) then
    return
  end
  local pos = api.nvim_win_get_cursor(win)
  local k = nearest(panel.doc, pos[1] - 1, pos[2])
  local row = k and panel.threads and panel.threads[k]
  if row then
    api.nvim_win_set_cursor(panel.win, { row, 0 })
  end
  show_current()
end

function M.toggle_panel()
  if shown() then
    return close()
  end
  local buf = api.nvim_get_current_buf()
  if not neo().docs[buf] then
    return vim.notify("[neoffice] Not a document buffer", vim.log.levels.WARN)
  end
  panel.doc = buf
  if not (panel.buf and api.nvim_buf_is_valid(panel.buf)) then
    panel.buf = api.nvim_create_buf(false, true)
    vim.bo[panel.buf].filetype = "neoffice-comments"
    -- keys of the panel only (it is neoffice's own buffer); set panel_keys.<action> = false to drop one
    local actions = {
      reply = M.reply, resolve = M.resolve, delete = M.delete, jump = M.jump, close = close,
      next = function() to_comment(1) end, prev = function() to_comment(-1) end,
    }
    for action, key in pairs(neo().config.panel_keys) do
      if key and actions[action] then
        vim.keymap.set("n", key, actions[action], { buffer = panel.buf, nowait = true, desc = action })
      end
    end
  end
  vim.cmd("botright 50vsplit")
  panel.win = api.nvim_get_current_win()
  api.nvim_win_set_buf(panel.win, panel.buf)
  vim.wo[panel.win].number, vim.wo[panel.win].relativenumber, vim.wo[panel.win].wrap = false, false, true
  vim.wo[panel.win].signcolumn, vim.wo[panel.win].winfixwidth, vim.wo[panel.win].cursorline = "no", true, true
  render()
  follow() -- (on the comment where you are; and as you move in the document, or in the panel)
  local group = api.nvim_create_augroup("neoffice_panel", {})
  api.nvim_create_autocmd("CursorMoved", { group = group, buffer = buf, callback = follow })
  api.nvim_create_autocmd("CursorMoved", { group = group, buffer = panel.buf, callback = reveal })
  api.nvim_create_autocmd("BufWinLeave", { group = group, buffer = panel.buf, callback = function()
    if panel.doc and api.nvim_buf_is_valid(panel.doc) then
      api.nvim_buf_clear_namespace(panel.doc, NS_CURRENT, 0, -1)
    end
  end })
end

-- ── Adding ──────────────────────────────────────────────────────────────────

-- Comment the visual selection, or the cursor position in normal mode.
-- `use_marks` takes the last selection instead (for :'<,'>Neoffice comment).
function M.add(use_marks)
  local buf = api.nvim_get_current_buf()
  if not neo().docs[buf] then
    return vim.notify("[neoffice] Not a document buffer", vim.log.levels.WARN)
  end
  local mode = use_marks == true and vim.fn.visualmode() or vim.fn.mode()
  local r1, c1, r2, c2
  if mode:match("^[vV\22]") then
    local s, e = vim.fn.getpos(use_marks == true and "'<" or "v"), vim.fn.getpos(use_marks == true and "'>" or ".")
    if s[2] > e[2] or (s[2] == e[2] and s[3] > e[3]) then
      s, e = e, s
    end
    local last_line = vim.fn.getline(e[2])
    r1, c1, r2 = s[2] - 1, s[3] - 1, e[2] - 1
    c2 = math.min(e[3] - 1 + #vim.fn.matchstr(last_line, ".", e[3] - 1), #last_line)
    if mode == "V" then
      c1, c2 = 0, #last_line
    end
    api.nvim_feedkeys(vim.keycode("<Esc>"), "nx", false)
  else
    local cursor = api.nvim_win_get_cursor(0)
    r1, c1, r2, c2 = cursor[1] - 1, cursor[2], cursor[1] - 1, cursor[2]
  end
  vim.ui.input({ prompt = "Comment: " }, function(text)
    if text and text ~= "" then
      neo().place(buf, "comment", { author = neo().config.author, date = now(), text = text, resolved = false, replies = {} }, r1, c1, r2, c2)
      M.modified(buf)
      render()
    end
  end)
end

return M
