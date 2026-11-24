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
    virt_text = { { ("  %s: %s%s"):format(c.author, (c.text:gsub("\n", " ")):sub(1, 50), more), "NeofficeCommentVirt" } },
  }
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
  local opts = vim.tbl_extend("force", M.style(m.c), { id = m.id, end_row = mark[3].end_row, end_col = mark[3].end_col })
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
    api.nvim_win_set_cursor(win, { m.row + 1, m.col })
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

local function close()
  if panel.win and api.nvim_win_is_valid(panel.win) then
    api.nvim_win_close(panel.win, true)
  end
  panel.win = nil
end

function M.toggle_panel()
  if panel.win and api.nvim_win_is_valid(panel.win) then
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
  vim.wo[panel.win].signcolumn, vim.wo[panel.win].winfixwidth = "no", true
  render()
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
