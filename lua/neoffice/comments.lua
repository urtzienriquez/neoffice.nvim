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
  vim.notify("neoffice: Comment text " .. (on and "shown" or "hidden"))
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

-- Its keys (keymaps.comments: two keys a family, as zotero.nvim's and fugitive's), in the order
-- shown at its end: { action, what it does }.
M.KEYS = {
  { "go", "go to it" }, { "reply", "reply" }, { "edit", "edit" }, { "resolve", "resolve" }, { "delete", "delete" },
  { "next", "next" }, { "prev", "previous" }, { "text", "text in lines" }, { "undo", "undo" }, { "redo", "redo" },
  { "close", "close" }, { "help", "help" },
}
function M.keys()
  local km = neo().config.keymaps
  return type(km) == "table" and type(km.comments) == "table" and km.comments or {}
end

local function render()
  if not (panel.buf and api.nvim_buf_is_valid(panel.buf) and neo().docs[panel.doc]) then
    return
  end
  -- as fugitive's status window: what it is about, its help, then its sections (highlighted by
  -- syntax/neoffice-comments.vim)
  local lines, rows = {}, {}
  local function add(text, mark, reply)
    lines[#lines + 1], rows[#lines + 1] = text, { mark, reply }
  end
  local help = M.keys().help
  help = type(help) == "table" and help[1] or help
  add("Document: " .. vim.fn.fnamemodify(api.nvim_buf_get_name(panel.doc), ":t"))
  if help then
    add("Help: " .. help)
  end
  local all = marks(panel.doc)
  panel.threads, panel.order = {}, {} -- (each comment's first line, by its place in the document; and as listed)
  for _, resolved in ipairs({ false, true }) do
    local these = vim.tbl_filter(function(m) return (m.c.resolved or false) == resolved end, all)
    if #these > 0 then
      add("")
      add(("%s (%d)"):format(resolved and "Resolved" or "Open", #these))
      for _, m in ipairs(these) do
        local c = m.c
        for k, x in ipairs(all) do
          if x == m then
            panel.threads[k] = #lines + 1
          end
        end
        panel.order[#panel.order + 1] = #lines + 1
        add(("%s, %s"):format(c.author, c.date:sub(1, 10)), m)
        for _, l in ipairs(vim.split(c.text, "\n")) do
          add("    " .. l, m)
        end
        for k, r in ipairs(c.replies) do
          add(("  %s, %s"):format(r.author, r.date:sub(1, 10)), m, k)
          for _, l in ipairs(vim.split(r.text, "\n")) do
            add("    " .. l, m, k)
          end
        end
      end
    end
  end
  if #all == 0 then
    add("")
    add("No comments. :Neoffice comment adds one.")
  end
  vim.bo[panel.buf].modifiable = true
  api.nvim_buf_set_lines(panel.buf, 0, -1, false, lines)
  vim.bo[panel.buf].modifiable = false
  vim.bo[panel.buf].modified = false -- (drawn, not changed: nothing of its own to save, :q closes it)
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
    return vim.notify("neoffice: Not a document buffer", vim.log.levels.WARN)
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
    vim.notify("neoffice: No comment here")
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


-- ── Undo ──────────────────────────────────────────────────────────────────
-- What you did to a document's comments (they aren't its text: its own undo doesn't reach them):
-- u and <C-r> in the panel undo and redo it. Each step { what, undo, redo }.
local history = {}
local function record(buf, step)
  history[buf] = history[buf] or { done = {}, undone = {} }
  table.insert(history[buf].done, step)
  history[buf].undone = {}
end
local function mark_of(buf, c) -- (a comment's mark, by its data: marks are made again when restored)
  for id, x in pairs(neo().docs[buf].comment) do
    if x == c then
      return id
    end
  end
end
local function refresh(buf, c) -- (after a step: its mark's look, the document to save, the panel)
  local id = c and mark_of(buf, c)
  if id then
    changed({ buf = buf, id = id, c = c })
  else
    M.modified(buf)
    render()
  end
end
-- A comment taken out of the document, and put back where it was (both: as steps to undo).
local function take_out(buf, c)
  local id = mark_of(buf, c)
  local pos = id and api.nvim_buf_get_extmark_by_id(buf, neo().ns.comment, id, { details = true })
  if not (pos and pos[1]) then
    return nil
  end
  api.nvim_buf_del_extmark(buf, neo().ns.comment, id)
  neo().docs[buf].comment[id] = nil
  refresh(buf, nil)
  return { pos[1], pos[2], pos[3].end_row, pos[3].end_col }
end
local function put_back(buf, c, where)
  local line_count = api.nvim_buf_line_count(buf)
  local r1, r2 = math.min(where[1], line_count - 1), math.min(where[3], line_count - 1)
  local function col(r, c0) return math.min(c0, #(api.nvim_buf_get_lines(buf, r, r + 1, false)[1] or "")) end
  neo().place(buf, "comment", c, r1, col(r1, where[2]), r2, col(r2, where[4]))
  refresh(buf, c)
end
-- Where its comments and its text were when it was opened or saved: undone back there (and nothing
-- else changed meanwhile: heading numbers, changes reviewed), the document is as saved again.
function M.reset(buf)
  history[buf] = nil
  M.saved(buf)
end
function M.saved(buf)
  local doc = neo().docs[buf]
  if doc then
    local h = history[buf]
    doc.comments_saved = h and h.done[#h.done] or false
    doc.text_saved = vim.fn.undotree(buf).seq_cur
  end
end
local function settle(buf)
  local doc = neo().docs[buf]
  local h = history[buf]
  if not doc or doc.unsaved_other then
    return
  end
  local at = h and h.done[#h.done] or false
  if at == doc.comments_saved and vim.fn.undotree(buf).seq_cur == doc.text_saved then
    vim.bo[buf].modified = false
  end
end

local function step(dir)
  local buf = panel.doc and api.nvim_buf_is_valid(panel.doc) and panel.doc or api.nvim_get_current_buf()
  local h = history[buf]
  local s = h and table.remove(dir > 0 and h.undone or h.done)
  if not s then
    return vim.notify(dir > 0 and "neoffice: Nothing to redo" or "neoffice: Nothing to undo")
  end
  if dir > 0 then
    s.redo()
    table.insert(h.done, s)
  else
    s.undo()
    table.insert(h.undone, s)
  end
  settle(buf)
  vim.notify(("neoffice: %s: %s"):format(dir > 0 and "Redone" or "Undone", s.what))
end
function M.undo() step(-1) end
function M.redo() step(1) end

-- A buffer to write a comment in (a reply, a new one, or one's text edited), as fugitive's commit
-- message: a split across the whole width, the text on top and what to do in "#" lines under it
-- (left out of what is written). :w writes it into the document (the next :w edits what it wrote),
-- :wq writes it and closes, :q! leaves it. spec = { title, about (lines: what it answers), lines,
-- write(text) }.
local composed = 0
local function compose(spec)
  composed = composed + 1
  local from = api.nvim_get_current_win()
  vim.cmd((vim.o.splitbelow and "botright" or "topleft") .. " split")
  local win = api.nvim_get_current_win()
  local buf = api.nvim_create_buf(true, false)
  api.nvim_win_set_buf(win, buf)
  pcall(api.nvim_buf_set_name, buf, ("neoffice://comment/%d"):format(composed))
  vim.bo[buf].buftype, vim.bo[buf].bufhidden, vim.bo[buf].swapfile = "acwrite", "wipe", false
  vim.bo[buf].filetype, vim.bo[buf].commentstring, vim.bo[buf].textwidth = "neoffice-comment", "# %s", 0
  local help = { "", "# " .. spec.title }
  for _, l in ipairs(spec.about or {}) do
    help[#help + 1] = "#   " .. l
  end
  vim.list_extend(help, {
    "#",
    "# Lines starting with '#' are left out. :wq (or :w, then :q) adds it to",
    "# the document; then save the document to keep it. An empty text adds",
    "# nothing (delete it and :wq to change your mind), nor does :q! before :w.",
  })
  local undolevels = vim.bo[buf].undolevels -- (what it opens with: not to be undone, only your changes)
  vim.bo[buf].undolevels = -1
  api.nvim_buf_set_lines(buf, 0, -1, false, vim.list_extend(vim.deepcopy(spec.lines), help))
  vim.bo[buf].undolevels = undolevels
  vim.bo[buf].modified = false
  api.nvim_win_set_cursor(win, { 1, 0 })
  -- (as git takes a commit's message: what :w wrote, once the buffer closes; :q! with nothing
  -- written leaves it)
  local written
  api.nvim_create_autocmd("BufWipeout", { buffer = buf, once = true, callback = function()
    vim.schedule(function()
      if api.nvim_win_is_valid(from) then -- (back where you were)
        api.nvim_set_current_win(from)
      end
      if written then
        local ok, err = pcall(spec.write, written)
        if not ok then
          vim.notify("neoffice: " .. tostring(err), vim.log.levels.ERROR)
        end
      end
    end)
  end })
  api.nvim_create_autocmd("BufWriteCmd", { buffer = buf, callback = function()
    local kept = vim.tbl_filter(function(l) return not l:match("^#") end, api.nvim_buf_get_lines(buf, 0, -1, false))
    local text = table.concat(kept, "\n"):gsub("^%s+", ""):gsub("%s+$", "")
    if text == "" then -- (as git: an empty message aborts, also one written before)
      written = nil
      vim.bo[buf].modified = false
      return vim.notify("neoffice: The text is empty: nothing will be added")
    end
    written = text
    vim.bo[buf].modified = false
  end })
end

-- Reply to the comment under the cursor: in a buffer (see compose), or `text` at once (as tests do).
function M.reply(text)
  neo().finish_all()
  local m = at_cursor()
  if not m then
    return vim.notify("neoffice: No comment here", vim.log.levels.WARN)
  end
  local function write(t)
    local c, reply = m.c, { author = neo().config.author, date = now(), text = t, replies = {} }
    table.insert(c.replies, reply)
    changed(m)
    record(m.buf, {
      what = "a reply",
      undo = function()
        for k, r in ipairs(c.replies) do
          if r == reply then
            table.remove(c.replies, k)
          end
        end
        refresh(m.buf, c)
      end,
      redo = function()
        table.insert(c.replies, reply)
        refresh(m.buf, c)
      end,
    })
  end
  if text then
    return text ~= "" and write(text) or nil
  end
  local about = { ("%s, %s:"):format(m.c.author, m.c.date:sub(1, 10)) }
  for _, l in ipairs(vim.split(m.c.text, "\n")) do
    about[#about + 1] = "  " .. l
  end
  compose({ title = "Replying to:", about = about, lines = { "" }, write = write })
end

-- Edit the text of the comment (or reply) under the cursor, in a buffer (see compose).
function M.edit()
  neo().finish_all()
  local m, reply = at_cursor()
  if not m then
    return vim.notify("neoffice: No comment here", vim.log.levels.WARN)
  end
  local x = reply and m.c.replies[reply] or m.c
  compose({ title = ("Editing %s's %s of %s."):format(x.author, reply and "reply" or "comment", x.date:sub(1, 10)), lines = vim.split(x.text, "\n"),
    write = function(t)
      local old = x.text
      x.text = t
      changed(m)
      record(m.buf, {
        what = "an edit",
        undo = function() x.text = old; refresh(m.buf, m.c) end,
        redo = function() x.text = t; refresh(m.buf, m.c) end,
      })
    end })
end

function M.resolve()
  neo().finish_all()
  local m = at_cursor()
  if m then
    local c = m.c
    c.resolved = not c.resolved
    changed(m)
    local function toggle() c.resolved = not c.resolved; refresh(m.buf, c) end
    record(m.buf, { what = c.resolved and "resolved" or "reopened", undo = toggle, redo = toggle })
  end
end

function M.delete()
  neo().finish_all()
  local m, reply = at_cursor()
  if not m or vim.fn.confirm(reply and "Delete reply?" or "Delete comment and its replies?", "&Yes\n&No") ~= 1 then
    return
  end
  local c, buf = m.c, m.buf
  if reply then
    local r = table.remove(c.replies, reply)
    changed(m)
    return record(buf, {
      what = "a reply deleted",
      undo = function() table.insert(c.replies, reply, r); refresh(buf, c) end,
      redo = function()
        for k, x in ipairs(c.replies) do
          if x == r then
            table.remove(c.replies, k)
          end
        end
        refresh(buf, c)
      end,
    })
  end
  local where = take_out(buf, c)
  record(buf, {
    what = "a comment deleted",
    undo = function() put_back(buf, c, where) end,
    redo = function() where = take_out(buf, c) or where end,
  })
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
  local threads = panel.order or {}
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
    return vim.notify("neoffice: Not a document buffer", vim.log.levels.WARN)
  end
  panel.doc = buf
  if not (panel.buf and api.nvim_buf_is_valid(panel.buf)) then
    panel.buf = api.nvim_create_buf(false, true)
    panel.keys = nil -- (a new buffer: none of neoffice's keys in it yet)
    vim.bo[panel.buf].filetype = "neoffice-comments"
    -- :w in it saves the document (its comments are the document's): through neoffice's save, called
    -- here (a :write of the document from this autocommand wouldn't run neoffice's own, and Vim would
    -- write its text over the file)
    api.nvim_buf_set_name(panel.buf, "neoffice://comments")
    vim.bo[panel.buf].buftype = "acwrite"
    api.nvim_create_autocmd("BufWriteCmd", { buffer = panel.buf, callback = function()
      local doc = panel.doc
      if doc and api.nvim_buf_is_valid(doc) and neo().docs[doc] then
        neo().save(doc, api.nvim_buf_get_name(doc))
      end
      vim.bo[panel.buf].modified = false
    end })
  end
  do -- keys of the panel only (it is neoffice's own buffer): keymaps.comments (false drops one); set
    -- each time it opens, as they are configured now
    -- (only neoffice's own, as set when it last opened: yours, from your filetype autocommands
    -- for neoffice-comments, stay)
    for _, lhs in ipairs(panel.keys or {}) do
      pcall(vim.keymap.del, "n", lhs, { buffer = panel.buf })
    end
    panel.keys = {}
    local actions = {
      go = M.jump, reply = M.reply, edit = M.edit, resolve = M.resolve, delete = M.delete, close = close, undo = M.undo, redo = M.redo,
      next = function() to_comment(1) end, prev = function() to_comment(-1) end,
      text = function() M.text() end, help = function() neo().help("neoffice-comments-panel") end,
    }
    for _, k in ipairs(M.KEYS) do
      local keys = M.keys()[k[1]]
      for _, key in ipairs(type(keys) == "table" and keys or { keys }) do
        if key and actions[k[1]] then
          vim.keymap.set("n", key, actions[k[1]], { buffer = panel.buf, nowait = true, silent = true, desc = "neoffice: " .. k[2] })
          table.insert(panel.keys, key)
        end
      end
    end
  end
  vim.cmd("botright 50vsplit")
  panel.win = api.nvim_get_current_win()
  api.nvim_win_set_buf(panel.win, panel.buf)
  -- (as :setlocal, for its window only: not inherited by the windows opened from it)
  for name, value in pairs({ number = false, relativenumber = false, wrap = true, signcolumn = "no", winfixwidth = true, cursorline = true,
    foldcolumn = "0", list = false, spell = false, statuscolumn = "", linebreak = true, breakindent = true, winbar = "" }) do
    api.nvim_set_option_value(name, value, { win = panel.win, scope = "local" })
  end
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
local NS_PENDING = api.nvim_create_namespace("neoffice_comment_pending") -- (where a comment being written goes)
function M.add(use_marks, text)
  neo().finish_all()
  local buf = api.nvim_get_current_buf()
  if not neo().docs[buf] then
    return vim.notify("neoffice: Not a document buffer", vim.log.levels.WARN)
  end
  local mode = use_marks == true and vim.fn.visualmode() or vim.fn.mode()
  local r1, c1, r2, c2, pending
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
  local c
  local function write(t)
    if c then -- (written already: edited)
      c.text = t
    else
      local pos = pending and api.nvim_buf_get_extmark_by_id(buf, NS_PENDING, pending, { details = true })
      if pos and pos[1] then -- (where its text is now: the document may have changed meanwhile)
        r1, c1, r2, c2 = pos[1], pos[2], pos[3].end_row, pos[3].end_col
        api.nvim_buf_del_extmark(buf, NS_PENDING, pending)
      end
      c = { author = neo().config.author, date = now(), text = t, resolved = false, replies = {} }
      neo().place(buf, "comment", c, r1, c1, r2, c2)
      local where
      record(buf, {
        what = "a comment",
        undo = function() where = take_out(buf, c) or where end,
        redo = function() put_back(buf, c, where or { r1, c1, r2, c2 }) end,
      })
    end
    M.modified(buf)
    render()
  end
  if text then
    return text ~= "" and write(text) or nil
  end
  pending = api.nvim_buf_set_extmark(buf, NS_PENDING, r1, c1, { end_row = r2, end_col = c2, right_gravity = false, end_right_gravity = true })
  local on = api.nvim_buf_get_text(buf, r1, c1, r2, c2, {})
  local quoted = vim.trim(table.concat(on, " "):gsub("%s+", " "))
  compose({ title = quoted ~= "" and "A comment on:" or "A comment here.", about = quoted ~= "" and { "  " .. quoted } or nil, lines = { "" }, write = write })
end

return M
