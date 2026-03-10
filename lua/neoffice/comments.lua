-- neoffice/comments.lua
-- Vertical split panel: threaded comments with reply / resolve / delete.
-- Write-back to .docx is done via extractor.write_comments_docx.

local extractor = require("neoffice.extractor")
local config = require("neoffice.config")
local M = {}

-- ── State ────────────────────────────────────────────────────────────────────

local state = {
  panel_buf = nil, -- comments panel buffer
  panel_win = nil, -- comments panel window
  main_buf = nil, -- proxy text buffer
  orig_path = nil,
  comments = {},
  comment_line_map = {}, -- 0-indexed line → comment id in panel buffer
}

local NS_ANCHORS = vim.api.nvim_create_namespace("neoffice_comment_anchors")

-- ── Highlights ───────────────────────────────────────────────────────────────

function M.setup_highlights()
  vim.api.nvim_set_hl(0, "NeofficeCommentPanel", { link = "Normal", default = true })
  vim.api.nvim_set_hl(0, "NeofficeCommentCursor", { bg = "#24283b", default = true })
  vim.api.nvim_set_hl(0, "NeofficeCommentSign", { fg = "#e0af68", default = true })
  vim.api.nvim_set_hl(0, "NeofficeCommentVirt", { fg = "#565f89", italic = true, default = true })
  vim.api.nvim_set_hl(0, "NeofficeCommentAuthor", { fg = "#7aa2f7", bold = true, default = true })
  vim.api.nvim_set_hl(0, "NeofficeCommentDate", { fg = "#565f89", default = true })
  vim.api.nvim_set_hl(0, "NeofficeCommentResolved", { fg = "#9ece6a", default = true })
end

local NS_PANEL = vim.api.nvim_create_namespace("neoffice_panel_hls")

-- ── Rendering ────────────────────────────────────────────────────────────────

local function wrap(text, width, indent)
  local out = {}
  local prefix = string.rep(" ", indent)
  local wrap_width = width - indent
  for _, line in ipairs(vim.split(text, "\n", { plain = true })) do
    while #line > wrap_width do
      local segment = line:sub(1, wrap_width + 1)
      local last_space = segment:find("%s[^%s]*$")
      local split_at = last_space or (wrap_width + 1)
      table.insert(out, prefix .. line:sub(1, split_at - 1))
      line = line:sub(split_at + 1)
    end
    if line ~= "" then
      table.insert(out, prefix .. line)
    end
  end
  return out
end

local function render()
  if not state.panel_buf or not vim.api.nvim_buf_is_valid(state.panel_buf) then
    return
  end

  local lines = {}
  local line_map = {}
  local hl_data = {}
  local W = 52

  -- Helper to queue an extmark
  local function add_extmark(group, line, start_col, end_col)
    table.insert(hl_data, {
      line = line,
      s = start_col,
      e = end_col,
      hl = group,
    })
  end

  if #state.comments == 0 then
    lines = { "", "  (no comments)", "" }
  else
    for _, cm in ipairs(state.comments) do
      line_map[#lines] = cm.id

      local author_str = "@" .. cm.author
      local date_str = (cm.date or ""):sub(1, 10)
      local resolved_tag = cm.resolved and "  ✓" or ""

      local header_prefix = "┌─ "
      local author_start = #header_prefix
      local author_end = author_start + #author_str

      local date_gap = "  "
      local date_start = author_end + #date_gap
      local date_end = date_start + #date_str

      add_extmark("NeofficeCommentAuthor", #lines, author_start, author_end)
      add_extmark("NeofficeCommentDate", #lines, date_start, date_end)

      if cm.resolved then
        add_extmark("NeofficeCommentResolved", #lines, date_end, date_end + #resolved_tag)
      end

      table.insert(lines, string.format("%s%s%s%s%s", header_prefix, author_str, date_gap, date_str, resolved_tag))

      for _, l in ipairs(wrap(cm.text or "", W, 3)) do
        table.insert(lines, "│" .. l)
      end

      for _, r in ipairs(cm.replies or {}) do
        table.insert(lines, "│")

        local r_author = "@" .. (r.author or "?")
        local r_date = (r.date or ""):sub(1, 10)
        local reply_prefix = "│  ↩ " -- Correctly handles multi-byte UTF-8

        local ra_start = #reply_prefix
        local ra_end = ra_start + #r_author
        local rd_gap = "  "
        local rd_start = ra_end + #rd_gap
        local rd_end = rd_start + #r_date

        add_extmark("NeofficeCommentAuthor", #lines, ra_start, ra_end)
        add_extmark("NeofficeCommentDate", #lines, rd_start, rd_end)

        table.insert(lines, string.format("%s%s%s%s", reply_prefix, r_author, rd_gap, r_date))

        for _, l in ipairs(wrap(r.text or "", W - 4, 5)) do
          table.insert(lines, "│" .. l)
        end
      end

      table.insert(lines, "└" .. string.rep("─", 10))
      table.insert(lines, "")
    end

    local footer_text = "  r=reply  <CR>=resolve  d=delete  q=close"
    table.insert(lines, footer_text)
    local footer_lnum = #lines - 1

    add_extmark("NeofficeCommentDate", footer_lnum, 0, #footer_text)

    for _, key in ipairs({ "r", "<CR>", "d", "q" }) do
      local s, e = footer_text:find(key .. "=", 1, true)
      if s and e then
        add_extmark("NeofficeCommentResolved", footer_lnum, s - 1, e - 1)
      end
    end
  end

  vim.api.nvim_set_option_value("modifiable", true, { buf = state.panel_buf })
  vim.api.nvim_buf_set_lines(state.panel_buf, 0, -1, false, lines)

  vim.api.nvim_buf_clear_namespace(state.panel_buf, NS_PANEL, 0, -1)
  for _, mark in ipairs(hl_data) do
    vim.api.nvim_buf_set_extmark(state.panel_buf, NS_PANEL, mark.line, mark.s, {
      end_col = mark.e,
      hl_group = mark.hl,
    })
  end

  vim.api.nvim_set_option_value("modifiable", false, { buf = state.panel_buf })
  state.comment_line_map = line_map
end

-- ── Load ─────────────────────────────────────────────────────────────────────

function M.load(orig_path)
  state.orig_path = orig_path
  local data = extractor.extract(orig_path)
  state.comments = data.comments or {}
end

-- ── Open / close ─────────────────────────────────────────────────────────────

function M.toggle(orig_path, main_buf)
  -- Close if open
  if state.panel_win and vim.api.nvim_win_is_valid(state.panel_win) then
    vim.api.nvim_win_close(state.panel_win, true)
    state.panel_win = nil
    return
  end

  if orig_path then
    state.orig_path = orig_path
    -- Only load from disk if we have no in-memory comments yet.
    -- If comments are already loaded (e.g. after a reply), do NOT reload
    -- from disk — that would throw away in-memory changes.
    if #state.comments == 0 then
      M.load(orig_path)
    end
  end
  if main_buf then
    state.main_buf = main_buf
  end

  -- Create panel buffer once
  if not state.panel_buf or not vim.api.nvim_buf_is_valid(state.panel_buf) then
    state.panel_buf = vim.api.nvim_create_buf(false, true)
    vim.api.nvim_buf_set_name(state.panel_buf, "[neoffice:comments]")
    vim.api.nvim_set_option_value("filetype", "neoffice_comments", { buf = state.panel_buf })
    vim.api.nvim_set_option_value("bufhidden", "hide", { buf = state.panel_buf })
    M._setup_keymaps()
  end

  render()

  local width = math.floor(vim.o.columns * 0.35)
  vim.cmd("botright " .. width .. "vsplit")
  state.panel_win = vim.api.nvim_get_current_win()
  vim.api.nvim_win_set_buf(state.panel_win, state.panel_buf)

  vim.api.nvim_set_option_value("number", false, { win = state.panel_win })
  vim.api.nvim_set_option_value("relativenumber", false, { win = state.panel_win })
  vim.api.nvim_set_option_value("signcolumn", "no", { win = state.panel_win })
  vim.api.nvim_set_option_value("cursorline", true, { win = state.panel_win })
  vim.api.nvim_set_option_value("wrap", true, { win = state.panel_win })
  vim.api.nvim_set_option_value("linebreak", true, { win = state.panel_win })
  vim.api.nvim_set_option_value(
    "winhighlight",
    "Normal:NeofficeCommentPanel,CursorLine:NeofficeCommentCursor",
    { win = state.panel_win }
  )
end

-- ── Helpers ───────────────────────────────────────────────────────────────────

local function comment_at_cursor()
  local row = vim.api.nvim_win_get_cursor(state.panel_win)[1] - 1
  for r = row, 0, -1 do
    local cid = state.comment_line_map[r]
    if cid then
      for _, cm in ipairs(state.comments) do
        if cm.id == cid then
          return cm
        end
      end
    end
  end
  return nil
end

-- flush() intentionally does NOT write to disk.
-- Comments are persisted only during the main save flow in convert.from_text(),
-- which re-injects the in-memory state after direct XML editing.
local function flush()
  if state.main_buf and vim.api.nvim_buf_is_valid(state.main_buf) then
    vim.cmd("redrawstatus")
  end
end

---Return the current in-memory comment list (called by convert.from_text).
function M.get_comments()
  return state.comments
end

-- ── User actions ─────────────────────────────────────────────────────────────

function M.add_comment(main_buf)
  main_buf = main_buf or state.main_buf
  if not main_buf then
    return
  end

  local win = vim.fn.bufwinid(main_buf)
  local line_text = ""
  if win ~= -1 then
    local row = vim.api.nvim_win_get_cursor(win)[1]
    line_text = vim.api.nvim_buf_get_lines(main_buf, row - 1, row, false)[1] or ""
  end

  vim.ui.input({ prompt = "New comment: " }, function(text)
    if not text or text == "" then
      return
    end
    local cm = {
      id = tostring(#state.comments + 1),
      author = vim.env.USER or "nvim-user",
      date = os.date("!%Y-%m-%dT%H:%M:%SZ"),
      text = text,
      replies = {},
      resolved = false,
      anchor = line_text:sub(1, 60),
    }
    table.insert(state.comments, cm)
    flush()
    render()
    -- Refresh anchor signs in main buffer
    if main_buf then
      local lines = vim.api.nvim_buf_get_lines(main_buf, 0, -1, false)
      M.draw_anchors(main_buf, lines)
    end
    vim.notify("[neoffice] Comment added", vim.log.levels.INFO)
  end)
end

function M.reply()
  local cm = comment_at_cursor()
  if not cm then
    vim.notify("[neoffice] No comment under cursor", vim.log.levels.WARN)
    return
  end
  vim.ui.input({ prompt = "Reply to @" .. cm.author .. ": " }, function(text)
    if not text or text == "" then
      return
    end
    table.insert(cm.replies, {
      author = vim.env.USER or "nvim-user",
      date = os.date("!%Y-%m-%dT%H:%M:%SZ"),
      text = text,
    })
    flush()
    render()
    vim.notify("[neoffice] Reply added", vim.log.levels.INFO)
  end)
end

function M.toggle_resolve()
  local cm = comment_at_cursor()
  if not cm then
    return
  end
  cm.resolved = not cm.resolved
  flush()
  render()
  vim.notify("[neoffice] Comment " .. (cm.resolved and "resolved ✓" or "reopened"), vim.log.levels.INFO)
end

function M.delete_comment()
  local cm = comment_at_cursor()
  if not cm then
    return
  end
  vim.ui.select({ "Delete", "Cancel" }, { prompt = "Delete comment?" }, function(choice)
    if choice ~= "Delete" then
      return
    end
    state.comments = vim.tbl_filter(function(c)
      return c.id ~= cm.id
    end, state.comments)
    flush()
    render()
    vim.notify("[neoffice] Comment deleted", vim.log.levels.INFO)
  end)
end

-- ── Anchor signs in main buffer ───────────────────────────────────────────────

function M.draw_anchors(buf, text_lines)
  vim.api.nvim_buf_clear_namespace(buf, NS_ANCHORS, 0, -1)
  for idx, cm in ipairs(state.comments) do
    if cm.anchor and cm.anchor ~= "" then
      local needle = cm.anchor:sub(1, 20)
      for lnum, line in ipairs(text_lines) do
        if line:find(needle, 1, true) then
          local total = 1 + #(cm.replies or {})
          vim.api.nvim_buf_set_extmark(buf, NS_ANCHORS, lnum - 1, 0, {
            sign_text = "💬",
            sign_hl_group = "NeofficeCommentSign",
            virt_text = {
              {
                string.format("  💬 %d comment%s", total, total > 1 and "s" or ""),
                "NeofficeCommentVirt",
              },
            },
            virt_text_pos = "eol",
          })
          break
        end
      end
    end
  end
end

-- ── Panel keymaps ─────────────────────────────────────────────────────────────

function M._setup_keymaps()
  local cfg = config.get()
  local buf = state.panel_buf
  local o = { buffer = buf, silent = true, noremap = true }

  vim.keymap.set("n", cfg.mappings.reply_comment, M.reply, vim.tbl_extend("force", o, { desc = "Reply to comment" }))
  vim.keymap.set(
    "n",
    cfg.mappings.resolve_comment,
    M.toggle_resolve,
    vim.tbl_extend("force", o, { desc = "Resolve/reopen comment" })
  )
  vim.keymap.set(
    "n",
    cfg.mappings.delete_comment,
    M.delete_comment,
    vim.tbl_extend("force", o, { desc = "Delete comment" })
  )
  vim.keymap.set("n", cfg.mappings.close_panel, function()
    if state.panel_win and vim.api.nvim_win_is_valid(state.panel_win) then
      vim.api.nvim_win_close(state.panel_win, true)
      state.panel_win = nil
    end
  end, vim.tbl_extend("force", o, { desc = "Close comments panel" }))
  vim.keymap.set("n", "<Esc>", function()
    if state.panel_win and vim.api.nvim_win_is_valid(state.panel_win) then
      vim.api.nvim_win_close(state.panel_win, true)
      state.panel_win = nil
    end
  end, o)
end

return M
