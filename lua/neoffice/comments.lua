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
  if win == -1 then
    vim.notify("[neoffice] Buffer not visible", vim.log.levels.WARN)
    return
  end

  -- Check if we're in visual mode (range selected)
  local mode = vim.api.nvim_get_mode().mode
  local is_visual = mode == "v" or mode == "V" or mode == "\22"

  local start_pos, end_pos
  if is_visual then
    -- Exit visual mode to capture marks
    vim.api.nvim_feedkeys(vim.api.nvim_replace_termcodes("<Esc>", true, false, true), "x", false)
    start_pos = vim.api.nvim_buf_get_mark(main_buf, "<")
    end_pos = vim.api.nvim_buf_get_mark(main_buf, ">")
  end

  vim.ui.input({ prompt = "New comment: " }, function(text)
    if not text or text == "" then
      return
    end

    local comment_id =
      string.format("__Annotation__%d_%d", math.random(10000, 99999), math.random(1000000000, 9999999999))

    local author = vim.env.USER or "nvim-user"
    local date = os.date("!%Y-%m-%dT%H:%M:%SZ")
    local initials = author:sub(1, 2):upper()

    local annotation_start = string.format(
      '<office:annotation office:name="%s" loext:resolved="false">'
        .. "<dc:creator>%s</dc:creator>"
        .. "<dc:date>%s</dc:date>"
        .. "<meta:creator-initials>%s</meta:creator-initials>"
        .. '<text:p text:style-name="Comment">%s</text:p>'
        .. "</office:annotation>",
      comment_id,
      author,
      date,
      initials,
      text:gsub("&", "&amp;"):gsub("<", "&lt;"):gsub(">", "&gt;")
    )

    if is_visual and start_pos and end_pos then
      local annotation_end = string.format('<office:annotation-end office:name="%s"/>', comment_id)

      local start_row, start_col = start_pos[1] - 1, start_pos[2]
      local end_row, end_col = end_pos[1] - 1, end_pos[2]

      -- Get the line to find the next character boundary
      local end_line = vim.api.nvim_buf_get_lines(main_buf, end_row, end_row + 1, false)[1]

      -- Move to next UTF-8 character boundary (after the selected character)
      local end_col_after = vim.str_byteindex(end_line, vim.str_utfindex(end_line, end_col) + 1)

      -- Insert end marker first
      vim.api.nvim_buf_set_text(main_buf, end_row, end_col_after, end_row, end_col_after, { annotation_end })

      -- Insert start marker
      vim.api.nvim_buf_set_text(main_buf, start_row, start_col, start_row, start_col, { annotation_start })

      vim.notify("[neoffice] Range comment added (save to persist)", vim.log.levels.INFO)
    else
      local row, col = unpack(vim.api.nvim_win_get_cursor(win))
      vim.api.nvim_buf_set_text(main_buf, row - 1, col, row - 1, col, { annotation_start })
      vim.notify("[neoffice] Comment added (save to persist)", vim.log.levels.INFO)
    end

    local anchor_row = start_pos and (start_pos[1] - 1) or (vim.api.nvim_win_get_cursor(win)[1] - 1)
    local anchor_line = vim.api.nvim_buf_get_lines(main_buf, anchor_row, anchor_row + 1, false)[1] or ""
    local cm = {
      id = comment_id,
      author = author,
      date = date,
      text = text,
      replies = {},
      resolved = false,
      anchor = anchor_line:sub(1, 60),
    }
    table.insert(state.comments, cm)

    flush()
    render()

    local lines = vim.api.nvim_buf_get_lines(main_buf, 0, -1, false)
    M.draw_anchors(main_buf, lines)
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

    -- Add to in-memory state
    local reply = {
      author = vim.env.USER or "nvim-user",
      date = os.date("!%Y-%m-%dT%H:%M:%SZ"),
      text = text,
    }
    table.insert(cm.replies, reply)

    -- Now update the XML in the buffer
    if not state.main_buf or not vim.api.nvim_buf_is_valid(state.main_buf) then
      vim.notify("[neoffice] Main buffer not available", vim.log.levels.WARN)
      return
    end

    -- Find the annotation in the buffer by ID
    local lines = vim.api.nvim_buf_get_lines(state.main_buf, 0, -1, false)
    local id_pattern = 'office:name="' .. cm.id:gsub("([%-%.%+%[%]%(%)%$%^%%%?%*])", "%%%1") .. '"'

    for lnum, line in ipairs(lines) do
      if line:find(id_pattern, 1, true) then
        -- Found the annotation - insert reply before closing tag
        local initials = (reply.author or "U"):sub(1, 2):upper()
        local reply_xml = string.format(
          '<office:annotation loext:parent-name="%s" loext:resolved="false">'
            .. "<dc:creator>%s</dc:creator>"
            .. "<dc:date>%s</dc:date>"
            .. "<meta:creator-initials>%s</meta:creator-initials>"
            .. '<text:p text:style-name="Comment">%s</text:p>'
            .. "</office:annotation>",
          cm.id,
          reply.author,
          reply.date,
          initials,
          text:gsub("&", "&amp;"):gsub("<", "&lt;"):gsub(">", "&gt;")
        )

        -- Insert reply annotation after parent closing tag
        local close_tag = "</office:annotation>"
        local insert_pos = line:find(close_tag, 1, true)
        if insert_pos then
          local new_line = line:sub(1, insert_pos + #close_tag - 1) .. reply_xml .. line:sub(insert_pos + #close_tag)
          vim.api.nvim_buf_set_lines(state.main_buf, lnum - 1, lnum, false, { new_line })
        end
        break
      end
    end

    flush()
    render()
    vim.notify("[neoffice] Reply added (save to persist)", vim.log.levels.INFO)
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

  for _, cm in ipairs(state.comments) do
    if not cm.id then
      goto continue
    end

    -- Search for the annotation by its exact ID in the XML
    -- Escape special pattern characters in the ID
    local id_escaped = cm.id:gsub("([%-%.%+%[%]%(%)%$%^%%%?%*])", "%%%1")
    local search_pattern = 'office:name="' .. id_escaped .. '"'

    for lnum, line in ipairs(text_lines) do
      -- Use plain text search (not pattern matching)
      if line:find(search_pattern, 1, true) then
        local total = 1 + #(cm.replies or {})
        vim.api.nvim_buf_set_extmark(buf, NS_ANCHORS, lnum - 1, 0, {
          sign_text = "💬",
          sign_hl_group = "NeofficeCommentSign",
          virt_text = {
            {
              string.format("   %d comment%s", total, total > 1 and "s" or ""),
              "NeofficeCommentVirt",
            },
          },
          virt_text_pos = "eol",
        })
        break -- Found it, move to next comment
      end
    end

    ::continue::
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
