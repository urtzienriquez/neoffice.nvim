-- neoffice/track_changes.lua
-- Fixed: non-paired tags, undo support, proper text restoration

local extractor = require("neoffice.extractor")
local zip = require("neoffice.zip")
local xml = require("neoffice.xml")
local M = {}

local NS = vim.api.nvim_create_namespace("neoffice_track_changes")

-- Per-buffer state:  buf  →  { changes = [], orig_path = "" }
local _state = {}

local HL = {
  insert = "NeofficeInsert",
  delete = "NeofficeDelete",
}

-- ── Highlights ───────────────────────────────────────────────────────────────

function M.setup_highlights()
  vim.api.nvim_set_hl(0, HL.insert, { fg = "#4ec994", underline = true, default = true })
  vim.api.nvim_set_hl(0, HL.delete, { fg = "#f87171", strikethrough = true, default = true })
end

-- ── Load ─────────────────────────────────────────────────────────────────────

-- Helper to convert byte offset to (line, col)
local function byte_to_pos(all_lines, byte_offset)
  local accumulated = 0
  for i, line_text in ipairs(all_lines) do
    local line_len = #line_text + 1 -- +1 for the \n
    if accumulated + line_len > byte_offset then
      return i - 1, byte_offset - accumulated
    end
    accumulated = accumulated + line_len
  end
  return #all_lines - 1, #all_lines[#all_lines]
end

function M.refresh(buf)
  buf = buf or vim.api.nvim_get_current_buf()
  local state = _state[buf]
  if not state or not state.changes then
    return
  end

  vim.api.nvim_buf_clear_namespace(buf, NS, 0, -1)

  local all_lines = vim.api.nvim_buf_get_lines(buf, 0, -1, false)
  local full_text = table.concat(all_lines, "\n")

  for s, tag in full_text:gmatch("()(<text:change[^>]-/>)") do
    local id = tag:match('text:change%-id="([^"]+)"')
    local is_primary = tag:find("change%-start") or not tag:find("change%-end")

    if id and is_primary then
      local ch = nil
      for _, c in ipairs(state.changes) do
        if c.id == id or c.id:match(id .. "$") then
          ch = c
          break
        end
      end

      local line, _ = byte_to_pos(all_lines, s - 1)
      local ch_type = (ch and ch.type) or (tag:find("change%-start") and "insert" or "delete")
      local author = (ch and ch.author) or "?"

      -- --- NEW: SANITIZE PREVIEW TEXT ---
      local display_text = (ch and ch.text) or ""
      if ch_type == "delete" then
        -- 1. Strip Author Name if it matches the metadata
        if ch and ch.author and ch.author ~= "?" then
          display_text = display_text:gsub("^" .. vim.pesc(ch.author), "")
        end
        -- 2. Strip ISO Timestamp (YYYY-MM-DDTHH:MM:SS)
        display_text = display_text:gsub("^%d%d%d%d%-%d%d%-%d%dT%d%d:%d%d:%d%d%.?%d*", "")
      end

      -- Clean up whitespace and limit length
      display_text = display_text:gsub("^%s+", ""):sub(1, 40):gsub("\n", " ")
      -- ----------------------------------

      local hl = ch_type == "insert" and HL.insert or HL.delete
      local icon = ch_type == "insert" and "▶" or "✕"
      local vt = string.format("%s %s (@%s)", icon, display_text, author)

      vim.api.nvim_buf_set_extmark(buf, NS, line, 0, {
        virt_text = { { vt, hl } },
        virt_text_pos = "eol",
        sign_text = ch_type == "insert" and "▶" or "✕",
        sign_hl_group = hl,
      })
    end
  end
end

function M.load(buf, orig_path, text_lines)
  _state[buf] = { changes = {}, orig_path = orig_path }
  local data = extractor.extract(orig_path)
  _state[buf].changes = data.track_changes or {}

  M.setup_highlights() -- Ensure highlights are defined
  M.refresh(buf)

  if #_state[buf].changes > 0 then
    vim.notify(string.format("[neoffice] %d track change(s) loaded", #_state[buf].changes))
  end
end

-- ── Navigation ───────────────────────────────────────────────────────────────

-- Helper to find every <text:change> occurrence in the buffer
local function get_all_tag_positions(buf)
  local all_lines = vim.api.nvim_buf_get_lines(buf, 0, -1, false)
  local full_text = table.concat(all_lines, "\n")
  local positions = {}

  -- Pattern: () captures start index, then we match the tag, then () captures end index
  -- Variable 's' will receive the start position (number)
  -- Variable '_' will receive the tag text (string) which we ignore
  for s, _, _ in full_text:gmatch("()(<text:change[^>]->)()") do
    local line, col = byte_to_pos(all_lines, s - 1)
    table.insert(positions, { line = line, col = col })
  end

  return positions
end

function M.next_change()
  local buf = vim.api.nvim_get_current_buf()
  local cursor = vim.api.nvim_win_get_cursor(0) -- {row, col} 1-indexed row
  local positions = get_all_tag_positions(buf)

  -- Find the first tag that is strictly AFTER the current cursor position
  for _, pos in ipairs(positions) do
    if pos.line > cursor[1] - 1 or (pos.line == cursor[1] - 1 and pos.col > cursor[2]) then
      vim.api.nvim_win_set_cursor(0, { pos.line + 1, pos.col })
      vim.notify("[neoffice] Next change tag", vim.log.levels.INFO)
      return
    end
  end
  vim.notify("[neoffice] No more changes", vim.log.levels.INFO)
end

function M.prev_change()
  local buf = vim.api.nvim_get_current_buf()
  local cursor = vim.api.nvim_win_get_cursor(0)
  local positions = get_all_tag_positions(buf)

  -- Find the first tag that is strictly BEFORE the current cursor position
  for i = #positions, 1, -1 do
    local pos = positions[i]
    if pos.line < cursor[1] - 1 or (pos.line == cursor[1] - 1 and pos.col < cursor[2]) then
      vim.api.nvim_win_set_cursor(0, { pos.line + 1, pos.col })
      vim.notify("[neoffice] Previous change tag", vim.log.levels.INFO)
      return
    end
  end
  vim.notify("[neoffice] No previous changes", vim.log.levels.INFO)
end

-- ── Change Detection and Selection ───────────────────────────────────────────

---Analyze the change at cursor to determine type and boundaries (inclusive of tags)
local function analyze_change_at_cursor()
  local buf = vim.api.nvim_get_current_buf()
  local cursor_pos = vim.api.nvim_win_get_cursor(0)
  local cursor_line_idx = cursor_pos[1] - 1
  local cursor_col = cursor_pos[2]

  local all_lines = vim.api.nvim_buf_get_lines(buf, 0, -1, false)
  local line_text = all_lines[cursor_line_idx + 1]
  if not line_text then
    return nil
  end

  -- 1. Check if cursor is directly on a tag
  local pattern = "(<text:change[^>]-/>)"
  local start_ptr = 1
  while true do
    local s, e, tag_content = line_text:find(pattern, start_ptr)
    if not s then
      break
    end

    if cursor_col >= (s - 1) and cursor_col <= e then
      local change_id = tag_content:match('text:change%-id="([^"]+)"')

      -- If it's a paired tag, use the robust finder to get full boundaries
      if tag_content:match("text:change%-start") or tag_content:match("text:change%-end") then
        return M.find_paired_range(all_lines, change_id)
      else
        -- It's a point deletion tag
        return {
          type = "point",
          change_id = change_id,
          line = cursor_line_idx,
          col_start = s - 1,
          col_end = e - 1,
        }
      end
    end
    start_ptr = e + 1
  end

  -- 2. Fallback: Search radius if not directly on a tag
  local search_radius = 5
  for offset = 1, search_radius do
    for _, check_idx in ipairs({ cursor_line_idx - offset, cursor_line_idx + offset }) do
      local l = all_lines[check_idx + 1]
      if l then
        local s, e, tag_content = l:find(pattern)
        if s then
          local change_id = tag_content:match('text:change%-id="([^"]+)"')
          if tag_content:match("text:change%-start") or tag_content:match("text:change%-end") then
            return M.find_paired_range(all_lines, change_id)
          else
            return { type = "point", change_id = change_id, line = check_idx, col_start = s - 1, col_end = e - 1 }
          end
        end
      end
    end
  end
  return nil
end

---Finds the exact start and end coordinates of a specific tag in the buffer
local function get_tag_boundaries(all_lines, pattern)
  local full_text = table.concat(all_lines, "\n")
  local s, e = full_text:find(pattern)
  if not s or not e then
    return nil
  end

  local l1, c1 = byte_to_pos(all_lines, s - 1)
  local l2, c2 = byte_to_pos(all_lines, e) -- 'e' is inclusive end, so this is the exclusive end for nvim_buf_set_text
  return { l1 = l1, c1 = c1, l2 = l2, c2 = c2 }
end

function M.find_paired_range(all_lines, change_id)
  local escaped_id = change_id:gsub("([%-%.%+%[%]%(%)%$%^%%%?%*])", "%%%1")

  local start_tag =
    get_tag_boundaries(all_lines, '<text:change%-start[^>]*text:change%-id="' .. escaped_id .. '"[^>]*/>')
  local end_tag = get_tag_boundaries(all_lines, '<text:change%-end[^>]*text:change%-id="' .. escaped_id .. '"[^>]*/>')

  if start_tag and end_tag then
    return {
      type = "paired",
      change_id = change_id,
      start_tag = start_tag,
      end_tag = end_tag,
    }
  end
  return nil
end

function M.select_change()
  local info = analyze_change_at_cursor()
  if not info then
    vim.notify("[neoffice] No change found at cursor", vim.log.levels.WARN)
    return
  end

  if info.type == "point" then
    -- Select the single <text:change ... /> tag
    vim.api.nvim_win_set_cursor(0, { info.line + 1, info.col_start })
    vim.cmd("normal! v")
    vim.api.nvim_win_set_cursor(0, { info.line + 1, info.col_end })
  elseif info.type == "paired" then
    -- UPDATED: Use the robust coordinates start_tag and end_tag
    -- Move to start of opening tag, enter visual, move to end of closing tag
    vim.api.nvim_win_set_cursor(0, { info.start_tag.l1 + 1, info.start_tag.c1 })
    vim.cmd("normal! v")
    -- c2 is exclusive, so we go to c2 - 1 to land on the '>'
    vim.api.nvim_win_set_cursor(0, { info.end_tag.l1 + 1, info.end_tag.c2 - 1 })
  end

  vim.notify(string.format("[neoffice] Selected: %s", info.change_id), vim.log.levels.INFO)
end

-- ── Change at Cursor ─────────────────────────────────────────────────────────

function M.change_at_cursor()
  local buf = vim.api.nvim_get_current_buf()

  -- FIXED: Don't rely on stale _state, analyze from buffer directly
  local info = analyze_change_at_cursor()
  if not info then
    return nil
  end

  -- Try to find matching change in state (if available)
  local state = _state[buf]
  if state and state.changes then
    for _, ch in ipairs(state.changes) do
      local ch_base_id = ch.id:match("^[^_]+_(.+)$") or ch.id
      if ch_base_id == info.change_id or ch.id == info.change_id then
        return ch
      end
    end
  end

  -- FIXED: If not found in state, create a temporary change object from buffer
  -- This handles the undo case where state is stale
  return {
    id = info.change_id,
    type = info.type == "point" and "delete" or "insert",
    author = "?",
    date = "",
    text = "",
    para = 0,
    status = "pending",
  }
end

-- ── Public API ────────────────────────────────────────────────────────────────

-- ── Accept / Reject (ODT) ─────────────────────────────────────────────────────

local function remove_metadata_from_buffer(buf, change_id)
  local lines = vim.api.nvim_buf_get_lines(buf, 0, -1, false)
  local content = table.concat(lines, "\n")
  local escaped_id = change_id:gsub("([%-%.%+%[%]%(%)%$%^%%%?%*])", "%%%1")

  -- The [^>]- pattern handles multi-line blocks and internal tags
  local pattern = '<text:changed%-region[^>]*text:id="' .. escaped_id .. '"[^>]*>.-</text:changed%-region>'

  -- We use a more aggressive match that ignores newlines
  local start_idx, end_idx = content:find(pattern)
  if start_idx and end_idx then
    local new_content = content:sub(1, start_idx - 1) .. content:sub(end_idx + 1)
    local new_lines = vim.split(new_content, "\n")
    vim.api.nvim_buf_set_lines(buf, 0, -1, false, new_lines)
  end
end

function M.accept(change_id)
  local buf = vim.api.nvim_get_current_buf()
  local all_lines = vim.api.nvim_buf_get_lines(buf, 0, -1, false)

  -- Find the tags using the robust search instead of guessing lengths
  local info = M.find_paired_range(all_lines, change_id) or analyze_change_at_cursor()
  if not info then
    vim.notify("[neoffice] Could not find tags for " .. change_id, vim.log.levels.ERROR)
    return
  end

  if info.type == "paired" then
    -- Remove end tag first so start tag position doesn't shift
    vim.api.nvim_buf_set_text(buf, info.end_tag.l1, info.end_tag.c1, info.end_tag.l2, info.end_tag.c2, {})
    vim.api.nvim_buf_set_text(buf, info.start_tag.l1, info.start_tag.c1, info.start_tag.l2, info.start_tag.c2, {})
  else
    -- Point tag (deletion): remove the single <text:change ... /> tag
    vim.api.nvim_buf_set_text(buf, info.line, info.col_start, info.line, info.col_end + 1, {})
  end

  remove_metadata_from_buffer(buf, change_id)
  vim.notify("[neoffice] Accepted: " .. change_id)
end

function M.reject(change_id)
  local buf = vim.api.nvim_get_current_buf()
  local all_lines = vim.api.nvim_buf_get_lines(buf, 0, -1, false)
  local info = M.find_paired_range(all_lines, change_id) or analyze_change_at_cursor()

  if not info then
    return
  end

  if info.type == "paired" then
    -- Reject Insertion: Just delete the tags and the text inside
    vim.api.nvim_buf_set_text(buf, info.start_tag.l1, info.start_tag.c1, info.end_tag.l2, info.end_tag.c2, {})
  else
    -- Reject Deletion: Restore the deleted text
    local ch = M.change_at_cursor()
    local raw_text = (ch and ch.text) or ""

    -- --- NEW: SANITIZATION LOGIC ---
    -- 1. Remove the author's name from the start if it exists
    if ch.author and ch.author ~= "?" then
      raw_text = raw_text:gsub("^" .. vim.pesc(ch.author), "")
    end
    -- 2. Remove the ISO date pattern (e.g., 2026-03-11T21:03:23.123)
    -- This pattern matches YYYY-MM-DDTHH:MM:SS and optional decimals
    raw_text = raw_text:gsub("^%d%d%d%d%-%d%d%-%d%dT%d%d:%d%d:%d%d%.?%d*", "")
    -- -------------------------------

    vim.api.nvim_buf_set_text(buf, info.line, info.col_start, info.line, info.col_end + 1, { raw_text })
  end

  remove_metadata_from_buffer(buf, change_id)
  vim.notify("[neoffice] Rejected: " .. change_id)
end

function M.accept_all()
  local buf = vim.api.nvim_get_current_buf()
  local state = _state[buf]
  if not state then
    return
  end

  for _, ch in ipairs(state.changes) do
    if ch.status == "pending" then
      M.accept(ch.id)
    end
  end
end

function M.reject_all()
  local buf = vim.api.nvim_get_current_buf()
  local state = _state[buf]
  if not state then
    return
  end

  for _, ch in ipairs(state.changes) do
    if ch.status == "pending" then
      M.reject(ch.id)
    end
  end
end

-- ── Summary ──────────────────────────────────────────────────────────────────

function M.show_summary()
  local buf = vim.api.nvim_get_current_buf()
  local state = _state[buf]
  local changes = state and state.changes or {}

  if #changes == 0 then
    vim.notify("[neoffice] No track changes", vim.log.levels.INFO)
    return
  end

  local lines = { " Track Changes ", string.rep("─", 56) }
  for _, ch in ipairs(changes) do
    local icon = ch.type == "insert" and "▶ +" or "✕ -"
    table.insert(lines, string.format(" %s %-32s  @%-14s", icon, ch.text:sub(1, 32):gsub("\n", " "), ch.author))
  end
  table.insert(lines, "")
  table.insert(lines, " <leader>dy=accept  <leader>dn=reject  :DocAccept all")
  table.insert(lines, " ]c=next  [c=prev  <leader>dv=select")

  local fbuf = vim.api.nvim_create_buf(false, true)
  vim.api.nvim_buf_set_lines(fbuf, 0, -1, false, lines)
  vim.api.nvim_set_option_value("modifiable", false, { buf = fbuf })

  local w = 60
  vim.api.nvim_open_win(fbuf, true, {
    relative = "editor",
    width = w,
    height = math.min(#lines + 1, 20),
    row = 3,
    col = math.floor((vim.o.columns - w) / 2),
    border = "rounded",
    title = " Track Changes ",
    title_pos = "center",
  })

  vim.keymap.set("n", "q", "<cmd>close<CR>", { buffer = fbuf, silent = true })
  vim.keymap.set("n", "<Esc>", "<cmd>close<CR>", { buffer = fbuf, silent = true })
end

return M
