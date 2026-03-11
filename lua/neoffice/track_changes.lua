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

function M.load(buf, orig_path, text_lines)
  vim.api.nvim_buf_clear_namespace(buf, NS, 0, -1)
  _state[buf] = { changes = {}, orig_path = orig_path }

  local data = extractor.extract(orig_path)
  local changes = data.track_changes
  _state[buf].changes = changes

  for _, ch in ipairs(changes) do
    local line = math.min(ch.para, math.max(#text_lines - 1, 0))
    local hl = ch.type == "insert" and HL.insert or HL.delete
    local icon = ch.type == "insert" and "▶ +" or "✕ -"
    local vt = string.format("  %s%s  (@%s)", icon, ch.text:sub(1, 40):gsub("\n", " "), ch.author)

    vim.api.nvim_buf_set_extmark(buf, NS, line, 0, {
      virt_text = { { vt, hl } },
      virt_text_pos = "eol",
      sign_text = ch.type == "insert" and "▶" or "✕",
      sign_hl_group = hl,
      hl_mode = "combine",
    })
  end

  if #changes > 0 then
    vim.notify(string.format("[neoffice] %d track change(s) loaded", #changes), vim.log.levels.INFO)
  end
end

-- ── Navigation ───────────────────────────────────────────────────────────────

---Find the next text:change* tag (any type)
---FIXED: Now correctly matches self-closing <text:change .../> tags
local function find_next_change_tag(buf, start_line)
  local total_lines = vim.api.nvim_buf_line_count(buf)

  for line_num = start_line, total_lines - 1 do
    local line = vim.api.nvim_buf_get_lines(buf, line_num, line_num + 1, false)[1]
    if line then
      -- Match any change tag: start, end, or point (self-closing)
      -- FIXED: Use [%s>] to match "text:change" followed by space or >
      -- This prevents matching "text:change-start" when looking for "text:change"
      if line:match("<text:change%-start") or line:match("<text:change%-end") or line:match("<text:change[%s>]") then
        return line_num, line
      end
    end
  end

  return nil, nil
end

local function find_prev_change_tag(buf, start_line)
  for line_num = start_line, 0, -1 do
    local line = vim.api.nvim_buf_get_lines(buf, line_num, line_num + 1, false)[1]
    if line then
      if line:match("<text:change%-start") or line:match("<text:change%-end") or line:match("<text:change[%s>]") then
        return line_num, line
      end
    end
  end

  return nil, nil
end

function M.next_change()
  local buf = vim.api.nvim_get_current_buf()
  local cursor_line = vim.api.nvim_win_get_cursor(0)[1] - 1

  local line_num, line = find_next_change_tag(buf, cursor_line + 1)

  if line_num then
    local tag_start = line:find("<text:change")
    vim.api.nvim_win_set_cursor(0, { line_num + 1, tag_start - 1 })
    vim.notify("[neoffice] Next change", vim.log.levels.INFO)
  else
    vim.notify("[neoffice] No more changes", vim.log.levels.INFO)
  end
end

function M.prev_change()
  local buf = vim.api.nvim_get_current_buf()
  local cursor_line = vim.api.nvim_win_get_cursor(0)[1] - 1

  local line_num, line = find_prev_change_tag(buf, cursor_line - 1)

  if line_num then
    local tag_start = line:find("<text:change")
    vim.api.nvim_win_set_cursor(0, { line_num + 1, tag_start - 1 })
    vim.notify("[neoffice] Previous change", vim.log.levels.INFO)
  else
    vim.notify("[neoffice] No previous changes", vim.log.levels.INFO)
  end
end

-- ── Change Detection and Selection ───────────────────────────────────────────

---Analyze the change at cursor to determine type and boundaries
local function analyze_change_at_cursor()
  local buf = vim.api.nvim_get_current_buf()
  local cursor_pos = vim.api.nvim_win_get_cursor(0)
  local cursor_line = cursor_pos[1] - 1

  local line = vim.api.nvim_buf_get_lines(buf, cursor_line, cursor_line + 1, false)[1]
  if not line then
    return nil
  end

  local change_id = nil
  local tag_type = nil

  -- FIXED: Better pattern matching for point deletions
  -- Check in order of specificity
  if line:match('<text:change%-start[^>]*text:change%-id="([^"]+)"') then
    change_id = line:match('<text:change%-start[^>]*text:change%-id="([^"]+)"')
    tag_type = "start"
  elseif line:match('<text:change%-end[^>]*text:change%-id="([^"]+)"') then
    change_id = line:match('<text:change%-end[^>]*text:change%-id="([^"]+)"')
    tag_type = "end"
  elseif line:match('<text:change[%s>][^>]*text:change%-id="([^"]+)"') then
    -- FIXED: Match "text:change" followed by space or >, not "text:change-"
    change_id = line:match('<text:change[%s>][^>]*text:change%-id="([^"]+)"')
    tag_type = "point"
  end

  if not change_id then
    -- Search nearby lines
    local search_radius = 5
    for offset = 1, search_radius do
      for _, check_line in ipairs({ cursor_line - offset, cursor_line + offset }) do
        if check_line >= 0 then
          line = vim.api.nvim_buf_get_lines(buf, check_line, check_line + 1, false)[1]
          if line then
            if line:match('<text:change%-start[^>]*text:change%-id="([^"]+)"') then
              change_id = line:match('<text:change%-start[^>]*text:change%-id="([^"]+)"')
              tag_type = "start"
              cursor_line = check_line
              break
            elseif line:match('<text:change[%s>][^>]*text:change%-id="([^"]+)"') then
              change_id = line:match('<text:change[%s>][^>]*text:change%-id="([^"]+)"')
              tag_type = "point"
              cursor_line = check_line
              break
            end
          end
        end
      end
      if change_id then
        break
      end
    end
  end

  if not change_id then
    return nil
  end

  -- For point deletions, return simple info
  if tag_type == "point" then
    local all_lines = vim.api.nvim_buf_get_lines(buf, 0, -1, false)
    local line_text = all_lines[cursor_line + 1]
    local tag_start, tag_end = line_text:find("<text:change[%s>][^>]*/>")

    if tag_start then
      return {
        type = "point",
        change_id = change_id,
        line = cursor_line,
        col_start = tag_start - 1,
        col_end = tag_end,
      }
    end
  end

  -- For paired changes, find the matching end tag
  local all_lines = vim.api.nvim_buf_get_lines(buf, 0, -1, false)
  local full_text = table.concat(all_lines, "\n")

  local escaped_id = change_id:gsub("([%-%.%+%[%]%(%)%$%^%%%?%*])", "%%%1")
  local start_pattern = '<text:change%-start[^>]*text:change%-id="' .. escaped_id .. '"[^>]*/>'
  local end_pattern = '<text:change%-end[^>]*text:change%-id="' .. escaped_id .. '"[^>]*/>'

  local start_pos = full_text:find(start_pattern)
  local _, start_end = full_text:find(start_pattern)
  local end_start, end_pos = full_text:find(end_pattern)

  if start_pos and end_pos then
    local function byte_to_pos(byte_offset)
      local line = 0
      local col = byte_offset
      local accumulated = 0

      for i, line_text in ipairs(all_lines) do
        local line_len = #line_text + 1
        if accumulated + line_len > byte_offset then
          col = byte_offset - accumulated
          line = i - 1
          break
        end
        accumulated = accumulated + line_len
      end

      return line, col
    end

    local start_line, start_col = byte_to_pos(start_end)
    local end_line, end_col = byte_to_pos(end_start - 1)

    return {
      type = "paired",
      change_id = change_id,
      start_line = start_line,
      start_col = start_col,
      end_line = end_line,
      end_col = end_col,
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
    vim.api.nvim_win_set_cursor(0, { info.line + 1, info.col_start })
    vim.cmd("normal! v")
    vim.api.nvim_win_set_cursor(0, { info.line + 1, info.col_end - 1 })
    vim.notify(string.format("[neoffice] Selected point deletion: %s", info.change_id), vim.log.levels.INFO)
  elseif info.type == "paired" then
    vim.api.nvim_win_set_cursor(0, { info.start_line + 1, info.start_col })
    vim.cmd("normal! v")
    vim.api.nvim_win_set_cursor(0, { info.end_line + 1, info.end_col })
    vim.notify(string.format("[neoffice] Selected paired change: %s", info.change_id), vim.log.levels.INFO)
  end
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

-- ── Accept / Reject (ODT) ─────────────────────────────────────────────────────

local function rewrite_odt(orig_path, change_id, action)
  local raw = zip.read_entry(orig_path, "content.xml")
  if not raw then
    vim.notify("[neoffice] Could not read content.xml", vim.log.levels.ERROR)
    return false
  end

  local base_id = change_id:match("^[^_]+_(.+)$") or change_id
  local root = xml.parse(raw)
  local modified_xml = raw

  -- Determine change type and extract deleted content
  local changed_regions = xml.find_all(root, "text:changed-region")
  local change_type = nil
  local deleted_content = nil

  for _, region in ipairs(changed_regions) do
    local region_id = xml.attr(region, "text:id")
    if region_id == base_id then
      if xml.find_first(region, "text:insertion") then
        change_type = "insertion"
      elseif xml.find_first(region, "text:deletion") then
        change_type = "deletion"

        -- FIXED: Extract only the text content, not metadata
        local deletion = xml.find_first(region, "text:deletion")
        if deletion then
          -- Get all text:p, text:h, etc. children (not change-info)
          local content_parts = {}
          for _, child in ipairs(deletion.children or {}) do
            if child.tag and child.tag:match("^text:") and child.tag ~= "text:change" then
              -- Serialize this text element
              table.insert(content_parts, xml.serialize(child))
            end
          end
          deleted_content = table.concat(content_parts, "")
        end
      end
      break
    end
  end

  local escaped_id = base_id:gsub("([%-%.%+%[%]%(%)%$%^%%%?%*])", "%%%1")

  if action == "accept" then
    if change_type == "insertion" then
      -- Accept insertion: Remove markers, keep content
      modified_xml = modified_xml:gsub('<text:change%-start[^>]*text:change%-id="' .. escaped_id .. '"[^>]*/>', "")
      modified_xml = modified_xml:gsub('<text:change%-end[^>]*text:change%-id="' .. escaped_id .. '"[^>]*/>', "")
    elseif change_type == "deletion" then
      -- Accept deletion: Remove the point marker
      modified_xml = modified_xml:gsub('<text:change[%s>][^>]*text:change%-id="' .. escaped_id .. '"[^>]*/>', "")
    end

    -- Remove the changed-region entry
    local region_pattern = '<text:changed%-region[^>]*text:id="' .. escaped_id .. '"[^>]*>.-</text:changed%-region>'
    modified_xml = modified_xml:gsub(region_pattern, "")
  elseif action == "reject" then
    if change_type == "insertion" then
      -- Reject insertion: Remove markers AND content
      local pattern = '<text:change%-start[^>]*text:change%-id="'
        .. escaped_id
        .. '"[^>]*/>.-<text:change%-end[^>]*text:change%-id="'
        .. escaped_id
        .. '"[^>]*/>'
      modified_xml = modified_xml:gsub(pattern, "")
    elseif change_type == "deletion" then
      -- FIXED: Reject deletion: Replace marker with actual deleted text
      if deleted_content and deleted_content ~= "" then
        vim.notify(string.format("[neoffice] Restoring: %s", deleted_content:sub(1, 50)), vim.log.levels.DEBUG)
        modified_xml =
          modified_xml:gsub('<text:change[%s>][^>]*text:change%-id="' .. escaped_id .. '"[^>]*/>', deleted_content)
      else
        -- No content to restore, just remove marker
        vim.notify("[neoffice] No deleted content found to restore", vim.log.levels.WARN)
        modified_xml = modified_xml:gsub('<text:change[%s>][^>]*text:change%-id="' .. escaped_id .. '"[^>]*/>', "")
      end
    end

    -- Remove the changed-region entry
    local region_pattern = '<text:changed%-region[^>]*text:id="' .. escaped_id .. '"[^>]*>.-</text:changed%-region>'
    modified_xml = modified_xml:gsub(region_pattern, "")
  end

  -- Clean up empty tracked-changes container
  if not modified_xml:match("<text:changed%-region") then
    modified_xml = modified_xml:gsub("<text:tracked%-changes[^>]*>%s*</text:tracked%-changes>", "")
  end

  return zip.write_entry(orig_path, "content.xml", modified_xml)
end

local function rewrite_docx(orig_path, change_id, action)
  local raw = zip.read_entry(orig_path, "word/document.xml")
  if not raw then
    vim.notify("[neoffice] Could not read document.xml", vim.log.levels.ERROR)
    return false
  end

  local id_num = change_id:match("_(%d+)$")
  if not id_num then
    return false
  end
  id_num = tonumber(id_num)

  local modified = raw

  if action == "accept" and change_id:match("^ins_") then
    local n = 0
    modified = raw:gsub("(<w:ins [^>]*>)(.-)(</w:ins>)", function(open, inner, close)
      if n == id_num then
        n = n + 1
        return inner
      end
      n = n + 1
      return open .. inner .. close
    end)
  elseif action == "reject" and change_id:match("^ins_") then
    local n = 0
    modified = raw:gsub("<w:ins [^>]*>.-</w:ins>", function(block)
      if n == id_num then
        n = n + 1
        return ""
      end
      n = n + 1
      return block
    end)
  elseif action == "accept" and change_id:match("^del_") then
    local n = 0
    modified = raw:gsub("<w:del [^>]*>.-</w:del>", function(block)
      if n == id_num then
        n = n + 1
        return ""
      end
      n = n + 1
      return block
    end)
  elseif action == "reject" and change_id:match("^del_") then
    local n = 0
    modified = raw:gsub("(<w:del [^>]*>)(.-)(</w:del>)", function(open, inner, close)
      if n == id_num then
        n = n + 1
        inner = inner:gsub("<w:delText", "<w:t"):gsub("</w:delText>", "</w:t>")
        return inner
      end
      n = n + 1
      return open .. inner .. close
    end)
  end

  return zip.write_entry(orig_path, "word/document.xml", modified)
end

-- ── Buffer Reload ─────────────────────────────────────────────────────────────

local function reload_buffer(orig_path)
  local target_buf = nil

  local buf_module = require("neoffice.buffer")
  if buf_module.get_all_meta then
    for buf, meta in pairs(buf_module.get_all_meta()) do
      if meta.orig_path == orig_path and vim.api.nvim_buf_is_valid(buf) then
        target_buf = buf
        break
      end
    end
  end

  if not target_buf then
    local meta = buf_module.get_meta()
    if meta and meta.orig_path == orig_path then
      target_buf = vim.api.nvim_get_current_buf()
    end
  end

  if target_buf then
    local convert = require("neoffice.convert")
    local text_path, para_map, original_root = convert.to_text(orig_path)

    if text_path then
      local cursor_pos = vim.api.nvim_win_get_cursor(0)
      local lines = vim.fn.readfile(text_path)

      vim.api.nvim_buf_set_lines(target_buf, 0, -1, false, lines)

      local new_line_count = vim.api.nvim_buf_line_count(target_buf)
      local new_line = math.min(cursor_pos[1], new_line_count)
      vim.api.nvim_win_set_cursor(0, { new_line, cursor_pos[2] })

      if require("neoffice.config").get().show_track_changes then
        M.load(target_buf, orig_path, lines)
      end

      require("neoffice.comments").load(orig_path)
      require("neoffice.comments").draw_anchors(target_buf, lines)

      vim.notify("[neoffice] Buffer reloaded", vim.log.levels.INFO)
    end
  end
end

-- ── Public API ────────────────────────────────────────────────────────────────

function M.accept(change_id, orig_path)
  orig_path = orig_path or (_state[vim.api.nvim_get_current_buf()] or {}).orig_path
  if not orig_path then
    vim.notify("[neoffice] No document path", vim.log.levels.ERROR)
    return
  end

  local ext = orig_path:match("%.(%w+)$"):lower()
  local success = false

  if ext == "docx" then
    success = rewrite_docx(orig_path, change_id, "accept")
  elseif ext == "odt" then
    success = rewrite_odt(orig_path, change_id, "accept")
  else
    vim.notify("[neoffice] Unsupported format: " .. ext, vim.log.levels.ERROR)
    return
  end

  if success then
    vim.notify("[neoffice] Accepted: " .. change_id, vim.log.levels.INFO)
    reload_buffer(orig_path)
  else
    vim.notify("[neoffice] Failed to accept: " .. change_id, vim.log.levels.ERROR)
  end
end

function M.reject(change_id, orig_path)
  orig_path = orig_path or (_state[vim.api.nvim_get_current_buf()] or {}).orig_path
  if not orig_path then
    vim.notify("[neoffice] No document path", vim.log.levels.ERROR)
    return
  end

  local ext = orig_path:match("%.(%w+)$"):lower()
  local success = false

  if ext == "docx" then
    success = rewrite_docx(orig_path, change_id, "reject")
  elseif ext == "odt" then
    success = rewrite_odt(orig_path, change_id, "reject")
  else
    vim.notify("[neoffice] Unsupported format: " .. ext, vim.log.levels.ERROR)
    return
  end

  if success then
    vim.notify("[neoffice] Rejected: " .. change_id, vim.log.levels.INFO)
    reload_buffer(orig_path)
  else
    vim.notify("[neoffice] Failed to reject: " .. change_id, vim.log.levels.ERROR)
  end
end

function M.accept_all(orig_path)
  local buf = vim.api.nvim_get_current_buf()
  local state = _state[buf]
  if not state then
    return
  end

  for _, ch in ipairs(state.changes) do
    if ch.status == "pending" then
      M.accept(ch.id, orig_path)
    end
  end
end

function M.reject_all(orig_path)
  local buf = vim.api.nvim_get_current_buf()
  local state = _state[buf]
  if not state then
    return
  end

  for _, ch in ipairs(state.changes) do
    if ch.status == "pending" then
      M.reject(ch.id, orig_path)
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
