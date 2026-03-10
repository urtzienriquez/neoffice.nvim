-- neoffice/track_changes.lua
-- Renders track-change insertions/deletions as extmarks in the proxy buffer.
-- accept/reject rewrite the source ZIP XML via extractor + zip modules.

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

---Load track changes for buf from orig_path and draw extmarks.
---@param buf       number
---@param orig_path string
---@param text_lines string[]
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

-- ── Query ────────────────────────────────────────────────────────────────────

---Return the change closest to the cursor line in the current buffer.
function M.change_at_cursor()
  local buf = vim.api.nvim_get_current_buf()
  local state = _state[buf]
  if not state or #state.changes == 0 then
    return nil
  end

  local line = vim.api.nvim_win_get_cursor(0)[1] - 1
  local best, best_dist = nil, math.huge
  for _, ch in ipairs(state.changes) do
    local d = math.abs(ch.para - line)
    if d < best_dist then
      best, best_dist = ch, d
    end
  end
  return best
end

-- ── Accept / Reject ───────────────────────────────────────────────────────────

local function rewrite_docx(orig_path, change_id, action)
  local raw = zip.read_entry(orig_path, "word/document.xml")
  if not raw then
    vim.notify("[neoffice] Could not read document.xml", vim.log.levels.ERROR)
    return false
  end

  local root = xml.parse(raw)
  local paras = xml.find_all(root, "w:p")
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

function M.accept(change_id, orig_path)
  orig_path = orig_path or (_state[vim.api.nvim_get_current_buf()] or {}).orig_path
  if not orig_path then
    return
  end
  local ext = orig_path:match("%.(%w+)$"):lower()
  if ext == "docx" then
    rewrite_docx(orig_path, change_id, "accept")
  end
  vim.notify("[neoffice] Accepted: " .. change_id, vim.log.levels.INFO)
end

function M.reject(change_id, orig_path)
  orig_path = orig_path or (_state[vim.api.nvim_get_current_buf()] or {}).orig_path
  if not orig_path then
    return
  end
  local ext = orig_path:match("%.(%w+)$"):lower()
  if ext == "docx" then
    rewrite_docx(orig_path, change_id, "reject")
  end
  vim.notify("[neoffice] Rejected: " .. change_id, vim.log.levels.INFO)
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

-- ── Summary float ────────────────────────────────────────────────────────────

function M.show_summary()
  local buf = vim.api.nvim_get_current_buf()
  local state = _state[buf]
  local changes = state and state.changes or {}

  if #changes == 0 then
    vim.notify("[neoffice] No track changes found", vim.log.levels.INFO)
    return
  end

  local lines = { " Track Changes ", string.rep("─", 56) }
  for _, ch in ipairs(changes) do
    local icon = ch.type == "insert" and "▶ +" or "✕ -"
    table.insert(lines, string.format(" %s %-32s  @%-14s", icon, ch.text:sub(1, 32):gsub("\n", " "), ch.author))
  end
  table.insert(lines, "")
  table.insert(lines, " <leader>dy = accept   <leader>dn = reject   :DocAccept all")

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
