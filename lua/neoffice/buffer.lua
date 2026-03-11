-- neoffice/buffer.lua
-- Creates and manages the proxy buffer for direct XML editing

local config = require("neoffice.config")
local M = {}

local _meta = {} -- buf  →  { orig_path, text_path, para_map, original_root }
local _conceal_state = {} -- buf → { enabled = bool, wrap = bool }

-- Add this function before M.open_proxy
local function setup_conceal(buf)
  -- Save original settings
  if not _conceal_state[buf] then
    _conceal_state[buf] = {
      wrap = vim.api.nvim_get_option_value("wrap", { win = 0 }),
      linebreak = vim.api.nvim_get_option_value("linebreak", { win = 0 }),
    }
  end

  -- Clear any existing syntax
  vim.api.nvim_buf_call(buf, function()
    vim.cmd("syntax clear")
  end)

  -- Conceal entire comment blocks - just show marker
  vim.api.nvim_buf_call(buf, function()
    -- Conceal entire annotation block (from opening to closing tag)
    -- NOTE: \\{-} is non-greedy match in Vim regex (backslash escaped for Lua)
    vim.fn.matchadd("Conceal", "<office:annotation[^>]*>.\\{-}</office:annotation>", 10, -1, { conceal = "💬" })

    -- Conceal annotation-end markers
    vim.fn.matchadd("Conceal", "<office:annotation-end[^>]\\{-}/>", 10, -1, { conceal = "💬" })
  end)

  -- Enable concealing with wrapping
  vim.api.nvim_set_option_value("conceallevel", 2, { win = 0 })
  vim.api.nvim_set_option_value("concealcursor", "n", { win = 0 })
  vim.api.nvim_set_option_value("wrap", true, { win = 0 })
  vim.api.nvim_set_option_value("linebreak", true, { win = 0 })

  _conceal_state[buf].enabled = true

  vim.notify("[neoffice] Comments concealed as markers", vim.log.levels.INFO)
end

local function disable_conceal(buf)
  -- Clear match highlighting
  vim.api.nvim_buf_call(buf, function()
    vim.fn.clearmatches()
  end)

  vim.api.nvim_set_option_value("conceallevel", 0, { win = 0 })
  vim.api.nvim_set_option_value("concealcursor", "", { win = 0 })

  -- Restore original settings
  if _conceal_state[buf] then
    vim.api.nvim_set_option_value("wrap", _conceal_state[buf].wrap, { win = 0 })
    vim.api.nvim_set_option_value("linebreak", _conceal_state[buf].linebreak, { win = 0 })
  end

  _conceal_state[buf].enabled = false
end

function M.toggle_conceal()
  local buf = vim.api.nvim_get_current_buf()

  if not _meta[buf] then
    vim.notify("[neoffice] Not a neoffice buffer", vim.log.levels.WARN)
    return
  end

  if _conceal_state[buf] and _conceal_state[buf].enabled then
    disable_conceal(buf)
    vim.notify("[neoffice] XML tags visible", vim.log.levels.INFO)
  else
    setup_conceal(buf)
    vim.notify("[neoffice] XML tags concealed (nowrap)", vim.log.levels.INFO)
  end
end

---Open (or reuse) a proxy buffer for direct XML editing
---@param orig_path      string
---@param text_path      string  path to temp .txt file
---@param para_map       table   paragraph mapping
---@param original_root  table   parsed XML root (for faster saves)
function M.open_proxy(orig_path, text_path, para_map, original_root)
  -- Create buffer and set name
  local buf = vim.api.nvim_create_buf(true, false)
  vim.api.nvim_buf_set_name(buf, "neoffice://" .. orig_path)

  -- Load XML content
  local lines = vim.fn.readfile(text_path)
  vim.api.nvim_buf_call(buf, function()
    local old_undolevels = vim.bo[buf].undolevels
    vim.bo[buf].undolevels = -1
    vim.api.nvim_buf_set_lines(buf, 0, -1, false, lines)
    vim.bo[buf].modified = false
    vim.bo[buf].undolevels = old_undolevels
  end)

  -- Set filetype to XML for base highlighting
  vim.api.nvim_set_option_value("filetype", "xml", { buf = buf })
  vim.api.nvim_set_option_value("bufhidden", "hide", { buf = buf })

  -- Store metadata
  _meta[buf] = {
    orig_path = orig_path,
    text_path = text_path,
    para_map = para_map,
    original_root = original_root,
  }

  -- Switch to the buffer so extmarks have a window context
  vim.api.nvim_set_current_buf(buf)

  -- 1. Setup Visuals (Concealing tags)
  setup_conceal(buf)

  -- 2. Initialize Logic Modules
  local tc = require("neoffice.track_changes")
  local comms = require("neoffice.comments")

  -- 3. Load Data & Draw Initial Signs
  -- We wrap this in schedule to ensure the buffer is fully "ready" in the UI
  vim.schedule(function()
    if vim.api.nvim_buf_is_valid(buf) then
      tc.load(buf, orig_path, lines)
      comms.load(orig_path)
      comms.draw_anchors(buf)
    end
  end)

  -- 4. Setup Keymaps
  M._setup_keymaps(buf, orig_path)

  -- 5. THE FIX: Reactive Refreshing
  -- This ensures signs move when you type and update when you save
  local refresh_group = vim.api.nvim_create_augroup("NeofficeRefresh_" .. buf, { clear = true })
  vim.api.nvim_create_autocmd({ "TextChanged", "TextChangedI", "BufWritePost" }, {
    group = refresh_group,
    buffer = buf,
    callback = function()
      -- Track changes refresh
      tc.refresh(buf)
      -- Comments refresh (re-scanning the buffer for tags)
      comms.draw_anchors(buf)
    end,
  })

  -- 6. Cleanup on buffer delete
  vim.api.nvim_create_autocmd("BufDelete", {
    group = refresh_group,
    buffer = buf,
    callback = function()
      _meta[buf] = nil
      _conceal_state[buf] = nil
    end,
  })

  return buf
end

---Return metadata for the current buffer, or nil.
function M.get_meta()
  return _meta[vim.api.nvim_get_current_buf()]
end

---Return metadata for a specific buffer handle.
function M.get_meta_for(buf)
  return _meta[buf]
end

function M.get_all_meta()
  return _meta
end

function M._setup_keymaps(buf, orig_path)
  local cfg = config.get()
  local o = { buffer = buf, silent = true, noremap = true }

  vim.keymap.set(
    "n",
    cfg.mappings.save,
    "<cmd>DocSave<CR>",
    vim.tbl_extend("force", o, { desc = "Save back to document" })
  )

  vim.keymap.set("n", "j", "gj", vim.tbl_extend("force", o, { desc = "Move down by visual line" }))
  vim.keymap.set("n", "k", "gk", vim.tbl_extend("force", o, { desc = "Move up by visual line" }))
  vim.keymap.set("v", "j", "gj", vim.tbl_extend("force", o, { desc = "Move down by visual line" }))
  vim.keymap.set("v", "k", "gk", vim.tbl_extend("force", o, { desc = "Move up by visual line" }))
  vim.keymap.set("n", "gj", "j", vim.tbl_extend("force", o, { desc = "Move down by logical line" }))
  vim.keymap.set("n", "gk", "k", vim.tbl_extend("force", o, { desc = "Move up by logical line" }))

  vim.keymap.set("n", cfg.mappings.toggle_comments, function()
    require("neoffice.comments").toggle(orig_path, buf)
  end, vim.tbl_extend("force", o, { desc = "Toggle comments panel" }))

  vim.keymap.set({ "n", "v" }, cfg.mappings.add_comment, function()
    require("neoffice.comments").add_comment(buf)
  end, vim.tbl_extend("force", o, { desc = "Add comment at cursor" }))

  vim.keymap.set(
    "n",
    cfg.mappings.show_changes,
    "<cmd>DocChanges<CR>",
    vim.tbl_extend("force", o, { desc = "Show track changes" })
  )

  vim.keymap.set(
    "n",
    "<leader>dt", -- or cfg.mappings.toggle_tags if you add it to config
    function()
      M.toggle_conceal()
    end,
    vim.tbl_extend("force", o, { desc = "Toggle XML tag concealing" })
  )

  vim.keymap.set(
    "n",
    "<leader>dr", -- or cfg.mappings.refresh_comments
    function()
      require("neoffice.comments").refresh_from_buffer()
    end,
    vim.tbl_extend("force", o, { desc = "Refresh comments from buffer" })
  )
  vim.keymap.set("n", "]c", function()
    require("neoffice.track_changes").next_change()
  end, vim.tbl_extend("force", o, { desc = "Jump to next track change" }))

  vim.keymap.set("n", "[c", function()
    require("neoffice.track_changes").prev_change()
  end, vim.tbl_extend("force", o, { desc = "Jump to previous track change" }))

  vim.keymap.set("n", "<leader>dv", function()
    require("neoffice.track_changes").select_change()
  end, vim.tbl_extend("force", o, { desc = "Select current track change" }))

  -- Visual mode accept/reject
  vim.keymap.set("v", cfg.mappings.accept_change, function()
    local tc = require("neoffice.track_changes")
    local ch = tc.change_at_cursor()
    if ch then
      vim.api.nvim_feedkeys(vim.api.nvim_replace_termcodes("<Esc>", true, false, true), "x", false)
      tc.accept(ch.id)
    else
      vim.notify("[neoffice] No change found", vim.log.levels.WARN)
    end
  end, vim.tbl_extend("force", o, { desc = "Accept selected track change" }))

  vim.keymap.set("v", cfg.mappings.reject_change, function()
    local tc = require("neoffice.track_changes")
    local ch = tc.change_at_cursor()
    if ch then
      vim.api.nvim_feedkeys(vim.api.nvim_replace_termcodes("<Esc>", true, false, true), "x", false)
      tc.reject(ch.id)
    else
      vim.notify("[neoffice] No change found", vim.log.levels.WARN)
    end
  end, vim.tbl_extend("force", o, { desc = "Reject selected track change" }))
end

return M
