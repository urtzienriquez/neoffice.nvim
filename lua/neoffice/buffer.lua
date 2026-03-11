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
  local cfg = config.get()

  -- Reuse existing buffer
  for buf, meta in pairs(_meta) do
    if meta.orig_path == orig_path and vim.api.nvim_buf_is_valid(buf) then
      vim.api.nvim_set_current_buf(buf)
      return buf
    end
  end

  local lines = vim.fn.readfile(text_path)
  local buf = vim.api.nvim_create_buf(true, false)

  -- Disable undo while setting initial content
  vim.api.nvim_buf_call(buf, function()
    local old_undolevels = vim.bo[buf].undolevels
    vim.bo[buf].undolevels = -1
    vim.api.nvim_buf_set_lines(buf, 0, -1, false, lines)
    vim.bo[buf].undolevels = old_undolevels
  end)

  local ext = (orig_path:match("%.(%w+)$") or ""):lower()
  local display = string.format("[%s] %s", ext, vim.fn.fnamemodify(orig_path, ":t"))
  vim.api.nvim_buf_set_name(buf, display)

  vim.api.nvim_set_option_value("filetype", "xml", { buf = buf })
  vim.api.nvim_set_option_value("buftype", "acwrite", { buf = buf })
  vim.api.nvim_set_option_value("modified", false, { buf = buf })

  _meta[buf] = {
    orig_path = orig_path,
    text_path = text_path,
    para_map = para_map,
    original_root = original_root, -- Cache for faster saves
  }

  vim.api.nvim_set_current_buf(buf)

  if cfg.conceal_tags_on_open then
    setup_conceal(buf)
  end

  -- :w  →  save back to document
  if cfg.auto_save then
    vim.api.nvim_create_autocmd("BufWriteCmd", {
      buffer = buf,
      callback = function()
        local current = vim.api.nvim_buf_get_lines(buf, 0, -1, false)
        vim.fn.writefile(current, text_path)
        require("neoffice").save()
        vim.api.nvim_set_option_value("modified", false, { buf = buf })
      end,
      desc = "neoffice: :w saves back to original document",
    })
  end

  -- Set up keymaps
  M._setup_keymaps(buf, orig_path)

  -- Cleanup on wipeout
  vim.api.nvim_create_autocmd("BufWipeout", {
    buffer = buf,
    once = true,
    callback = function()
      vim.fn.delete(text_path)
      _meta[buf] = nil
    end,
    desc = "neoffice: cleanup temp file on buffer close",
  })

  -- After creating the buffer, add change tracking
  vim.api.nvim_create_autocmd("BufWritePost", {
    buffer = buf,
    callback = function()
      -- Refresh comments after save
      require("neoffice.comments").refresh_from_buffer()
    end,
    desc = "neoffice: refresh comments after save",
  })

  vim.notify(string.format("[neoffice] %s opened", vim.fn.fnamemodify(orig_path, ":t")), vim.log.levels.INFO)

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
