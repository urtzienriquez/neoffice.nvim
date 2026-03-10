-- neoffice/buffer.lua
-- Creates and manages the proxy buffer for direct XML editing

local config = require("neoffice.config")
local M = {}

-- buf  →  { orig_path, text_path, para_map, original_root }
local _meta = {}

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

  vim.api.nvim_buf_set_lines(buf, 0, -1, false, lines)

  local ext = (orig_path:match("%.(%w+)$") or ""):lower()
  local display = string.format("[%s] %s", ext, vim.fn.fnamemodify(orig_path, ":t"))
  vim.api.nvim_buf_set_name(buf, display)

  vim.api.nvim_set_option_value("filetype", "text", { buf = buf })
  vim.api.nvim_set_option_value("buftype", "acwrite", { buf = buf })
  vim.api.nvim_set_option_value("modified", false, { buf = buf })

  _meta[buf] = {
    orig_path = orig_path,
    text_path = text_path,
    para_map = para_map,
    original_root = original_root, -- Cache for faster saves
  }

  vim.api.nvim_set_current_buf(buf)

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

  vim.notify(
    string.format(
      "[neoffice] %s opened",
      vim.fn.fnamemodify(orig_path, ":t")
    ),
    vim.log.levels.INFO
  )

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

function M._setup_keymaps(buf, orig_path)
  local cfg = config.get()
  local o = { buffer = buf, silent = true, noremap = true }

  vim.keymap.set(
    "n",
    cfg.mappings.save,
    "<cmd>DocSave<CR>",
    vim.tbl_extend("force", o, { desc = "Save back to document" })
  )

  vim.keymap.set("n", cfg.mappings.toggle_comments, function()
    require("neoffice.comments").toggle(orig_path, buf)
  end, vim.tbl_extend("force", o, { desc = "Toggle comments panel" }))

  vim.keymap.set("n", cfg.mappings.add_comment, function()
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
    cfg.mappings.accept_change,
    "<cmd>DocAccept<CR>",
    vim.tbl_extend("force", o, { desc = "Accept change at cursor" })
  )

  vim.keymap.set(
    "n",
    cfg.mappings.reject_change,
    "<cmd>DocReject<CR>",
    vim.tbl_extend("force", o, { desc = "Reject change at cursor" })
  )
end

return M
