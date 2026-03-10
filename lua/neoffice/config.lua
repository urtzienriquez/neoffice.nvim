-- neoffice/config.lua
-- Configuration for direct XML editing mode

local M = {}

M.defaults = {
  -- Automatically save back to document on :w
  auto_save = true,

  -- Show track-change extmarks when opening a document
  show_track_changes = true,

  -- Open the comments panel automatically when comments are found
  auto_open_comments = false,

  -- Temporary directory for proxy files (nil = vim.fn.tempname parent)
  tmp_dir = nil,

  mappings = {
    -- In the main proxy buffer
    save = "<leader>ds",
    toggle_comments = "<leader>dc",
    add_comment = "<leader>da",
    show_changes = "<leader>dT",
    accept_change = "<leader>dy",
    reject_change = "<leader>dn",

    -- Inside the comments panel
    reply_comment = "r",
    resolve_comment = "<CR>",
    delete_comment = "d",
    close_panel = "q",
  },
}

M.options = vim.deepcopy(M.defaults)

function M.setup(user_opts)
  M.options = vim.tbl_deep_extend("force", M.defaults, user_opts or {})
end

function M.get()
  return M.options
end

return M
