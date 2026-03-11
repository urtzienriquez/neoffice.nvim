-- neoffice/init.lua
-- Main coordinator module for direct XML editing mode

local M = {}

local function cfg()
  return require("neoffice.config")
end
local function convert()
  return require("neoffice.convert")
end
local function buf_mod()
  return require("neoffice.buffer")
end
local function tc()
  return require("neoffice.track_changes")
end
local function comments()
  return require("neoffice.comments")
end

-- ── Setup ────────────────────────────────────────────────────────────────────

function M.setup(user_opts)
  cfg().setup(user_opts)
  tc().setup_highlights()
  comments().setup_highlights()
end

-- ── Open ─────────────────────────────────────────────────────────────────────

function M.open(path)
  path = path or vim.fn.expand("%:p")
  if not path or path == "" then
    vim.notify("[neoffice] No file specified", vim.log.levels.ERROR)
    return
  end

  local ext = (path:match("%.(%w+)$") or ""):lower()
  if ext ~= "docx" and ext ~= "odt" and ext ~= "doc" then
    vim.notify("[neoffice] Unsupported format: " .. ext, vim.log.levels.WARN)
    return
  end

  if not convert().check_deps() then
    return
  end

  -- Extract text using direct XML mode
  local text_path, para_map, original_root = convert().to_text(path)
  if not text_path then
    vim.notify("[neoffice] Failed to extract text from document", vim.log.levels.ERROR)
    return
  end

  -- --- CLEANED UP OPEN LOGIC ---
  -- All initial loading of markers and autocommands now happens inside open_proxy
  local proxy_buf = buf_mod().open_proxy(path, text_path, para_map, original_root)

  -- Handle optional UI elements
  if cfg().get().auto_open_comments then
    local data = require("neoffice.extractor").extract(path)
    if #(data.comments or {}) > 0 then
      vim.schedule(function()
        comments().toggle(path, proxy_buf)
      end)
    end
  end
end

-- ── Save ─────────────────────────────────────────────────────────────────────

function M.save()
  local meta = buf_mod().get_meta()
  if not meta then
    vim.notify("[neoffice] Not a neoffice buffer", vim.log.levels.WARN)
    return
  end

  -- Get live comments to re-inject
  local live_comments = comments().get_comments()

  -- Save using direct XML mode
  convert().from_text(meta.text_path, meta.orig_path, meta.para_map, meta.original_root, live_comments)
end

-- ── Track changes ────────────────────────────────────────────────────────────

function M.show_changes()
  tc().show_summary()
end

function M.accept(args)
  local meta = buf_mod().get_meta()
  if not meta then
    return
  end

  if args == "all" then
    tc().accept_all() -- No longer need to pass path, handled by current buffer
  else
    local ch = tc().change_at_cursor()
    if ch then
      tc().accept(ch.id)
    end
  end
end

function M.reject(args)
  local meta = buf_mod().get_meta()
  if not meta then
    return
  end

  if args == "all" then
    tc().reject_all()
  else
    local ch = tc().change_at_cursor()
    if ch then
      tc().reject(ch.id)
    end
  end
end

-- ── Comments ─────────────────────────────────────────────────────────────────

function M.toggle_comments()
  local meta = buf_mod().get_meta()
  local buf = vim.api.nvim_get_current_buf()
  comments().toggle(meta and meta.orig_path or nil, buf)
end

function M.add_comment()
  local buf = vim.api.nvim_get_current_buf()
  comments().add_comment(buf)
end

return M
