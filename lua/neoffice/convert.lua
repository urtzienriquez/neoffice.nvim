-- neoffice/convert.lua
-- Direct XML editing - preserves ALL formatting

local config = require("neoffice.config")
local M = {}

-- ── Direct XML editing ───────────────────────────────────────────────────────

function M.to_text(orig_path)
  local direct = require("neoffice.direct")
  local text, para_map, root = direct.to_text(orig_path)

  if not text then
    vim.notify("[neoffice] " .. (para_map or "Failed to extract text"), vim.log.levels.ERROR)
    return nil
  end

  -- Write to temp file
  local tmp = vim.fn.tempname() .. ".txt"
  vim.fn.writefile(vim.split(text, "\n", { plain = true }), tmp)

  return tmp, para_map, root
end

function M.from_text(text_path, orig_path, para_map, original_root, comments)
  local direct = require("neoffice.direct")

  -- Read edited text
  local edited_text = table.concat(vim.fn.readfile(text_path), "\n")

  -- Replace text in original XML
  local ok, err = direct.from_text(orig_path, edited_text, para_map, original_root)

  if not ok then
    vim.notify("[neoffice] Save failed: " .. (err or "unknown error"), vim.log.levels.ERROR)
    return
  end

  -- Re-inject comments if needed
  if comments and #comments > 0 then
    local ext = (orig_path:match("%.(%w+)$") or ""):lower()
    if ext == "odt" then
      -- For ODT, inject comments into content.xml
      local zip = require("neoffice.zip")
      local extractor = require("neoffice.extractor")
      local content = zip.read_entry(orig_path, "content.xml")
      if content then
        local merged = extractor.inject_annotations_odt(content, comments)
        zip.write_entry(orig_path, "content.xml", merged)
      end
    else
      -- For DOCX, write comments.xml
      require("neoffice.extractor").write_comments_docx(orig_path, comments)
    end
  end

  vim.notify("[neoffice] Saved → " .. vim.fn.fnamemodify(orig_path, ":t"), vim.log.levels.INFO)
end

-- ── Dependency check ─────────────────────────────────────────────────────────

function M.check_deps()
  if vim.fn.executable("unzip") ~= 1 then
    vim.notify("[neoffice] unzip not found (required for reading documents)", vim.log.levels.ERROR)
    return false
  end
  if vim.fn.executable("zip") ~= 1 then
    vim.notify("[neoffice] zip not found (required for saving documents)", vim.log.levels.ERROR)
    return false
  end
  return true
end

return M
