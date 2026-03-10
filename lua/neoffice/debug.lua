-- neoffice/debug.lua
-- :DocDebug  –  opens a scratch buffer with a full diagnostic report

local zip = require("neoffice.zip")
local xml = require("neoffice.xml")
local extractor = require("neoffice.extractor")
local M = {}

function M.run(orig_path)
  orig_path = orig_path or (require("neoffice.buffer").get_meta() or {}).orig_path or vim.fn.expand("%:p")

  if not orig_path or orig_path == "" then
    vim.notify("[neoffice] DocDebug: no document path found", vim.log.levels.ERROR)
    return
  end

  local lines = {
    "# neoffice debug report",
    "# document: " .. orig_path,
    "# " .. os.date("%Y-%m-%d %H:%M:%S"),
    "",
  }

  table.insert(lines, "## ZIP entries")
  local entries = zip.list_entries(orig_path)
  if #entries == 0 then
    table.insert(lines, "  ERROR: no entries found – is unzip installed and is the path correct?")
    table.insert(lines, "  Tried: unzip -Z1 " .. orig_path)
  else
    for _, e in ipairs(entries) do
      table.insert(lines, "  " .. e)
    end
  end
  table.insert(lines, "")

  local ext = (orig_path:match("%.(%w+)$") or ""):lower()
  local comments_entry = ext == "odt" and "content.xml" or "word/comments.xml"

  table.insert(lines, "## " .. comments_entry .. "  (first 2000 chars)")
  local raw = zip.read_entry(orig_path, comments_entry)
  if not raw then
    table.insert(lines, "  NOT FOUND in archive.")
    table.insert(lines, "  This means the document has no comments,")
    table.insert(lines, "  OR the unzip command failed.")
    local test_cmd = "unzip -p '" .. orig_path:gsub("'", "'\"'\"'") .. "' '" .. comments_entry .. "' 2>&1 | head -5"
    local test_out = vim.fn.system(test_cmd)
    table.insert(lines, "  unzip test output: " .. (test_out ~= "" and test_out or "(empty)"))
  else
    table.insert(lines, "  LENGTH: " .. #raw .. " bytes")
    table.insert(lines, "")
    for _, l in ipairs(vim.split(raw:sub(1, 2000), "\n", { plain = true })) do
      table.insert(lines, "  " .. l)
    end
    if #raw > 2000 then
      table.insert(lines, "  ... (truncated)")
    end
  end
  table.insert(lines, "")

  if raw then
    table.insert(lines, "## XML parse tree (first 3000 chars)")
    local ok, tree = pcall(xml.parse, raw)
    if not ok then
      table.insert(lines, "  PARSE ERROR: " .. tostring(tree))
    else
      local dump = xml.dump(tree)
      for _, l in ipairs(vim.split(dump:sub(1, 3000), "\n", { plain = true })) do
        table.insert(lines, l)
      end
      if #dump > 3000 then
        table.insert(lines, "  ... (truncated)")
      end

      local tag = ext == "odt" and "text:annotation" or "w:comment"
      local found = xml.find_all(tree, tag)
      table.insert(lines, "")
      table.insert(lines, string.format("  find_all(%q) → %d node(s)", tag, #found))
    end
    table.insert(lines, "")
  end

  table.insert(lines, "## extractor.extract() result")
  local ok2, data = pcall(extractor.extract, orig_path)
  if not ok2 then
    table.insert(lines, "  ERROR: " .. tostring(data))
  else
    table.insert(lines, string.format("  track_changes: %d", #(data.track_changes or {})))
    table.insert(lines, string.format("  comments:      %d", #(data.comments or {})))
    table.insert(lines, "")
    for i, cm in ipairs(data.comments or {}) do
      table.insert(
        lines,
        string.format("  [%d] id=%s  author=%s  text=%s", i, cm.id or "?", cm.author or "?", (cm.text or ""):sub(1, 60))
      )
    end
  end
  table.insert(lines, "")
  table.insert(lines, "# end of report")

  local buf = vim.api.nvim_create_buf(false, true)
  vim.api.nvim_buf_set_lines(buf, 0, -1, false, lines)
  vim.api.nvim_set_option_value("filetype", "markdown", { buf = buf })
  vim.api.nvim_set_option_value("modifiable", false, { buf = buf })
  vim.api.nvim_buf_set_name(buf, "[neoffice:debug]")

  vim.cmd("botright 20split")
  local win = vim.api.nvim_get_current_win()
  vim.api.nvim_win_set_buf(win, buf)
  vim.keymap.set("n", "q", "<cmd>close<CR>", { buffer = buf, silent = true })
  vim.keymap.set("n", "<Esc>", "<cmd>close<CR>", { buffer = buf, silent = true })
end

return M
