-- lua/neoffice/diagnose_annotations.lua
-- Diagnostic to inspect annotation XML structure in ODT files

local M = {}

function M.show_saved_annotations(odt_path)
  local zip = require("neoffice.zip")
  local xml = require("neoffice.xml")

  odt_path = odt_path or vim.fn.expand("%:p")
  if not odt_path:match("%.odt$") then
    vim.notify("[neoffice] Not an ODT file", vim.log.levels.ERROR)
    return
  end

  local content = zip.read_entry(odt_path, "content.xml")
  if not content then
    vim.notify("[neoffice] Could not read content.xml", vim.log.levels.ERROR)
    return
  end

  local root = xml.parse(content)
  local office_text = xml.find_first(root, "office:text") or xml.find_first(root, "office:body") or root

  local buf = vim.api.nvim_create_buf(false, true)
  local lines = {
    "=== ANNOTATION STRUCTURE IN " .. vim.fn.fnamemodify(odt_path, ":t") .. " ===",
    "",
  }

  local ann_count = 0
  local reply_count = 0

  for _, para in ipairs(office_text.children or {}) do
    if para.tag == "text:p" or para.tag == "text:h" then
      for _, child in ipairs(para.children or {}) do
        if child.tag == "office:annotation" then
          ann_count = ann_count + 1
          local ann_name = xml.attr(child, "office:name") or "?"
          local creator = xml.inner_text(xml.find_first(child, "dc:creator") or {}) or "?"

          table.insert(lines, string.format("Annotation #%d: name=%s author=%s", ann_count, ann_name, creator))

          local nested = 0
          for _, nested_child in ipairs(child.children or {}) do
            if nested_child.tag == "office:annotation" then
              nested = nested + 1
              reply_count = reply_count + 1
              local reply_name = xml.attr(nested_child, "office:name") or "?"
              local reply_creator = xml.inner_text(xml.find_first(nested_child, "dc:creator") or {}) or "?"
              table.insert(lines, string.format("  └─ Reply: name=%s author=%s", reply_name, reply_creator))
            end
          end

          if nested == 0 then
            table.insert(lines, "  (no nested annotations/replies)")
          end

          if nested > 0 and reply_count == nested then
            table.insert(lines, "")
            table.insert(lines, "  RAW XML:")
            local raw_ann = content:match(
              "<office:annotation[^>]*office:name=['\"]"
                .. ann_name:gsub("%-", "%%-"):gsub("%_", "%%_")
                .. "['\"][^>]*>.-</office:annotation>"
            )
            if raw_ann then
              for line in raw_ann:gmatch("[^\n]+") do
                table.insert(lines, "  " .. line)
              end
            end
          end

          table.insert(lines, "")
        end
      end
    end
  end

  table.insert(lines, string.format("SUMMARY: %d annotations, %d replies", ann_count, reply_count))
  table.insert(lines, "")

  if ann_count == 0 then
    table.insert(lines, "WARNING: No annotations found!")
    table.insert(lines, "The injection may have failed or pandoc removed them.")
  end

  vim.api.nvim_buf_set_lines(buf, 0, -1, false, lines)
  vim.api.nvim_set_option_value("filetype", "text", { buf = buf })
  vim.api.nvim_set_option_value("modifiable", false, { buf = buf })
  vim.api.nvim_buf_set_name(buf, "[annotations-dump]")

  vim.cmd("botright 30split")
  local win = vim.api.nvim_get_current_win()
  vim.api.nvim_win_set_buf(win, buf)
  vim.keymap.set("n", "q", "<cmd>close<CR>", { buffer = buf, silent = true })

  vim.notify(string.format("[neoffice] Found %d annotations, %d replies", ann_count, reply_count), vim.log.levels.INFO)
end

return M
