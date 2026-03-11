-- neoffice/direct.lua
local zip = require("neoffice.zip")
local xml = require("neoffice.xml")
local M = {}

local function odt_to_text(zip_path)
  local raw = zip.read_entry(zip_path, "content.xml")
  if not raw then
    return nil, "Could not read content.xml"
  end

  local root = xml.parse(raw)
  local body = xml.find_first_local(root, "body")
  if not body then
    return nil, "Could not find body element"
  end

  local office_text = xml.find_first_local(body, "text") or body

  local lines = {}
  local para_map = {}

  for i, node in ipairs(office_text.children or {}) do
    -- Double check node is a table before processing
    if type(node) == "table" then
      local tag_local = (node.tag or ""):match(":(.+)$") or node.tag
      if tag_local == "p" or tag_local == "h" then
        local content = xml.serialize_inner(node)
        -- Ensure we don't return an empty string (Neovim needs at least a space for the line)
        if content == "" then
          content = " "
        end

        lines[#lines + 1] = content
        para_map[#lines] = i
      end
    end
  end

  if #lines == 0 then
    return nil, "No paragraphs found"
  end
  return table.concat(lines, "\n"), para_map, root
end

local function docx_to_text(zip_path)
  local raw = zip.read_entry(zip_path, "word/document.xml")
  if not raw then
    return nil, "Could not read document.xml"
  end

  local root = xml.parse(raw)
  local body = xml.find_first_local(root, "body")
  if not body then
    return nil, "Could not find body element in DOCX"
  end

  local lines = {}
  local para_map = {}
  for i, node in ipairs(body.children or {}) do
    local tag_local = node.tag:match(":(.+)$") or node.tag
    if tag_local == "p" then
      table.insert(lines, xml.serialize_inner(node))
      para_map[#lines] = i
    end
  end
  return table.concat(lines, "\n"), para_map, root
end

function M.from_text(zip_path, edited_text, para_map, root)
  local ext = (zip_path:match("%.(%w+)$") or ""):lower()

  -- Use flexible find to locate the container for saving
  local container_tag = (ext == "odt") and "text" or "body"
  local container = xml.find_first_local(root, container_tag)

  if not container then
    container = root
  end

  local children = container.children
  local edited_lines = vim.split(edited_text, "\n", { plain = true })

  for line_num, new_inner_xml in ipairs(edited_lines) do
    local child_idx = para_map[line_num]
    if child_idx and children[child_idx] then
      local fragment = xml.parse("<tmp>" .. new_inner_xml .. "</tmp>")
      -- The children of <tmp> are the nodes we want
      children[child_idx].children = fragment.children[1].children
    end
  end

  local final_xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' .. xml.serialize(root)
  local entry = (ext == "odt") and "content.xml" or "word/document.xml"
  return zip.write_entry(zip_path, entry, final_xml)
end

function M.to_text(orig_path)
  local ext = (orig_path:match("%.(%w+)$") or ""):lower()
  if ext == "odt" then
    return odt_to_text(orig_path)
  end
  if ext == "docx" then
    return docx_to_text(orig_path)
  end
  return nil, "Unsupported format"
end

return M
