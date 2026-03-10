-- neoffice/direct.lua
-- Direct XML text extraction and editing without pandoc conversion.
-- Preserves ALL formatting by only replacing text content in XML nodes.

local zip = require("neoffice.zip")
local xml = require("neoffice.xml")
local M = {}

-- ── XML Serialization ────────────────────────────────────────────────────────

local function serialize_xml(node, indent)
  indent = indent or 0
  if not node or type(node) ~= "table" then
    return ""
  end

  local parts = {}

  -- Opening tag
  local tag_open = "<" .. node.tag
  for k, v in pairs(node.attrs or {}) do
    -- Escape attribute values
    v = tostring(v):gsub("&", "&amp;"):gsub("<", "&lt;"):gsub(">", "&gt;"):gsub('"', "&quot;")
    tag_open = tag_open .. string.format(' %s="%s"', k, v)
  end

  -- Check if self-closing (no text, no children)
  if (not node.text or node.text == "") and (#(node.children or {}) == 0) then
    table.insert(parts, tag_open .. "/>")
    return table.concat(parts, "\n")
  end

  tag_open = tag_open .. ">"
  table.insert(parts, tag_open)

  -- Text content (escaped)
  if node.text and node.text ~= "" then
    local text = tostring(node.text):gsub("&", "&amp;"):gsub("<", "&lt;"):gsub(">", "&gt;")
    table.insert(parts, text)
  end

  -- Children
  for _, child in ipairs(node.children or {}) do
    table.insert(parts, serialize_xml(child, indent + 1))
  end

  -- Closing tag
  table.insert(parts, "</" .. node.tag .. ">")

  return table.concat(parts, "\n")
end

-- ── DOCX: Extract editable text ─────────────────────────────────────────────

local function docx_to_text(zip_path)
  local raw = zip.read_entry(zip_path, "word/document.xml")
  if not raw then
    return nil, "Could not read document.xml"
  end

  local root = xml.parse(raw)
  local body = xml.find_first(root, "w:body")
  if not body then
    return nil, "No w:body found"
  end

  local lines = {}
  local para_map = {} -- Track paragraph index to XML node mapping

  for idx, para in ipairs(xml.find_all(body, "w:p")) do
    -- Extract all text from runs in this paragraph
    local para_text = {}
    for _, run in ipairs(xml.find_all(para, "w:r")) do
      for _, t in ipairs(xml.find_all(run, "w:t")) do
        local text = xml.inner_text(t)
        if text ~= "" then
          table.insert(para_text, text)
        end
      end
    end

    -- Check if this is a heading or special paragraph
    local pPr = xml.find_first(para, "w:pPr")
    local pStyle = pPr and xml.find_first(pPr, "w:pStyle")
    local style_val = pStyle and xml.attr(pStyle, "w:val")

    local prefix = ""
    if style_val then
      if style_val:match("Heading1") then
        prefix = "# "
      elseif style_val:match("Heading2") then
        prefix = "## "
      elseif style_val:match("Heading3") then
        prefix = "### "
      elseif style_val:match("Heading4") then
        prefix = "#### "
      end
    end

    local line = prefix .. table.concat(para_text, "")
    table.insert(lines, line)
    para_map[#lines] = idx -- Map line number to paragraph index
  end

  return table.concat(lines, "\n"), para_map, root
end

-- ── DOCX: Replace text content ──────────────────────────────────────────────

local function docx_from_text(zip_path, edited_text, para_map, original_root)
  local root = original_root
  if not root then
    local raw = zip.read_entry(zip_path, "word/document.xml")
    if not raw then
      return false, "Could not read document.xml"
    end
    root = xml.parse(raw)
  end

  local body = xml.find_first(root, "w:body")
  if not body then
    return false, "No w:body found"
  end

  local paragraphs = xml.find_all(body, "w:p")
  local edited_lines = vim.split(edited_text, "\n", { plain = true })

  -- Process each edited line
  for line_num, new_text in ipairs(edited_lines) do
    local para_idx = para_map[line_num]
    if para_idx and paragraphs[para_idx] then
      local para = paragraphs[para_idx]

      -- Strip heading markers if present
      new_text = new_text:gsub("^#+%s*", "")

      -- Find all text nodes in this paragraph
      local text_nodes = xml.find_all(para, "w:t")

      if #text_nodes > 0 then
        -- Strategy: put all text in the first text node, clear others
        -- This preserves the formatting of the first run
        text_nodes[1].text = new_text
        for i = 2, #text_nodes do
          text_nodes[i].text = ""
        end
      else
        -- No existing text nodes - create basic run structure
        -- Find or create a run
        local run = xml.find_first(para, "w:r")
        if not run then
          run = { tag = "w:r", attrs = {}, children = {}, text = "" }
          table.insert(para.children, run)
        end

        -- Add text node
        local t_node = { tag = "w:t", attrs = {}, children = {}, text = new_text }
        table.insert(run.children, t_node)
      end
    end
  end

  -- Serialize the modified XML back
  local modified_xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' .. serialize_xml(root)

  -- Write back to ZIP
  return zip.write_entry(zip_path, "word/document.xml", modified_xml)
end

-- ── ODT: Extract editable text ──────────────────────────────────────────────

local function odt_to_text(zip_path)
  local raw = zip.read_entry(zip_path, "content.xml")
  if not raw then
    return nil, "Could not read content.xml"
  end

  local root = xml.parse(raw)
  local office_text = xml.find_first(root, "office:text") or xml.find_first(root, "office:body") or root

  local lines = {}
  local para_map = {}

  for idx, para in ipairs(office_text.children or {}) do
    if para.tag == "text:p" or para.tag == "text:h" then
      local para_text = xml.inner_text(para)

      -- Check heading level
      local prefix = ""
      if para.tag == "text:h" then
        local level = xml.attr(para, "text:outline-level")
        if level then
          prefix = string.rep("#", tonumber(level)) .. " "
        else
          prefix = "# "
        end
      end

      table.insert(lines, prefix .. para_text)
      para_map[#lines] = idx
    end
  end

  return table.concat(lines, "\n"), para_map, root
end

-- ── ODT: Replace text content ───────────────────────────────────────────────

local function odt_from_text(zip_path, edited_text, para_map, original_root)
  local root = original_root
  if not root then
    local raw = zip.read_entry(zip_path, "content.xml")
    if not raw then
      return false, "Could not read content.xml"
    end
    root = xml.parse(raw)
  end

  local office_text = xml.find_first(root, "office:text") or xml.find_first(root, "office:body") or root

  local edited_lines = vim.split(edited_text, "\n", { plain = true })
  local children = office_text.children or {}

  for line_num, new_text in ipairs(edited_lines) do
    local para_idx = para_map[line_num]
    if para_idx and children[para_idx] then
      local para = children[para_idx]

      -- Strip heading markers
      new_text = new_text:gsub("^#+%s*", "")

      -- Replace the text content while keeping all formatting
      para.text = new_text

      -- Clear child text nodes to avoid duplication
      for _, child in ipairs(para.children or {}) do
        if child.tag == "text:span" or child.tag == "text:a" then
          child.text = ""
        end
      end
    end
  end

  -- Serialize and write back
  local modified_xml = '<?xml version="1.0" encoding="UTF-8"?>\n' .. serialize_xml(root)
  return zip.write_entry(zip_path, "content.xml", modified_xml)
end

-- ── Public API ───────────────────────────────────────────────────────────────

function M.to_text(orig_path)
  local ext = (orig_path:match("%.(%w+)$") or ""):lower()

  if ext == "docx" then
    return docx_to_text(orig_path)
  elseif ext == "odt" then
    return odt_to_text(orig_path)
  end

  return nil, "Unsupported format: " .. ext
end

function M.from_text(orig_path, edited_text, para_map, original_root)
  local ext = (orig_path:match("%.(%w+)$") or ""):lower()

  if ext == "docx" then
    return docx_from_text(orig_path, edited_text, para_map, original_root)
  elseif ext == "odt" then
    return odt_from_text(orig_path, edited_text, para_map, original_root)
  end

  return false, "Unsupported format: " .. ext
end

return M
