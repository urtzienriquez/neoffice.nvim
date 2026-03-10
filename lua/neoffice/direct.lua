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

  -- Remove comment range markers and other non-content elements
  raw = raw:gsub("<w:commentRangeStart[^>]*/>", "")
  raw = raw:gsub("<w:commentRangeEnd[^>]*/>", "")
  raw = raw:gsub("<w:commentReference[^>]*/>", "")

  -- Parse for structure (root return value)
  local root = xml.parse(raw)

  local lines = {}
  local para_map = {}

  -- Process all w:p elements in document order from raw XML
  local pos = 1
  local para_idx = 0

  while pos <= #raw do
    -- Find next paragraph (including self-closing)
    local p_start = raw:find("<w:p[%s>/]", pos)
    if not p_start then
      break
    end

    para_idx = para_idx + 1

    -- Check for self-closing paragraph (rare but possible)
    local self_closing = raw:match("^<w:p[^>]*/>", p_start)
    if self_closing then
      table.insert(lines, "")
      para_map[#lines] = para_idx
      pos = p_start + #self_closing
      goto continue
    end

    -- Extract paragraph content
    local open_tag, content = raw:match("(<w:p[^>]*>)(.-)</w:p>", p_start)
    if not open_tag then
      pos = p_start + 1
      goto continue
    end

    -- Check for heading style
    local prefix = ""
    if content:match('<w:pStyle[^>]*w:val="Heading1"') then
      prefix = "# "
    elseif content:match('<w:pStyle[^>]*w:val="Heading2"') then
      prefix = "## "
    elseif content:match('<w:pStyle[^>]*w:val="Heading3"') then
      prefix = "### "
    elseif content:match('<w:pStyle[^>]*w:val="Heading4"') then
      prefix = "#### "
    end

    -- Replace XML tags with spaces to preserve word boundaries
    -- BUT: self-closing tags (like bookmarks, breaks) should not add spaces

    -- First remove self-closing tags without adding space
    local text = content:gsub("<[^>]+/>", "")

    -- Then replace remaining tags (open/close pairs) with spaces
    -- This prevents "<w:t>word1</w:t><w:t>word2</w:t>" from becoming "word1word2"
    text = text:gsub("<[^>]+>", " ")

    -- Decode XML entities
    text = text:gsub("&lt;", "<"):gsub("&gt;", ">"):gsub("&amp;", "&"):gsub("&quot;", '"'):gsub("&apos;", "'")

    -- Normalize whitespace: collapse multiple spaces/newlines into single space
    text = text:gsub("%s+", " ")

    -- Trim leading/trailing whitespace
    text = text:match("^%s*(.-)%s*$") or ""

    table.insert(lines, prefix .. text)
    para_map[#lines] = para_idx

    -- Move past this paragraph
    local close_pos = raw:find("</w:p>", p_start, true)
    pos = close_pos and (close_pos + 6) or (p_start + 1)

    ::continue::
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

  -- First, remove all annotation blocks (comments) from raw XML
  raw = raw:gsub("<office:annotation[^>]*>.-</office:annotation>", "")
  raw = raw:gsub("<office:annotation%-end[^>]*/>", "")

  -- Parse to get structure (for root return value)
  local root = xml.parse(raw)

  local lines = {}
  local para_map = {}

  -- Process all text:p and text:h elements in document order
  local pos = 1
  local para_idx = 0

  while pos <= #raw do
    -- Look for next paragraph or heading (including self-closing)
    local p_start = raw:find("<text:p[%s>/]", pos)
    local h_start = raw:find("<text:h[%s>/]", pos)

    local next_start, tag_type
    if p_start and h_start then
      if p_start < h_start then
        next_start, tag_type = p_start, "p"
      else
        next_start, tag_type = h_start, "h"
      end
    elseif p_start then
      next_start, tag_type = p_start, "p"
    elseif h_start then
      next_start, tag_type = h_start, "h"
    else
      break -- No more paragraphs or headings
    end

    para_idx = para_idx + 1

    -- Extract the element (handle both regular and self-closing tags)
    local pattern, close_tag
    if tag_type == "p" then
      pattern = "(<text:p[^>]*>)(.-)(</text:p>)"
      close_tag = "</text:p>"
    else
      pattern = "(<text:h[^>]*>)(.-)(</text:h>)"
      close_tag = "</text:h>"
    end

    -- First check for self-closing tag
    local self_closing = raw:match("^<text:[ph][^>]*/>", next_start)
    if self_closing then
      -- Self-closing tag means empty paragraph
      table.insert(lines, "")
      para_map[#lines] = para_idx
      pos = next_start + #self_closing
      goto continue
    end

    local open_tag, content, end_tag_match = raw:match(pattern, next_start)
    if not open_tag then
      pos = next_start + 1
      goto continue
    end

    -- For headings, extract the level
    local prefix = ""
    if tag_type == "h" then
      local level = open_tag:match('text:outline%-level="(%d+)"') or "1"
      prefix = string.rep("#", tonumber(level)) .. " "
    end

    -- Remove all XML tags without inserting spaces.
    -- Tags often split words due to formatting spans.    
    local text = content:gsub("<[^>]+>", "")

    -- Decode XML entities
    text = text:gsub("&lt;", "<"):gsub("&gt;", ">"):gsub("&amp;", "&"):gsub("&quot;", '"'):gsub("&apos;", "'")

    -- Normalize whitespace: collapse multiple spaces/newlines/tabs into single space
    text = text:gsub("%s+", " ")

    -- Trim leading/trailing whitespace
    text = text:match("^%s*(.-)%s*$") or ""

    -- Always add the line (even if empty) to preserve blank lines
    table.insert(lines, prefix .. text)
    para_map[#lines] = para_idx

    -- Move past this element
    local close_pos = raw:find(close_tag, next_start, true)
    pos = close_pos and (close_pos + #close_tag) or (next_start + 1)

    ::continue::
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
