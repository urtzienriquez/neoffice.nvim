-- neoffice/extractor.lua
-- Reads track changes and comments from .docx / .odt files.
-- Also provides write-back functions used by convert.lua.

local zip = require("neoffice.zip")
local xml = require("neoffice.xml")
local M = {}

-- ── DOCX: track changes ───────────────────────────────────────────────────────

local function docx_track_changes(zip_path)
  local raw = zip.read_entry(zip_path, "word/document.xml")
  if not raw then
    return {}
  end

  local root = xml.parse(raw)
  local changes = {}
  local id = 0

  local paras = xml.find_all(root, "w:p")
  for para_idx, para in ipairs(paras) do
    for _, child in ipairs(para.children or {}) do
      local tag = child.tag
      if tag == "w:ins" or tag == "w:del" then
        local text_tag = tag == "w:del" and "w:delText" or "w:t"
        local text = ""
        for _, t in ipairs(xml.find_all(child, text_tag)) do
          text = text .. xml.inner_text(t)
        end
        table.insert(changes, {
          id = (tag == "w:ins" and "ins_" or "del_") .. id,
          type = tag == "w:ins" and "insert" or "delete",
          author = xml.attr(child, "w:author") or "?",
          date = xml.attr(child, "w:date") or "",
          text = text,
          para = para_idx,
          status = "pending",
        })
        id = id + 1
      end
    end
  end
  return changes
end

-- ── DOCX: comments ────────────────────────────────────────────────────────────

local function docx_comments(zip_path)
  local raw = zip.read_entry(zip_path, "word/comments.xml")
  if not raw then
    return {}
  end

  local root = xml.parse(raw)
  local comments = {}

  for _, c in ipairs(xml.find_all(root, "w:comment")) do
    local text_parts = {}
    for _, t in ipairs(xml.find_all(c, "w:t")) do
      table.insert(text_parts, xml.inner_text(t))
    end
    table.insert(comments, {
      id = xml.attr(c, "w:id") or tostring(#comments),
      author = xml.attr(c, "w:author") or "?",
      date = xml.attr(c, "w:date") or "",
      text = table.concat(text_parts, " "),
      replies = {},
      resolved = false,
      anchor = nil,
    })
  end

  local ext_raw = zip.read_entry(zip_path, "word/commentsExtended.xml")
  if ext_raw then
    local ext_root = xml.parse(ext_raw)
    for _, ce in ipairs(xml.find_all(ext_root, "w15:commentEx")) do
      if xml.attr(ce, "w15:done") == "1" then
        local pid = xml.attr(ce, "w15:paraId")
        for _, cm in ipairs(comments) do
          if cm.id == pid then
            cm.resolved = true
          end
        end
      end
    end
  end

  return comments
end

-- ── ODT: track changes ────────────────────────────────────────────────────────

local function odt_track_changes(zip_path)
  local raw = zip.read_entry(zip_path, "content.xml")
  if not raw then
    return {}
  end

  local root = xml.parse(raw)
  local changes = {}
  local id = 0

  for _, region in ipairs(xml.find_all(root, "text:changed-region")) do
    local rid = xml.attr(region, "text:id") or tostring(id)

    for _, ins in ipairs(xml.find_all(region, "text:insertion")) do
      local info = xml.find_first(ins, "office:change-info")
      local author = info and xml.inner_text(xml.find_first(info, "dc:creator") or {}) or "?"
      local date = info and xml.inner_text(xml.find_first(info, "dc:date") or {}) or ""
      table.insert(changes, {
        id = "ins_" .. rid,
        type = "insert",
        author = author,
        date = date,
        text = "",
        para = 0,
        status = "pending",
      })
    end

    for _, del in ipairs(xml.find_all(region, "text:deletion")) do
      local info = xml.find_first(del, "office:change-info")
      local author = info and xml.inner_text(xml.find_first(info, "dc:creator") or {}) or "?"
      local date = info and xml.inner_text(xml.find_first(info, "dc:date") or {}) or ""
      table.insert(changes, {
        id = "del_" .. rid,
        type = "delete",
        author = author,
        date = date,
        text = xml.inner_text(del),
        para = 0,
        status = "pending",
      })
    end

    id = id + 1
  end

  return changes
end

-- ── ODT: comments (read) ──────────────────────────────────────────────────────

local function odt_comments(zip_path)
  local raw = zip.read_entry(zip_path, "content.xml")
  if not raw then
    return {}
  end

  local root = xml.parse(raw)
  local comments = {}

  local function extract_replies(parent_node)
    local replies = {}
    for _, child in ipairs(parent_node.children or {}) do
      if child.tag == "office:annotation" then
        local reply_body = {}
        for _, tp in ipairs(child.children or {}) do
          if tp.tag == "text:p" then
            local t = xml.inner_text(tp)
            if t ~= "" then
              table.insert(reply_body, t)
            end
          end
        end
        if #reply_body == 0 and child.text and child.text ~= "" then
          table.insert(reply_body, child.text)
        end
        table.insert(replies, {
          author = xml.inner_text(xml.find_first(child, "dc:creator") or {}) or "?",
          date = xml.inner_text(xml.find_first(child, "dc:date") or {}) or "",
          text = table.concat(reply_body, "\n"),
        })
      end
    end
    return replies
  end

  -- NEW: Recursive function to find all annotations in a node
  local function find_annotations_in_node(node)
    local found = {}
    if node.tag == "office:annotation" then
      table.insert(found, node)
    end
    for _, child in ipairs(node.children or {}) do
      local nested = find_annotations_in_node(child)
      for _, ann in ipairs(nested) do
        table.insert(found, ann)
      end
    end
    return found
  end

  local office_text = xml.find_first(root, "office:text") or xml.find_first(root, "office:body") or root

  for para_idx, para in ipairs(office_text.children or {}) do
    if para.tag ~= "text:p" and para.tag ~= "text:h" then
      goto continue
    end

    -- Extract XML anchor to match buffer content
    local anchor = xml.serialize_inner(para):sub(1, 60)

    -- NEW: Find all annotations recursively, not just direct children
    local annotations = find_annotations_in_node(para)

    for _, child in ipairs(annotations) do
      local body_parts = {}
      for _, tp in ipairs(child.children or {}) do
        if tp.tag == "text:p" then
          local t = xml.inner_text(tp)
          if t ~= "" then
            table.insert(body_parts, t)
          end
        end
      end
      if #body_parts == 0 and child.text and child.text ~= "" then
        table.insert(body_parts, child.text)
      end

      local replies = extract_replies(child)

      table.insert(comments, {
        id = xml.attr(child, "office:name") or tostring(#comments + 1),
        author = xml.inner_text(xml.find_first(child, "dc:creator") or {}) or "?",
        date = xml.inner_text(xml.find_first(child, "dc:date") or {}) or "",
        text = table.concat(body_parts, "\n"),
        replies = replies,
        resolved = false,
        anchor = anchor ~= "" and anchor or nil,
      })
    end

    ::continue::
  end

  return comments
end

-- ── ODT: inject annotations into content.xml string ──────────────────────────

function M.inject_annotations_odt(content_xml, comments)
  -- Parse the XML properly
  local root = xml.parse(content_xml)

  -- Find the office:text container
  local office_text = xml.find_first(root, "office:text") or xml.find_first_local(root, "text") or root

  -- Helper to create annotation node
  local function create_annotation(cm)
    local ann = {
      tag = "office:annotation",
      attrs = { ["office:name"] = cm.id },
      children = {},
    }

    -- Add metadata
    table.insert(ann.children, {
      tag = "dc:creator",
      attrs = {},
      children = { { tag = "_TEXT", text = cm.author or "nvim" } },
    })

    table.insert(ann.children, {
      tag = "dc:date",
      attrs = {},
      children = { { tag = "_TEXT", text = cm.date or os.date("!%Y-%m-%dT%H:%M:%SZ") } },
    })

    -- Add comment text as paragraph
    table.insert(ann.children, {
      tag = "text:p",
      attrs = {},
      children = { { tag = "_TEXT", text = cm.text or "" } },
    })

    return ann
  end

  -- Find paragraphs and inject comments
  for _, cm in ipairs(comments) do
    local anchor = cm.anchor
    if anchor and anchor ~= "" then
      local needle = anchor:sub(1, 30)

      -- Search through all paragraphs
      for _, para in ipairs(office_text.children or {}) do
        local tag_local = (para.tag or ""):match(":(.+)$") or para.tag
        if tag_local == "p" or tag_local == "h" then
          -- Check if this paragraph matches the anchor
          local para_text = xml.serialize_inner(para)
          if para_text:find(needle, 1, true) then
            -- Insert annotation as a child of this paragraph
            table.insert(para.children, create_annotation(cm))

            -- Add replies if any
            for _, reply in ipairs(cm.replies or {}) do
              local reply_ann = create_annotation({
                id = cm.id .. "_reply",
                author = reply.author,
                date = reply.date,
                text = reply.text,
              })
              table.insert(para.children, reply_ann)
            end

            break -- Found the anchor, move to next comment
          end
        end
      end
    end
  end

  -- Serialize back to XML
  return xml.serialize(root)
end

-- ── DOCX: write comments.xml ──────────────────────────────────────────────────

function M.write_comments_docx(zip_path, comments)
  local W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  local lines = {
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    string.format('<w:comments xmlns:w="%s">', W),
  }

  for _, cm in ipairs(comments) do
    local e = function(s)
      return tostring(s):gsub("&", "&amp;"):gsub("<", "&lt;"):gsub(">", "&gt;")
    end
    table.insert(
      lines,
      string.format(
        '  <w:comment w:id="%s" w:author="%s" w:date="%s">',
        e(cm.id),
        e(cm.author or "nvim"),
        cm.date or os.date("!%Y-%m-%dT%H:%M:%SZ")
      )
    )
    table.insert(lines, "    <w:p><w:r>")
    table.insert(lines, "      <w:t>" .. e(cm.text or "") .. "</w:t>")
    table.insert(lines, "    </w:r></w:p>")
    for _, reply in ipairs(cm.replies or {}) do
      table.insert(
        lines,
        string.format(
          '  <w:comment w:id="%s_r" w:author="%s" w:date="%s">',
          e(cm.id),
          e(reply.author or "nvim"),
          reply.date or os.date("!%Y-%m-%dT%H:%M:%SZ")
        )
      )
      table.insert(lines, "    <w:p><w:r>")
      table.insert(lines, "      <w:t>" .. e(reply.text or "") .. "</w:t>")
      table.insert(lines, "    </w:r></w:p>")
      table.insert(lines, "  </w:comment>")
    end
    table.insert(lines, "  </w:comment>")
  end

  table.insert(lines, "</w:comments>")
  return zip.write_entry(zip_path, "word/comments.xml", table.concat(lines, "\n"))
end

-- ── Public API ────────────────────────────────────────────────────────────────

function M.extract(path)
  local ext = (path:match("%.(%w+)$") or ""):lower()
  if ext == "docx" or ext == "doc" then
    return {
      track_changes = docx_track_changes(path),
      comments = docx_comments(path),
    }
  elseif ext == "odt" then
    return {
      track_changes = odt_track_changes(path),
      comments = odt_comments(path),
    }
  end
  return { track_changes = {}, comments = {} }
end

return M
