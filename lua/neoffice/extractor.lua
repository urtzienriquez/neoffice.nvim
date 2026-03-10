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

  local office_text = xml.find_first(root, "office:text") or xml.find_first(root, "office:body") or root

  for para_idx, para in ipairs(office_text.children or {}) do
    if para.tag ~= "text:p" and para.tag ~= "text:h" then
      goto continue
    end

    local anchor_parts = {}
    if para.text and para.text ~= "" then
      table.insert(anchor_parts, para.text)
    end
    for _, child in ipairs(para.children or {}) do
      if child.tag ~= "office:annotation" then
        local t = xml.inner_text(child)
        if t ~= "" then
          table.insert(anchor_parts, t)
        end
      end
    end
    local anchor = table.concat(anchor_parts, ""):match("^%s*(.-)%s*$")

    for _, child in ipairs(para.children or {}) do
      if child.tag == "office:annotation" then
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
    end

    ::continue::
  end

  return comments
end

-- ── ODT: inject annotations into content.xml string ──────────────────────────

function M.inject_annotations_odt(content_xml, comments)
  content_xml = content_xml:gsub("<office:annotation.-</office:annotation>", "")

  local function ann_xml(cm)
    local e = function(s)
      return (tostring(s or ""):gsub("&", "&amp;"):gsub("<", "&lt;"):gsub(">", "&gt;"))
    end

    local function get_initials(name)
      local initials = ""
      for word in (name or "U"):gmatch("%S+") do
        initials = initials .. word:sub(1, 1):upper()
      end
      return (initials ~= "" and initials:sub(1, 2) or "U")
    end

    local parts = {}

    table.insert(parts, string.format('<office:annotation office:name="%s">', e(cm.id)))
    table.insert(parts, string.format("  <dc:creator>%s</dc:creator>", e(cm.author or "nvim")))
    table.insert(
      parts,
      string.format("  <dc:date>%s</dc:date>", (cm.date ~= "" and cm.date) or os.date("!%Y-%m-%dT%H:%M:%S"))
    )
    table.insert(
      parts,
      string.format("  <meta:creator-initials>%s</meta:creator-initials>", e(get_initials(cm.author)))
    )

    for _, body_line in ipairs(vim.split(cm.text or "", "\n", { plain = true })) do
      table.insert(parts, string.format("  <text:p>%s</text:p>", e(body_line)))
    end
    table.insert(parts, "</office:annotation>")

    for _, reply in ipairs(cm.replies or {}) do
      table.insert(parts, string.format('<office:annotation loext:parent-name="%s" loext:resolved="false">', e(cm.id)))
      table.insert(parts, string.format("  <dc:creator>%s</dc:creator>", e(reply.author or "nvim")))
      table.insert(parts, string.format("  <dc:date>%s</dc:date>", reply.date or os.date("!%Y-%m-%dT%H:%M:%S")))
      table.insert(
        parts,
        string.format("  <meta:creator-initials>%s</meta:creator-initials>", e(get_initials(reply.author)))
      )

      table.insert(parts, "  <text:p>")
      local short_date = (cm.date or ""):sub(1, 16)
      table.insert(
        parts,
        string.format(
          '    <text:span text:style-name="Quoted">Reply to %s (%s)</text:span>',
          e(cm.author),
          e(short_date)
        )
      )
      table.insert(parts, "    <text:line-break/>")
      table.insert(parts, "    " .. e(reply.text or ""))
      table.insert(parts, "  </text:p>")
      table.insert(parts, "</office:annotation>")
    end

    return table.concat(parts, "\n")
  end

  local lines = vim.split(content_xml, "\n", { plain = true })
  local injections = {}
  local text_p_closes = {}
  local buf = {}

  for i, line in ipairs(lines) do
    table.insert(buf, line)
    if line:find("</text:p>", 1, true) then
      table.insert(text_p_closes, { idx = i, text = table.concat(buf, " ") })
      buf = {}
    end
  end

  local last_close = text_p_closes[#text_p_closes]

  for _, cm in ipairs(comments) do
    local target_idx = last_close and last_close.idx or #lines
    local anchor = cm.anchor
    if anchor and anchor ~= "" then
      local needle = anchor:sub(1, 30)
      for _, entry in ipairs(text_p_closes) do
        if entry.text:find(needle, 1, true) then
          target_idx = entry.idx
          break
        end
      end
    end
    injections[target_idx] = injections[target_idx] or {}
    table.insert(injections[target_idx], ann_xml(cm))
  end

  local result = {}
  for i, line in ipairs(lines) do
    local anns = injections[i]
    if anns then
      local close_pos = line:find("</text:p>", 1, true)
      if close_pos then
        local before_close = line:sub(1, close_pos - 1)
        local after_close = line:sub(close_pos)
        table.insert(result, before_close)
        local indent = line:match("^(%s*)") or "      "
        for _, a in ipairs(anns) do
          for _, ann_line in ipairs(vim.split(a, "\n", { plain = true })) do
            table.insert(result, indent .. ann_line)
          end
        end
        table.insert(result, indent:sub(1, -3) .. after_close)
      else
        table.insert(result, line)
      end
    else
      table.insert(result, line)
    end
  end
  return table.concat(result, "\n")
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
