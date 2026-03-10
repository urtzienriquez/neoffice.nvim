-- neoffice/xml.lua
-- Robust XML tokeniser + tree builder for Office Open XML / ODF.
-- Handles: namespace prefixes, quoted attrs containing >, CDATA, self-closing tags.

local M = {}

-- ── Low-level tokeniser ───────────────────────────────────────────────────────

local function read_tag(s, pos)
  local len = #s
  local buf = {}
  local in_q = nil

  while pos <= len do
    local c = s:sub(pos, pos)
    if in_q then
      buf[#buf + 1] = c
      if c == in_q then
        in_q = nil
      end
    else
      if c == '"' or c == "'" then
        in_q = c
        buf[#buf + 1] = c
      elseif c == ">" then
        pos = pos + 1
        break
      else
        buf[#buf + 1] = c
      end
    end
    pos = pos + 1
  end

  return table.concat(buf), pos
end

local function parse_attrs(s)
  local attrs = {}
  local pos = 1
  local len = #s
  while pos <= len do
    local ws = s:match("^%s+", pos)
    if ws then
      pos = pos + #ws
    end
    if pos > len then
      break
    end

    local name = s:match("^([%w:%-_.]+)", pos)
    if not name then
      break
    end
    pos = pos + #name

    local eq = s:match("^%s*=%s*", pos)
    if eq then
      pos = pos + #eq
      local q = s:sub(pos, pos)
      if q == '"' or q == "'" then
        pos = pos + 1
        local val_end = s:find(q, pos, true)
        if val_end then
          attrs[name] = s:sub(pos, val_end - 1)
          pos = val_end + 1
        end
      else
        local val = s:match("^([^%s>]+)", pos) or ""
        attrs[name] = val
        pos = pos + #val
      end
    else
      attrs[name] = name
    end
  end
  return attrs
end

-- ── Tree builder ─────────────────────────────────────────────────────────────

function M.parse(src)
  src = src:gsub("^%s*<%?[^%?]*%?>", "")
  src = src:gsub("<!%-%-(.-)%-%->", "")
  src = src:gsub("<!%[CDATA%[(.-)%]%]>", function(c)
    return c
  end)
  src = src:gsub("<!DOCTYPE[^>]*>", "")

  local pos = 1
  local len = #src

  local function parse_node()
    if pos > len then
      return nil
    end

    local lt = src:find("<", pos, true)
    if not lt then
      local text = src:sub(pos):match("^%s*(.-)%s*$")
      pos = len + 1
      return text ~= "" and { tag = "#text", text = text, attrs = {}, children = {} } or nil
    end

    if lt > pos then
      local text = src:sub(pos, lt - 1):match("^%s*(.-)%s*$")
      pos = lt
      if text ~= "" then
        return { tag = "#text", text = text, attrs = {}, children = {} }
      end
    end

    local c2 = src:sub(pos + 1, pos + 1)

    if c2 == "/" then
      return nil
    end

    if c2 == "?" then
      local e = src:find("?>", pos + 2, true)
      pos = e and e + 2 or len + 1
      return parse_node()
    end

    pos = pos + 1
    local raw, new_pos = read_tag(src, pos)
    pos = new_pos

    local self_closing = raw:sub(-1) == "/"
    if self_closing then
      raw = raw:sub(1, -2)
    end

    local tag_name = raw:match("^([%w:%-_.]+)")
    if not tag_name then
      return parse_node()
    end

    local attr_str = raw:sub(#tag_name + 1)
    local node = {
      tag = tag_name,
      attrs = parse_attrs(attr_str),
      children = {},
      text = "",
    }

    if self_closing then
      return node
    end

    while pos <= len do
      local close_pat = "^</%s*" .. tag_name:gsub("([%-:.])", "%%%1") .. "%s*>"
      local cm = src:match(close_pat, pos)
      if cm then
        pos = pos + #cm
        break
      end

      local child = parse_node()
      if child == nil then
        local closing_end = src:find(">", pos, true)
        pos = closing_end and closing_end + 1 or len + 1
        break
      end
      if child.tag == "#text" then
        node.text = (node.text ~= "" and node.text .. " " or "") .. child.text
      else
        table.insert(node.children, child)
      end
    end

    return node
  end

  return parse_node() or { tag = "root", attrs = {}, children = {}, text = "" }
end

-- ── Query helpers ─────────────────────────────────────────────────────────────

function M.find_all(node, tag, acc)
  acc = acc or {}
  if not node or type(node) ~= "table" then
    return acc
  end
  if node.tag == tag then
    table.insert(acc, node)
  end
  for _, child in ipairs(node.children or {}) do
    M.find_all(child, tag, acc)
  end
  return acc
end

function M.find_first(node, tag)
  return M.find_all(node, tag)[1]
end

function M.inner_text(node)
  if not node then
    return ""
  end
  local parts = {}
  if node.text and node.text ~= "" then
    parts[#parts + 1] = node.text
  end
  for _, child in ipairs(node.children or {}) do
    local t = M.inner_text(child)
    if t ~= "" then
      parts[#parts + 1] = t
    end
  end
  return table.concat(parts, "")
end

function M.attr(node, key)
  if not node or not node.attrs then
    return nil
  end
  if node.attrs[key] then
    return node.attrs[key]
  end
  local local_key = key:match(":(.+)$") or key
  for k, v in pairs(node.attrs) do
    local lk = k:match(":(.+)$") or k
    if lk == local_key then
      return v
    end
  end
  return nil
end

function M.dump(node, indent)
  indent = indent or 0
  if not node or type(node) ~= "table" then
    return ""
  end
  local pad = string.rep("  ", indent)
  local line = pad .. "<" .. (node.tag or "?")
  for k, v in pairs(node.attrs or {}) do
    line = line .. string.format(" %s=%q", k, v)
  end
  line = line .. ">"
  if node.text and node.text ~= "" then
    line = line .. " [" .. node.text:sub(1, 40) .. "]"
  end
  local parts = { line }
  for _, child in ipairs(node.children or {}) do
    parts[#parts + 1] = M.dump(child, indent + 1)
  end
  return table.concat(parts, "\n")
end

return M
