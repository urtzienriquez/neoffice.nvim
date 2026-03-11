-- neoffice/xml.lua
local M = {}

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
  while pos <= #s do
    local ws = s:match("^%s+", pos)
    if ws then
      pos = pos + #ws
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
      local val = ""
      if q == '"' or q == "'" then
        val = s:match("^" .. q .. "(.-)" .. q, pos)
        pos = pos + #val + 2
      else
        val = s:match("^([^%s>]+)", pos)
        pos = pos + #val
      end
      attrs[name] = val:gsub("&amp;", "&"):gsub("&lt;", "<"):gsub("&gt;", ">"):gsub("&quot;", '"')
    else
      attrs[name] = true
    end
  end
  return attrs
end

function M.parse(s)
  local pos = 1
  -- Start with a clean ROOT to capture top-level elements
  local root = { tag = "ROOT", children = {}, attrs = {} }
  local stack = { root }

  while pos <= #s do
    local start_tag = s:find("<", pos)
    if not start_tag then
      local text = s:sub(pos)
      if text ~= "" then
        table.insert(stack[#stack].children, { tag = "_TEXT", text = text })
      end
      break
    end

    if start_tag > pos then
      local text = s:sub(pos, start_tag - 1)
      if text:match("%S") then
        table.insert(stack[#stack].children, { tag = "_TEXT", text = text })
      end
    end

    local tag_str, next_pos = read_tag(s, start_tag + 1)
    pos = next_pos

    if tag_str:sub(1, 1) == "/" then
      if #stack > 1 then
        table.remove(stack)
      end
    elseif tag_str:sub(-1) == "/" then
      local name = tag_str:match("^([%w:%-_.]+)")
      local attrs = parse_attrs(tag_str:sub(#name + 1, -2))
      table.insert(stack[#stack].children, { tag = name, attrs = attrs, children = {} })
    elseif tag_str:sub(1, 1) == "?" or tag_str:sub(1, 1) == "!" then
      -- Skip
    else
      local name = tag_str:match("^([%w:%-_.]+)")
      local attrs = parse_attrs(tag_str:sub(#name + 1))
      local node = { tag = name, attrs = attrs, children = {} }
      table.insert(stack[#stack].children, node)
      table.insert(stack, node)
    end
  end
  return root
end

function M.serialize(node)
  if not node then
    return ""
  end

  -- 1. Handle Virtual ROOT
  if node.tag == "ROOT" then
    local res = ""
    for _, child in ipairs(node.children or {}) do
      res = res .. M.serialize(child)
    end
    return res
  end

  -- 2. Handle Text Nodes
  if node.tag == "_TEXT" then
    return tostring(node.text or ""):gsub("&", "&amp;"):gsub("<", "&lt;"):gsub(">", "&gt;")
  end

  -- 3. Build Tag
  local parts = {}
  -- Use table.concat logic but append manually to be safe
  parts[#parts + 1] = "<" .. node.tag

  -- Attributes
  for k, v in pairs(node.attrs or {}) do
    local val = tostring(v):gsub("&", "&amp;"):gsub("<", "&lt;"):gsub(">", "&gt;"):gsub('"', "&quot;")
    parts[#parts + 1] = string.format(' %s="%s"', k, val)
  end

  if #(node.children or {}) == 0 then
    parts[#parts + 1] = "/>"
  else
    parts[#parts + 1] = ">"
    for _, child in ipairs(node.children) do
      parts[#parts + 1] = M.serialize(child)
    end
    parts[#parts + 1] = "</" .. node.tag .. ">"
  end

  return table.concat(parts)
end

function M.serialize_inner(node)
  if not node or not node.children then
    return ""
  end
  local parts = {}
  for _, child in ipairs(node.children) do
    -- Using the #parts+1 syntax is safer than table.insert
    -- because it's impossible to pass a "bad argument #2"
    parts[#parts + 1] = M.serialize(child)
  end
  return table.concat(parts)
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
  -- Handle namespace matching: office:name matches both "office:name" and "name"
  local local_key = key:match(":(.+)$") or key
  for k, v in pairs(node.attrs) do
    local lk = k:match(":(.+)$") or k
    if lk == local_key then
      return v
    end
  end
  return nil
end

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
  local results = M.find_all(node, tag)
  return results[1]
end

function M.find_first_local(node, local_tag)
  if not node or type(node) ~= "table" then
    return nil
  end
  local current_local = node.tag:match(":(.+)$") or node.tag
  if current_local == local_tag then
    return node
  end
  for _, child in ipairs(node.children or {}) do
    local res = M.find_first_local(child, local_tag)
    if res then
      return res
    end
  end
  return nil
end

return M
