-- A tiny XML DOM: just enough to edit .docx/.odt parts without losing anything.
--   element: { tag = "w:p", attrs = { { "w:val", "x" }, ... }, children = { ... }, parent = <element> }
--   text:    { text = "decoded text" }
--   other:   { raw = "<?...?>" }  (declarations, comments and CDATA are kept verbatim)
local M = {}

local ENTITIES = { lt = "<", gt = ">", amp = "&", quot = '"', apos = "'" }

local function decode(s)
  return (s:gsub("&(#?)(x?)(%w+);", function(hash, hex, v)
    if hash == "" then
      return ENTITIES[v]
    end
    return vim.fn.nr2char(tonumber(v, hex == "x" and 16 or 10))
  end))
end

function M.escape(s, in_attr)
  s = s:gsub("[%z\1-\8\11\12\14-\31]", "") -- (control characters XML can't have at all)
  s = s:gsub("&", "&amp;"):gsub("<", "&lt;"):gsub(">", "&gt;"):gsub("\r", "&#13;")
  if in_attr then -- (kept as they are: written plainly, a parser would read them as spaces)
    s = s:gsub('"', "&quot;"):gsub("\n", "&#10;"):gsub("\t", "&#9;")
  end
  return s
end

function M.parse(s)
  local doc = { children = {} }
  local node, pos = doc, 1
  while pos <= #s do
    local lt = s:find("<", pos, true) or #s + 1
    if lt > pos then
      M.append(node, { text = decode(s:sub(pos, lt - 1)) })
    end
    if lt > #s then
      break
    end
    local next_char = s:sub(lt + 1, lt + 1)
    if next_char == "/" then -- closing tag
      pos = s:find(">", lt, true) + 1
      node = node.parent
    elseif next_char == "?" or next_char == "!" then -- <?..?>, <!--..-->, <![CDATA[..]]>, <!DOCTYPE..>
      local close = (s:sub(lt, lt + 3) == "<!--" and "-->") or (s:sub(lt, lt + 8) == "<![CDATA[" and "]]>") or ">"
      local _, stop = s:find(close, lt, true)
      M.append(node, { raw = s:sub(lt, stop) })
      pos = stop + 1
    else -- opening or self-closing tag
      local tag, p = s:match("^([^%s/>]+)()", lt + 1)
      local el = M.append(node, M.elem(tag))
      while true do
        local name, _, value, after = s:match("^%s*([^%s=/>]+)%s*=%s*([\"'])(.-)%2()", p)
        if not name then
          break
        end
        table.insert(el.attrs, { name, decode(value) })
        p = after
      end
      local slash, after = s:match("^%s*(/?)>()", p)
      if slash == "" then
        node = el
      end
      pos = after
    end
  end
  return doc
end

function M.serialize(n)
  if n.text then
    return M.escape(n.text)
  elseif n.raw then
    return n.raw
  end
  local inner = {}
  for _, child in ipairs(n.children) do
    inner[#inner + 1] = M.serialize(child)
  end
  inner = table.concat(inner)
  if not n.tag then -- the document node
    return inner
  end
  local attrs = {}
  for _, kv in ipairs(n.attrs) do
    attrs[#attrs + 1] = string.format(' %s="%s"', kv[1], M.escape(kv[2], true))
  end
  local open = "<" .. n.tag .. table.concat(attrs)
  return inner == "" and open .. "/>" or open .. ">" .. inner .. "</" .. n.tag .. ">"
end

-- ── Building ────────────────────────────────────────────────────────────────

-- elem("w:t", { { "xml:space", "preserve" } }, { child, ... }); nil children are skipped.
function M.elem(tag, attrs, children)
  local el = { tag = tag, attrs = attrs or {}, children = {} }
  for i = 1, table.maxn(children or {}) do
    if children[i] then
      M.append(el, children[i])
    end
  end
  return el
end

function M.append(parent, child, index)
  child.parent = parent
  table.insert(parent.children, index or #parent.children + 1, child)
  return child
end

function M.index(n)
  for i, c in ipairs(n.parent.children) do
    if c == n then
      return i
    end
  end
end

function M.remove(n)
  table.remove(n.parent.children, M.index(n))
  n.parent = nil
end

function M.copy(n)
  if not n.tag then
    return { text = n.text, raw = n.raw }
  end
  local el = M.elem(n.tag, vim.deepcopy(n.attrs))
  for _, c in ipairs(n.children) do
    M.append(el, M.copy(c))
  end
  return el
end

-- ── Querying ────────────────────────────────────────────────────────────────

function M.attr(n, name)
  for _, kv in ipairs(n.attrs or {}) do
    if kv[1] == name then
      return kv[2]
    end
  end
  return nil
end

function M.set_attr(n, name, value) -- value nil removes the attribute
  for i, kv in ipairs(n.attrs) do
    if kv[1] == name then
      if value then
        kv[2] = value
      else
        table.remove(n.attrs, i)
      end
      return
    end
  end
  if value then
    table.insert(n.attrs, { name, value })
  end
end

function M.child(n, tag)
  for _, c in ipairs(n.children or {}) do
    if c.tag == tag then
      return c
    end
  end
end

function M.find(n, tag) -- first descendant (or self) with this tag
  if n.tag == tag then
    return n
  end
  for _, c in ipairs(n.children or {}) do
    local found = M.find(c, tag)
    if found then
      return found
    end
  end
end

-- All descendants whose tag is in `tags` (without looking inside them), skipping `skip` subtrees.
function M.collect(n, tags, skip, out)
  out = out or {}
  for _, c in ipairs(n.children or {}) do
    if tags[c.tag] then
      out[#out + 1] = c
    elseif c.tag and not (skip and skip[c.tag]) then
      M.collect(c, tags, skip, out)
    end
  end
  return out
end

function M.text(n) -- all text inside a node
  if n.text then
    return n.text
  end
  local parts = {}
  for _, c in ipairs(n.children or {}) do
    parts[#parts + 1] = M.text(c)
  end
  return table.concat(parts)
end

return M
