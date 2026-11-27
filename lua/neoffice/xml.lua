-- A tiny XML DOM: just enough to edit .docx/.odt parts without losing anything.
--   element: { tag = "w:p", attrs = { { "w:val", "x" }, ... }, children = { ... }, parent = <element> }
--   text:    { text = "decoded text" }
--   other:   { raw = "<?...?>" }  (declarations, comments and CDATA are kept verbatim)
local M = {}

local ENTITIES = { lt = "<", gt = ">", amp = "&", quot = '"', apos = "'" }

local function decode(s)
  if not s:find("&", 1, true) then -- (most text: nothing to decode)
    return s
  end
  return (s:gsub("&(#?)(x?)(%w+);", function(hash, hex, v)
    if hash == "" then
      return ENTITIES[v]
    end
    return vim.fn.nr2char(tonumber(v, hex == "x" and 16 or 10))
  end))
end

function M.escape(s, in_attr)
  if not s:find('[%c&<>"]') then -- (most text: nothing to escape)
    return s
  end
  s = s:gsub("[%z\1-\8\11\12\14-\31]", "") -- (control characters XML can't have at all)
  s = s:gsub("&", "&amp;"):gsub("<", "&lt;"):gsub(">", "&gt;")
  if in_attr then -- (kept as they are: written plainly, a parser would read them as spaces)
    s = s:gsub('"', "&quot;"):gsub("\r", "&#13;"):gsub("\n", "&#10;"):gsub("\t", "&#9;")
  end
  return s
end

-- Whether XML text is a document Word and LibreOffice can read: one root element, and every
-- prefix it uses (w:, text:, svg: ...) declared where it is used. true, or nil and why.
local prefixes = {} -- (a name's prefix, "w" of "w:p", or false: names repeat, so each is read once)
local function prefix_of(name)
  local prefix = prefixes[name]
  if prefix == nil then
    prefix = name:match("^([^:]+):") or false
    prefixes[name] = prefix
  end
  return prefix
end
function M.check(s) -- (s: the text, or the tree it was written from, as M.parse makes them)
  local ok, doc = true, s
  if type(s) == "string" then
    ok, doc = pcall(M.parse, s)
  end
  if not ok or type(doc) ~= "table" then
    return nil, "it can't be read back"
  end
  local roots = 0
  for _, c in ipairs(doc.children) do
    roots = roots + (c.tag and 1 or 0)
  end
  if roots ~= 1 then
    return nil, ("%d root elements"):format(roots)
  end
  local problem
  local function used(name, declared, el)
    local prefix = prefix_of(name)
    if prefix and prefix ~= "xml" and prefix ~= "xmlns" and not declared[prefix] then
      problem = problem or ("the prefix %s (in <%s>) isn't declared"):format(prefix, el.tag)
    end
  end
  local function walk(el, scope)
    local declared, attrs = scope, el.attrs
    for k = 1, #attrs do
      local name = attrs[k][1]
      local prefix = name:byte(1) == 120 and name:match("^xmlns:(.+)$") -- ("x")
      if prefix then
        declared = declared == scope and setmetatable({}, { __index = scope }) or declared
        declared[prefix] = true
      end
    end
    used(el.tag, declared, el)
    for k = 1, #attrs do
      used(attrs[k][1], declared, el)
    end
    local kids = el.children
    for k = 1, #kids do
      local c = kids[k]
      if c.tag and not problem then
        walk(c, declared)
      end
    end
  end
  for _, c in ipairs(doc.children) do
    if c.tag then
      walk(c, {})
    end
  end
  if problem then
    return nil, problem
  end
  return true
end

function M.parse(s)
  local doc = { children = {} }
  local node, pos, n = doc, 1, #s
  local find, match, sub, byte = string.find, string.match, string.sub, string.byte
  while pos <= n do
    local lt = find(s, "<", pos, true) or n + 1
    if lt > pos then
      local kids = node.children
      kids[#kids + 1] = { text = decode(sub(s, pos, lt - 1)), parent = node }
    end
    if lt > n then
      break
    end
    local next_char = byte(s, lt + 1)
    if next_char == 47 then -- closing tag ("/")
      pos = find(s, ">", lt, true) + 1
      node = node.parent
    elseif next_char == 63 or next_char == 33 then -- <?..?>, <!--..-->, <![CDATA[..]]>, <!DOCTYPE..> ("?", "!")
      local close = (sub(s, lt, lt + 3) == "<!--" and "-->") or (sub(s, lt, lt + 8) == "<![CDATA[" and "]]>") or ">"
      local _, stop = find(s, close, lt, true)
      local kids = node.children
      kids[#kids + 1] = { raw = sub(s, lt, stop), parent = node }
      pos = stop + 1
    else -- opening or self-closing tag
      local tag, p = match(s, "^([^%s/>]+)()", lt + 1)
      local attrs = {}
      local el = { tag = tag, attrs = attrs, children = {}, parent = node }
      local kids = node.children
      kids[#kids + 1] = el
      while true do
        local name, _, value, after = match(s, "^%s*([^%s=/>]+)%s*=%s*([\"'])(.-)%2()", p)
        if not name then
          break
        end
        attrs[#attrs + 1] = { name, decode(value) }
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

local function write(n, out, k) -- (k: how many pieces are in out; returns how many after n's)
  if n.text then
    k = k + 1
    out[k] = M.escape(n.text)
    return k
  elseif n.raw then
    k = k + 1
    out[k] = n.raw
    return k
  end
  local tag = n.tag
  local kids = n.children
  if tag then
    k = k + 1
    out[k] = "<" .. tag
    local attrs = n.attrs
    for i = 1, #attrs do
      local kv = attrs[i]
      k = k + 1
      out[k] = " " .. kv[1] .. '="' .. M.escape(kv[2], true) .. '"'
    end
    if #kids == 0 then
      k = k + 1
      out[k] = "/>"
      return k
    end
    k = k + 1
    out[k] = ">"
  end
  local before = k
  for i = 1, #kids do
    k = write(kids[i], out, k)
  end
  if tag then
    if k == before then -- (children that write nothing: an empty element)
      out[before] = "/>"
    else
      k = k + 1
      out[k] = "</" .. tag .. ">"
    end
  end
  return k
end

function M.serialize(n)
  local out = {}
  local k = write(n, out, 0)
  for i = #out, k + 1, -1 do -- (nothing past what was written)
    out[i] = nil
  end
  return table.concat(out)
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
  for i, c in ipairs(n.parent and n.parent.children or {}) do
    if c == n then
      return i
    end
  end
end

function M.remove(n)
  local at = n.parent and M.index(n)
  if at then -- (one no longer among its parent's children: nothing else is taken out)
    table.remove(n.parent.children, at)
  end
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
  local attrs = n.attrs
  if attrs then
    for i = 1, #attrs do
      local kv = attrs[i]
      if kv[1] == name then
        return kv[2]
      end
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
