-- :Neoffice merge {file.Rnw/.tex ...}: bring a revised document back into the LaTeX (or Rnoweb)
-- it was made from. Each paragraph of the document is paired with its paragraph of the source (by
-- its words); its text as it was sent (all tracked changes rejected) tells what each citation,
-- \Sexpr value, reference... of the source looked like; and its differences with the text as
-- revised (all changes accepted) are made in the source, in LaTeX: the words that changed, as
-- little as possible around them, never cutting a command or an \emph{} in two. Citations typed
-- in the document are looked up in the .bib. The result opens in a diff view against the source;
-- the comments, and what needs a human, go to the quickfix list.
local api = vim.api
local markup = require("neoffice.markup")
local M = {}

local function neo()
  return require("neoffice")
end

-- ── Text: tokens ────────────────────────────────────────────────────────────

-- Characters that separate words, punctuation, and how quotes compare.
local SPACE = { [" "] = " ", ["\t"] = " ", ["\n"] = " ", ["\194\160"] = "~", ["\226\128\175"] = "~", ["\226\129\159"] = " ", ["\227\128\128"] = " " }
for b = 0x80, 0x8A do -- (U+2000..U+200A: thin, hair, en spaces...: a thin space)
  SPACE["\226\128" .. string.char(b)] = "\\,"
end
local PUNCT = {}
for _, c in ipairs({ "–", "—", "‘", "’", "“", "”", "°", "−", "×", "…", "±", "«", "»", "¶" }) do
  PUNCT[c] = true
end
local SAME = { ["‘"] = "'", ["’"] = "'", ["“"] = '"', ["”"] = '"', ["−"] = "-" }

local function char_at(s, p)
  return s:match("^.[\128-\191]*", p) or ""
end
local function is_word(c)
  return c:match("^[%w]") ~= nil or (c:byte() or 0) >= 128 and not PUNCT[c] and not SPACE[c]
end
local function flags_key(f)
  return (f.b and "b" or "") .. (f.i and "i" or "") .. (f.sup and "^" or "") .. (f.sub and "_" or "")
end

-- The tokens of a list of characters { ch, b, i, sup, sub, ... } (a paragraph of the document): words
-- and punctuation, each { key, text, i1, i2 (its characters) }; images, footnotes... one token each.
local function char_tokens(chars)
  local out, k = {}, 1
  while k <= #chars do
    local c = chars[k]
    if c.ch == "" then -- (an image, equation, footnote, field)
      local what = c.footnote and "footnote:" .. c.footnote.text or c.field and "field:" .. c.field.text
        or c.math and "math:" .. c.math.tex or c.image and "image:" .. c.image.path or "object"
      out[#out + 1] = { key = "\1" .. what, text = "", i1 = k, i2 = k, object = c }
      k = k + 1
    elseif SPACE[c.ch] then
      k = k + 1
    else
      local j = k
      if is_word(c.ch) then
        while chars[j + 1] and chars[j + 1].ch ~= "" and is_word(chars[j + 1].ch) and flags_key(chars[j + 1]) == flags_key(c) do
          j = j + 1
        end
      end
      local text = table.concat(vim.tbl_map(function(x) return x.ch end, vim.list_slice(chars, k, j)))
      out[#out + 1] = { key = (SAME[text] or text) .. "|" .. flags_key(c), text = text, i1 = k, i2 = j }
      k = j + 1
    end
  end
  return out
end

-- Where the group that starts at `p` ("{...}" or "[...]") ends (the position after it).
local function group_end(src, p)
  local open, close = src:sub(p, p), src:sub(p, p) == "{" and "}" or "]"
  local depth = 0
  while p <= #src do
    local c = src:sub(p, p)
    if c == "\\" then
      p = p + 1
    elseif c == open then
      depth = depth + 1
    elseif c == close then
      depth = depth - 1
      if depth == 0 then
        return p + 1
      end
    end
    p = p + 1
  end
  return p
end

-- What LaTeX commands do to the text (see tex_tokens).
local FORMAT = { emph = "i", textit = "i", textbf = "b", textsuperscript = "sup", textsubscript = "sub",
  texttt = "", textrm = "", textsf = "", textup = "", textnormal = "", mbox = "", text = "" }
local KIND = { Sexpr = "sexpr", cite = "cite", parencite = "cite", textcite = "cite", autocite = "cite", citep = "cite",
  citet = "cite", Cite = "cite", Parencite = "cite", Textcite = "cite", footcite = "cite",
  ref = "ref", cref = "ref", Cref = "ref", autoref = "ref", eqref = "ref", pageref = "ref", nameref = "ref",
  si = "unit", SI = "unit", num = "unit", qty = "unit", unit = "unit", ang = "unit", footnote = "footnote" }
local INVISIBLE = { label = true, index = true, begin = true, ["end"] = true, noindent = true, centering = true,
  linenumbers = true, nolinenumbers = true, singlespacing = true, doublespacing = true, maketitle = true, par = true }
local ESCAPED = { ["%"] = true, ["&"] = true, ["_"] = true, ["#"] = true, ["$"] = true, ["{"] = true, ["}"] = true }
local LIGATURE = { ["---"] = "—", ["--"] = "–", ["``"] = "“", ["''"] = "”", ["`"] = "‘" }

-- The tokens of LaTeX source src[s, e): words and punctuation as the document shows them, each
-- { key, text, s, e (its source), wraps (the formatting commands it is in: { s, e }) }; a
-- citation, \Sexpr, reference, unit, math... one token { opaque = its kind, s, e }.
local function tex_tokens(src, s, e, flags, wraps, out)
  out, flags, wraps = out or {}, flags or {}, wraps or {}
  local p = s
  local function add(text, a, b, opaque)
    local key = opaque and ("\0" .. #out) or ((SAME[text] or text) .. "|" .. flags_key(flags))
    out[#out + 1] = { key = key, text = text, s = a, e = b, opaque = opaque, wraps = wraps, flags = flags }
  end
  while p < e do
    local c = char_at(src, p)
    local name = c == "\\" and src:match("^\\([a-zA-Z]+)%*?", p)
    if c == "%" then -- a comment: to the end of the line
      p = (src:find("\n", p, true) or e) + 1
    elseif c == "\\" and not name then -- \% \& ... (escaped), \, \  (spaces)
      local x = src:sub(p + 1, p + 1)
      if ESCAPED[x] then
        add(x, p, p + 2)
      end
      p = p + 2
    elseif name then
      local after = p + 1 + #name + (src:sub(p + 1 + #name, p + 1 + #name) == "*" and 1 or 0)
      local args = after
      while src:sub(args, args) == "[" or src:sub(args, args) == "{" do
        args = group_end(src, args)
        if KIND[name] ~= "cite" and KIND[name] ~= "unit" and name ~= "SI" and name ~= "href" and src:sub(args, args) == "{" and FORMAT[name] then
          break
        end
      end
      if FORMAT[name] and src:sub(after, after) == "{" then -- formatting: its text, with the formatting
        local close = group_end(src, after)
        local f = vim.tbl_extend("force", {}, flags)
        if FORMAT[name] ~= "" then
          f[FORMAT[name]] = true
        end
        tex_tokens(src, after + 1, close - 1, f, vim.list_extend(vim.list_slice(wraps), { { s = p, e = close, flags = f } }), out)
        p = close
      elseif KIND[name] or (args > after and not INVISIBLE[name]) then -- one token: what it shows is known later
        add("", p, args, KIND[name] or "command")
        p = args
      else -- (\noindent, \label{...}, \begin{abstract}...: nothing shown)
        p = args
      end
    elseif c == "$" then -- math
      local stop = src:find(src:sub(p, p + 1) == "$$" and "$$" or "$", p + (src:sub(p, p + 1) == "$$" and 2 or 1), true)
      local close = stop and stop + (src:sub(p, p + 1) == "$$" and 2 or 1) or e
      add("", p, close, "math")
      p = close
    elseif c == "{" or c == "}" or SPACE[c] or c == "~" then
      p = p + #c
    elseif LIGATURE[src:sub(p, p + 2)] or LIGATURE[src:sub(p, p + 1)] or LIGATURE[c] then
      local lig = LIGATURE[src:sub(p, p + 2)] and src:sub(p, p + 2) or LIGATURE[src:sub(p, p + 1)] and src:sub(p, p + 1) or c
      add(LIGATURE[lig], p, p + #lig)
      p = p + #lig
    elseif is_word(c) then
      local q = p
      while q < e and is_word(char_at(src, q)) do
        q = q + #char_at(src, q)
      end
      add(src:sub(p, q - 1), p, q)
      p = q
    else
      add(c, p, p + #c)
      p = p + #c
    end
  end
  return out
end

-- ── The document: each paragraph as sent and as revised ─────────────────────

local function covers(m, row, col)
  local d = m[4]
  return (m[2] < row or m[2] == row and m[3] <= col) and (d.end_row > row or d.end_row == row and d.end_col > col)
end

-- The characters of a paragraph as sent (sent: insertions out, deletions in, the formatting from
-- before formatting changes) and as revised (deletions out).
local function versions(buf, block)
  local doc = neo().docs[buf]
  local _, chars = markup.parse(block.text)
  local lines = api.nvim_buf_get_lines(buf, block.first - 1, block.last, true)
  local row, start = 1, -(block.col or 0)
  for _, ch in ipairs(chars) do
    while row < #lines and ch.col > start + #lines[row] do
      start, row = start + #lines[row] + 1, row + 1
    end
    ch.row, ch.col = block.first - 2 + row, ch.col - start
  end
  local marks = api.nvim_buf_get_extmarks(buf, neo().ns.change, { block.first - 1, 0 }, { block.last - 1, -1 }, { details = true, overlap = true })
  local sent, revised = {}, {}
  for _, ch in ipairs(chars) do
    local inserted, deleted, format
    for _, m in ipairs(marks) do
      local change = doc.change[m[1]]
      if change and covers(m, ch.row, ch.col) then
        inserted = inserted or change.type == "insert"
        deleted = deleted or change.type == "delete"
        format = change.type == "format" and change or format
      end
    end
    if not inserted then
      local was = vim.tbl_extend("force", {}, ch)
      if format and format.old_flags then
        for _, f in ipairs({ "b", "i", "u", "sup", "sub" }) do
          was[f] = format.old_flags[f]
        end
      end
      was.ch = (deleted and ch.ch == "¶" and not ch.esc) and " " or was.ch -- (a deleted paragraph break)
      sent[#sent + 1] = was
    end
    if not deleted then
      revised[#revised + 1] = ch
    end
  end
  return sent, revised
end

-- ── The source: its paragraphs ──────────────────────────────────────────────

-- The paragraphs of a LaTeX/Rnoweb file (not code chunks): blocks between blank lines, and the
-- arguments of \caption{} and \section{}. Each { path, s, e (in text), line }.
local function regions(path, text)
  local out, line_of = {}, {}
  local function add(s, e, least) -- (with at least `least` words)
    local words = vim.tbl_filter(function(t) return not t.opaque and t.text:match("%a") end, tex_tokens(text, s, e))
    if #words >= least then
      out[#out + 1] = { path = path, s = s, e = e, line = select(2, text:sub(1, s):gsub("\n", "")) + 1 }
    end
  end
  local function block(s, e)
    local b = text:sub(s, e - 1)
    local found = false
    for _, cmd in ipairs({ "caption", "section", "subsection", "subsubsection", "paragraph" }) do
      local from = 1
      while true do
        local a, b2 = b:find("\\" .. cmd .. "%*?%b{}", from)
        if not a then
          break
        end
        local open = b:find("{", a, true)
        add(s + open, s + group_end(b, open) - 2, 1) -- (a heading of one word too: see M.compute)
        from, found = b2 + 1, true
      end
    end
    if not found and not b:match("^%s*\\begin{figure") and not b:match("^%s*\\begin{table") then
      add(s, e, 2)
    end
  end
  local start, chunk, pos = nil, false, 1
  for line in (text .. "\n"):gmatch("([^\n]*)\n") do
    local s = pos
    pos = pos + #line + 1
    if chunk then
      chunk = not line:match("^@%s*$")
    elseif line:match("^%s*<<.*>>=%s*$") then
      if start then
        block(start, s)
      end
      start, chunk = nil, true
    elseif line:match("^%s*$") then
      if start then
        block(start, s - 1)
      end
      start = nil
    else
      start = start or s
    end
  end
  if start then
    block(start, #text + 1)
  end
  return out
end

-- ── Citations ───────────────────────────────────────────────────────────────

local ACCENTS = {}
for plain, list in pairs({ a = "áàâäãå", e = "éèêë", i = "íìîïı", o = "óòôöõø", u = "úùûü", n = "ñ", c = "ç", s = "š", z = "ž", l = "ł" }) do
  for c in list:gmatch("[^\128-\191][\128-\191]*") do
    ACCENTS[c], ACCENTS[vim.fn.toupper(c)] = plain, plain
  end
end
local LETTERS = { i = "i", j = "j", o = "o", O = "O", l = "l", L = "L", ss = "ss", ae = "ae", AE = "AE", oe = "oe", OE = "OE", aa = "a", AA = "A" }
local function fold(s) -- (lower case, without accents, braces or LaTeX accent commands)
  s = s:gsub("\204[\128-\191]", ""):gsub("\205[\128-\175]", "") -- (combining accents, U+0300-036F)
  s = s:gsub("\\(%a+)%f[%A]", LETTERS):gsub("\\[`'^\"~=.uvHc]%s*{?(%a)}?", "%1"):gsub("[{}]", ""):gsub("[^\128-\191][\128-\191]+", function(c) return ACCENTS[c] or c end)
  return s:lower()
end

-- The entries of a .bib: key -> { surname (folded, of the first author), year }.
local function read_bib(paths)
  local bib = {}
  for _, path in ipairs(paths) do
    local f = io.open(path)
    local text = f and f:read("*a") or ""
    if f then
      f:close()
    end
    for kind, entry in text:gmatch("@(%a+)%s*(%b{})") do -- (each entry, by its braces: on one line or many)
      local key, body = entry:match("^{%s*([^,%s]+)%s*,(.*)}$")
      if key and not ({ comment = true, string = true, preamble = true })[kind:lower()] then
        local function field(name)
          local v = body:match("[%s,]" .. name .. "%s*=%s*(%b{})") or body:match("[%s,]" .. name .. '%s*=%s*"([^"]*)"')
          return v and v:gsub("^{(.*)}$", "%1")
        end
        local author = (field("author") or field("editor") or ""):gsub("%s+", " ")
        local first = vim.split(author, " and ", { plain = true })[1] or ""
        local surname = first:find(",") and first:match("^%s*(.-)%s*,") or first:match("(%S+)%s*$") or ""
        local year = (field("year") or field("date") or ""):match("%d%d%d%d")
        local _, ands = author:gsub(" and ", "")
        bib[key] = { surname = fold(surname), year = year, authors = ands + 1 }
      end
    end
  end
  return bib
end

-- Citations written as text: "(Freitas et al., 2023; Scaramuzzi et al., 2023)". Each { a, b
-- (positions), items = { { name, year } } }.
local function written_citations(text)
  local out = {}
  for a, inside, b in text:gmatch("()%(([^()]-%d%d%d%d%a?[^()]-)%)()") do
    local items = {}
    for part in (inside .. ";"):gmatch("([^;]+);") do
      local name, years = part:match("^%s*(.-),?%s+(%d%d%d%d%a?[%d%a,%s]*)%s*$")
      for year in (years or ""):gmatch("%d%d%d%d%a?") do
        items[#items + 1] = { name = name, year = year }
      end
    end
    if #items > 0 then
      out[#out + 1] = { a = a, b = b, items = items }
    end
  end
  return out
end

-- The key of a citation item: among `prefer` (the keys the text had) first, then in the .bib.
local function resolve(item, bib, prefer)
  local name = fold(item.name or ""):gsub(" et al%.?", ""):gsub("%s*&.*", ""):gsub("%s+and%s.*", "")
  local year, letter = item.year:match("^(%d%d%d%d)(%a?)$")
  for _, keys in ipairs({ prefer or {}, vim.tbl_keys(bib) }) do
    local found = {}
    for _, key in ipairs(keys) do
      local e = bib[key]
      if e and e.year == year and e.surname ~= "" and (name:sub(1, #e.surname) == e.surname or name:find(e.surname, 1, true)) then
        found[#found + 1] = key
      end
    end
    if #found > 1 then -- (one, two or more authors: "Name", "Name & Other", "Name et al.")
      local raw = fold(item.name or "")
      local n = raw:find("et al", 1, true) and 3 or (raw:find("&", 1, true) or raw:find(" and ", 1, true)) and 2 or 1
      local by_count = vim.tbl_filter(function(k) return n == 3 and bib[k].authors >= 3 or bib[k].authors == n end, found)
      found = #by_count > 0 and by_count or found
    end
    if #found > 1 and letter ~= "" then
      found = vim.tbl_filter(function(k) return k:sub(-#(year .. letter)) == year .. letter end, found)
    end
    if #found == 1 then
      return found[1]
    elseif #found > 1 then
      return nil, found
    end
  end
end

-- ── LaTeX from the document's text ──────────────────────────────────────────

local function escape(c)
  if c == "\\" then
    return "\\textbackslash{}"
  elseif ESCAPED[c] then
    return "\\" .. c
  elseif c == "~" then
    return "\\textasciitilde{}"
  elseif c == "^" then
    return "\\textasciicircum{}"
  end
  return SPACE[c] or c
end

local WRAP = { i = "\\emph{", b = "\\textbf{", sup = "\\textsuperscript{", sub = "\\textsubscript{" }

-- LaTeX for characters of the document: formatting (but what the source has around them already,
-- `base`), escapes; citations written as text become \parencite{...} when the .bib knows them.
local function latex(chars, base, bib, prefer, issues, where)
  local plain = table.concat(vim.tbl_map(function(c) return c.ch ~= "" and c.ch or "\1" end, chars))
  local cites, at = {}, {} -- (by byte position in plain -> character index)
  local pos = 1
  for k, c in ipairs(chars) do
    at[pos] = k
    pos = pos + math.max(#c.ch, 1)
  end
  at[pos] = #chars + 1
  for _, cite in ipairs(written_citations(plain)) do
    local keys, missing, first = {}, {}, nil
    for _, item in ipairs(cite.items) do
      local key, options = resolve(item, bib, prefer)
      keys[#keys + 1] = key
      if not key then
        missing[#missing + 1] = item.name .. " " .. item.year .. (options and (" (" .. table.concat(options, ", ") .. "?)") or "")
        first = first or item
      end
    end
    if #missing == 0 and at[cite.a] and at[cite.b] then
      cites[at[cite.a]] = { stop = at[cite.b], text = "\\parencite{" .. table.concat(keys, ",") .. "}" }
    elseif #missing > 0 then
      issues[#issues + 1] = vim.tbl_extend("force", where, {
        text = "citation not found in the .bib: " .. table.concat(missing, "; "),
        cite = { name = first.name, year = first.year }, -- (to find it in the proposal)
      })
    end
  end
  local out, open, k = {}, {}, 1
  local function set(f) -- (open and close \emph{} ... to have formatting f)
    for n = #open, 1, -1 do
      if not f[open[n]] then
        for _ = n, #open do
          out[#out + 1] = "}"
          table.remove(open)
        end
        break
      end
    end
    for _, flag in ipairs({ "b", "i", "sup", "sub" }) do
      if f[flag] and not base[flag] and not vim.tbl_contains(open, flag) then
        out[#out + 1] = WRAP[flag]
        open[#open + 1] = flag
      end
    end
  end
  while k <= #chars do
    local c = chars[k]
    if cites[k] then
      set({})
      out[#out + 1] = cites[k].text
      k = cites[k].stop
    else
      if not SPACE[c.ch] then
        set(c)
      else -- (a space: out of the formatting the next word isn't in, not into the one it is)
        local n = k + 1
        while chars[n] and SPACE[chars[n].ch] do
          n = n + 1
        end
        local f = {}
        for _, flag in ipairs(open) do
          f[flag] = chars[n] and chars[n][flag] and not cites[n] or nil
        end
        set(f)
      end
      if c.footnote then
        out[#out + 1] = "\\footnote{" .. c.footnote.text:gsub("[\\%%&_#$]", "\\%0") .. "}"
      elseif c.math then
        out[#out + 1] = (c.math.display and "$$" or "$") .. c.math.tex .. (c.math.display and "$$" or "$")
      elseif c.ch ~= "" then
        out[#out + 1] = escape(c.ch)
      end
      k = k + 1
    end
  end
  set({})
  return table.concat(out)
end

-- The bracketed citations of tokens of the document: { from, to } (token indices), each a "(" to
-- its ")" with a year inside.
local function cited(tokens)
  local out, opened = {}, {}
  for k, t in ipairs(tokens) do
    if t.text == "(" then
      opened[#opened + 1] = k
    elseif t.text == ")" and #opened > 0 then
      local from = table.remove(opened)
      for j = from, k do
        if tokens[j].text:match("^%d%d%d%d%a?$") then
          out[#out + 1] = { from, k }
          break
        end
      end
    end
  end
  return out
end

-- ── A paragraph ─────────────────────────────────────────────────────────────

local function keys(list)
  return vim.tbl_map(function(t) return t.key end, list)
end

-- The source of a paragraph (src[s, e)) with the edits of the document (sent -> revised), and
-- what needs a human.
local function merge_paragraph(src, s, e, sent, revised, bib, where)
  local issues = {}
  local R, O, N = tex_tokens(src, s, e), char_tokens(sent), char_tokens(revised)
  -- 1. what the source's citations, values... showed as: where it and the sent text differ
  -- where the source already reads as the revised text (its edits taken already: a merge done
  -- again): at_n[i], the revised token that source token i is
  local index, at_n = {}, {}
  for i, t in ipairs(R) do
    index[t] = i
  end
  do
    local a, b = 1, 1
    local hunks = neo().diff(keys(R), keys(N))
    hunks[#hunks + 1] = { #R + 1, 0, #N + 1, 0 }
    for _, h in ipairs(hunks) do
      while a < h[1] do
        at_n[a], a, b = b, a + 1, b + 1
      end
      a, b = h[1] + h[2], h[3] + h[4]
    end
  end
  local function as_revised(tokens) -- (the source has these as the revised text does, there)
    for k, t in ipairs(tokens) do
      local i = index[t]
      if not at_n[i] or k > 1 and at_n[i] ~= at_n[index[tokens[k - 1]]] + 1 then
        return false
      end
    end
    local first, last = index[tokens[1]], index[tokens[#tokens]]
    local p, q = at_n[first - 1], at_n[last + 1]
    return (not p or p == at_n[first] - 1) and (not q or q == at_n[last] + 1)
  end
  local function gone(i) -- (the revised text, like the source, has nothing before source token i)
    local p, q = i > 1 and at_n[i - 1] or 0, i <= #R and at_n[i] or #N + 1
    return p and q and q == p + 1
  end

  local seq = {} -- the source as the document showed it: { key, owner = token or group }
  local ours = {} -- (source tokens the sent text doesn't have)
  local function align(R, O) -- (source tokens R, which the document showed as its tokens O)
    local a = 1
    local hunks = neo().diff(keys(R), keys(O))
    hunks[#hunks + 1] = { #R + 1, 0, #O + 1, 0 }
    for _, h in ipairs(hunks) do
      while a < h[1] do
        seq[#seq + 1], a = { key = R[a].key, owner = R[a] }, a + 1
      end
      local part, shown = vim.list_slice(R, h[1], h[1] + h[2] - 1), vim.list_slice(O, h[3], h[3] + h[4] - 1)
      local opaque = vim.iter(part):find(function(t) return t.opaque end)
      if opaque and #shown > 0 then -- (citations...: the text they showed, as groups)
        local cites = cited(shown)
        local mine = vim.tbl_filter(function(t) return t.opaque == "cite" end, part)
        local pieces = {} -- { R tokens, sent tokens }
        if #mine > 0 and #mine == #cites then -- (each \parencite: its "(...)"; between them, the rest)
          local r, o = 1, 1
          for n, c in ipairs(cites) do
            local at = vim.fn.index(part, mine[n]) + 1
            pieces[#pieces + 1] = { vim.list_slice(part, r, at - 1), vim.list_slice(shown, o, c[1] - 1) }
            pieces[#pieces + 1] = { { mine[n] }, vim.list_slice(shown, c[1], c[2]) }
            r, o = at + 1, c[2] + 1
          end
          pieces[#pieces + 1] = { vim.list_slice(part, r), vim.list_slice(shown, o) }
          for n = #pieces, 1, -1 do -- (text with no source of its own: a narrative citation's names)
            local piece = pieces[n]
            if #piece[1] == 0 and #piece[2] > 0 then
              local into = pieces[n + 1] and #pieces[n + 1][1] > 0 and pieces[n + 1] or pieces[n - 1]
              if into == pieces[n + 1] then
                into[2] = vim.list_extend(vim.list_slice(piece[2]), into[2])
              elseif into then
                vim.list_extend(into[2], piece[2])
              end
              table.remove(pieces, n)
            elseif #piece[1] == 0 then
              table.remove(pieces, n)
            end
          end
        else
          pieces = { { part, shown } }
        end
        for _, piece in ipairs(pieces) do
          local tokens, text = piece[1], piece[2]
          local kind = vim.iter(tokens):find(function(t) return t.opaque end)
          if kind and #text > 0 then
            local group = { s = tokens[1].s, e = tokens[#tokens].e, kind = kind.opaque, tokens = tokens, wraps = tokens[1].wraps, flags = tokens[1].flags }
            for _, t in ipairs(text) do
              seq[#seq + 1] = { key = t.key, owner = group }
            end
          elseif kind then -- (shown as nothing)
            vim.list_extend(ours, tokens)
          else -- (words: as any others)
            align(tokens, text)
          end
        end
      elseif #part == 0 and (R[h[1]] and gone(index[R[h[1]]]) or not R[h[1]] and R[h[1] - 1] and gone(index[R[h[1] - 1]] + 1)) then
        -- (deleted in the document, and in the source already)
      elseif #part > 0 and as_revised(part) then -- (as the document has it now: its edit taken already)
        for _, t in ipairs(part) do
          seq[#seq + 1] = { key = t.key, owner = t }
        end
      elseif #part == 0 then -- (text the source doesn't have: written by the export, "Figure 1:")
        local at = R[h[1] - 1] and R[h[1] - 1].e or R[h[1]] and R[h[1]].s or s
        local generated = { s = at, e = at, kind = "generated", wraps = {}, tokens = {} }
        for _, t in ipairs(shown) do
          seq[#seq + 1] = { key = t.key, owner = generated }
        end
      elseif #shown == 0 then -- (words the source has and the sent text doesn't: written since)
        vim.list_extend(ours, part)
      else -- (words that differ: written since; the source's, unless the document changes them too)
        local group = { s = part[1].s, e = part[#part].e, kind = "text", tokens = part, wraps = part[1].wraps, flags = part[1].flags }
        for _, t in ipairs(shown) do
          seq[#seq + 1] = { key = t.key, owner = group }
        end
      end
      a = h[1] + h[2]
    end
  end
  align(R, O)

  -- 2. what changed from that to the revised text: cores (grown to whole citations)
  local cores, hunks2 = {}, {}
  for _, h in ipairs(neo().diff(keys(seq), keys(N))) do
    cores[#cores + 1] = { r1 = h[1], r2 = h[1] + h[2] - 1, n1 = h[3], n2 = h[3] + h[4] - 1 }
    hunks2[#hunks2 + 1] = vim.deepcopy(cores[#cores]) -- (what really changed: cores grow below)
  end
  for _, c in ipairs(cores) do
    local lo, hi = c.r1, c.r2
    local p, q = seq[lo - 1], seq[lo]
    if hi < lo and p and q and p.owner == q.owner and p.owner.kind then -- (typed inside a citation...)
      lo, hi = lo - 1, lo
    end
    for x = lo, hi do
      local g = seq[x].owner
      if g.kind == "cite" then
        while lo > 1 and seq[lo - 1].owner == g do
          lo = lo - 1
        end
        while hi < #seq and seq[hi + 1].owner == g do
          hi = hi + 1
        end
      end
    end
    c.n1, c.n2, c.r1, c.r2 = c.n1 - (c.r1 - lo), c.n2 + (hi - c.r2), lo, hi
  end
  for _, c in ipairs(cores) do -- (and whole citations of the revised text)
    for _, span in ipairs(cited(N)) do
      if c.n1 <= span[2] and c.n2 >= span[1] and c.n2 >= c.n1 then
        local left, right = math.max(c.n1 - span[1], 0), math.max(span[2] - c.n2, 0)
        c.n1, c.n2, c.r1, c.r2 = c.n1 - left, c.n2 + right, math.max(c.r1 - left, 1), math.min(c.r2 + right, #seq)
      end
    end
  end
  table.sort(cores, function(x, y) return x.r1 < y.r1 end)
  for k = #cores, 2, -1 do -- (cores grown into each other: one)
    local p, c = cores[k - 1], cores[k]
    if c.r1 <= p.r2 + 1 then
      p.r1, p.r2, p.n1, p.n2 = math.min(p.r1, c.r1), math.max(p.r2, c.r2), math.min(p.n1, c.n1), math.max(p.n2, c.n2)
      table.remove(cores, k)
    end
  end
  local changed = {} -- (seq entries that really changed; citations they touch: all of them)
  for _, h in ipairs(hunks2) do
    for x = h.r1, h.r2 do
      changed[x] = true
    end
    local p, q = seq[h.r1 - 1], seq[h.r1]
    if h.r2 < h.r1 and p and q and p.owner == q.owner and p.owner.kind then -- (typed inside one)
      changed[h.r1 - 1], changed[h.r1] = true, true
    end
  end
  for x, entry in ipairs(seq) do
    if changed[x] and entry.owner.kind == "cite" then
      for y, other in ipairs(seq) do
        if other.owner == entry.owner then
          changed[y] = true
        end
      end
    end
  end
  local function shift(x) -- (the revised token that the unchanged seq[x] is)
    local d = 0
    for _, h in ipairs(hunks2) do
      if h.r2 < x and h.r1 <= x then
        d = d + (h.n2 - h.n1) - (h.r2 - h.r1)
      end
    end
    return x + d
  end
  do -- (a citation of the revised text with something new in it: all of it written again)
    local copied = {}
    for x = 1, #seq do
      if not changed[x] then
        copied[shift(x)] = x
      end
    end
    for _, span in ipairs(cited(N)) do
      local new = false
      for y = span[1], span[2] do
        new = new or not copied[y]
      end
      if new then
        for y = span[1], span[2] do
          if copied[y] then
            changed[copied[y]] = true
          end
        end
      end
    end
  end

  -- 3. each core's source span: whole formatting commands where it would cut one; cores whose
  -- spans meet make one edit
  local function inside(w, pos) -- (pos within the formatting w's text)
    return w.s < pos and pos < w.e
  end
  local edits = {}
  for _, c in ipairs(cores) do
    local A, B, wa, wb
    if c.r2 >= c.r1 then
      A, B, wa, wb = seq[c.r1].owner.s, seq[c.r2].owner.e, seq[c.r1].owner.wraps or {}, seq[c.r2].owner.wraps or {}
    else
      local before, after = seq[c.r1 - 1], seq[c.r1]
      A = before and before.owner.e or after and after.owner.s or s
      B, wa = A, vim.tbl_filter(function(w) return inside(w, A) end, before and before.owner.wraps or after and after.owner.wraps or {})
      wb = wa
      for _, w in ipairs(wa) do -- (typed inside \emph{}... but not in its formatting: it is rewritten)
        local typed = vim.list_slice(revised, c.n1 <= c.n2 and N[c.n1].i1 or 1, c.n1 <= c.n2 and N[c.n2].i2 or 0)
        for flag in pairs(w.flags) do
          if vim.iter(typed):any(function(ch) return not SPACE[ch.ch] and not ch[flag] end) then
            A, B = math.min(A, w.s), math.max(B, w.e)
          end
        end
      end
    end
    for _, list in ipairs({ wa, wb }) do
      for _, w in ipairs(list) do
        if not (vim.tbl_contains(wa, w) and vim.tbl_contains(wb, w)) then
          A, B = math.min(A, w.s), math.max(B, w.e)
        end
      end
    end
    local last = edits[#edits]
    if last and A < last.B then
      last.B, last.cores[#last.cores + 1] = math.max(last.B, B), c
      last.A = math.min(last.A, A)
    else
      edits[#edits + 1] = { A = A, B = B, cores = { c } }
    end
  end

  -- 4. the new source of each edit (from the end, so positions stay right): unchanged words as
  -- they are in the source, changed ones in LaTeX from the revised text
  local out = src:sub(s, e - 1)
  for k = #edits, 1, -1 do
    local edit = edits[k]
    local x1, x2 = nil, nil -- (the seq entries in the span)
    for x, entry in ipairs(seq) do
      if entry.owner.s >= edit.A and entry.owner.e <= edit.B then
        x1, x2 = x1 or x, x
      end
    end
    for _, c in ipairs(edit.cores) do
      x1 = math.min(x1 or c.r1, c.r1)
      x2 = math.max(x2 or c.r2, c.r2)
    end
    local common = {} -- (the formatting around all of the edit: not written again)
    local first = seq[x1] or seq[x1 - 1]
    for _, w in ipairs(first and first.owner.wraps or {}) do
      if w.s < edit.A and edit.B < w.e then
        common[w] = true
      end
    end
    local base = {}
    for w in pairs(common) do
      base = vim.tbl_extend("force", base, w.flags)
    end
    local parts, open, ok, last_y, done_group = {}, {}, true, nil, nil
    local first_converted, last_converted -- (whether the edit starts, ends with revised text)
    local function close_all()
      for _ = 1, #open do
        parts[#parts + 1] = "}"
      end
      open = {}
    end
    local function spacing(y) -- (the revised text's spacing before its token y)
      if last_y and N[last_y] and N[y] and #parts > 0 then
        local between = vim.list_slice(revised, N[last_y].i2 + 1, N[y].i1 - 1)
        parts[#parts + 1] = table.concat(vim.tbl_map(function(ch) return escape(ch.ch) end, between))
      end
    end
    local function converted(y1, y2, prefer) -- (revised tokens y1..y2, in LaTeX)
      if y2 < y1 then
        return
      end
      close_all()
      spacing(y1)
      prefer = vim.fn.uniq(vim.fn.sort(prefer or {}))
      local text = latex(vim.list_slice(revised, N[y1].i1, N[y2].i2), base, bib, #prefer > 0 and prefer or nil, issues, where)
      if #prefer > 0 and not text:find("\\parencite") then
        issues[#issues + 1] = vim.tbl_extend("force", where, { text = "a citation was edited, check it: " .. text })
      end
      if first_converted == nil then
        first_converted = true
      end
      parts[#parts + 1], last_y, done_group, last_converted = text, y2, nil, true
    end
    local function source(x, y) -- (an unchanged token: as it is in the source)
      local o = seq[x].owner
      if o.kind == "generated" then -- (not in the source)
        return
      end
      first_converted, last_converted = first_converted or false, false
      if not (done_group and o == done_group) then
        local want = vim.tbl_filter(function(w) return not common[w] end, o.wraps or {})
        for n = #open, 1, -1 do -- (formatting it is not in: closed)
          if not vim.tbl_contains(want, open[n]) then
            for _ = n, #open do
              parts[#parts + 1] = "}"
              table.remove(open)
            end
            break
          end
        end
        spacing(y)
        for _, w in ipairs(want) do
          if not vim.tbl_contains(open, w) then
            parts[#parts + 1] = src:sub(w.s, src:find("{", w.s, true))
            open[#open + 1] = w
          end
        end
        parts[#parts + 1] = src:sub(o.s, o.e - 1)
        done_group = o.kind and o or nil
      end
      last_y = y
    end
    -- the revised tokens of the edit
    local y1, y2 = math.huge, -math.huge
    for _, c in ipairs(edit.cores) do
      y1, y2 = math.min(y1, c.n1), math.max(y2, c.n2)
    end
    if x1 <= x2 and not changed[x1] then
      y1 = math.min(y1, shift(x1))
    end
    if x1 <= x2 and not changed[x2] then
      y2 = math.max(y2, shift(x2))
    end
    local y, x = y1, x1
    while x <= x2 do
      if not changed[x] then
        local yx = shift(x)
        converted(y, yx - 1) -- (typed before it)
        source(x, yx)
        y, x = yx + 1, x + 1
      else -- a run of changes: from the revised text (a citation in it: looked up)
        local xb, prefer = x, {}
        while xb < x2 and changed[xb + 1] do
          xb = xb + 1
        end
        local ynext = xb < x2 and shift(xb + 1) or y2 + 1
        local deleted = ynext - 1 < y -- (only deleted: a value, a reference... may go with its text)
        for z = x, xb do
          local g = seq[z].owner
          if g.kind == "text" and not g.reported then -- (changed both in the source and in the document)
            g.reported, ok = true, false
            issues[#issues + 1] = vim.tbl_extend("force", where, {
              text = "conflict: the source changed here since the document was sent, and the document too: merge by hand (" .. src:sub(g.s, g.e - 1) .. ")",
            })
          elseif g.kind == "text" then
            ok = false
          elseif g.kind == "cite" then
            for key in src:sub(g.s, g.e - 1):gmatch("{([^{}]*)}%s*$") do
              vim.list_extend(prefer, vim.split(key:gsub("%s", ""), ",", { trimempty = true }))
            end
          elseif g.kind and g.kind ~= "generated" and not (g.reported or deleted) then -- (a computed value, a reference... edited)
            g.reported, ok = true, false
            issues[#issues + 1] = vim.tbl_extend("force", where, {
              text = ("%s changed in the document: the source keeps it (%s)"):format(g.kind == "sexpr" and "a computed value" or g.kind, src:sub(g.s, g.e - 1)),
            })
          elseif g.kind == "generated" and not deleted and not g.reported then
            g.reported, ok = true, false
            issues[#issues + 1] = vim.tbl_extend("force", where, { text = "the document changed text the source doesn't have (written by the export?): merge by hand" })
          elseif g.kind and g.reported and not deleted then
            ok = false
          end
        end
        converted(y, ynext - 1, prefer)
        y, x = ynext, xb + 1
      end
    end
    converted(y, y2) -- (typed at the end)
    close_all()
    for _, t in ipairs(ours) do -- (words written in the source since: not overwritten)
      if ok and t.s >= edit.A and t.e <= edit.B then
        ok = false
        issues[#issues + 1] = vim.tbl_extend("force", where, {
          text = "conflict: the source changed here since the document was sent (" .. src:sub(t.s, t.e - 1) .. "), and the document too: merge by hand",
        })
      end
    end
    if ok then
      local text = table.concat(parts)
      for _, cmd in ipairs({ "\\emph{", "\\textbf{" }) do -- ("\emph{a} \emph{b}", from two places: one)
        text = text:gsub(cmd .. "([^{}]*)}(%s+)" .. cmd, cmd .. "%1%2")
      end
      local a, b = edit.A - s + 1, edit.B - s + 1
      local function between(p, q) -- (the revised text's spacing between its tokens p and q)
        return N[p] and N[q] and table.concat(vim.tbl_map(function(ch) return escape(ch.ch) end, vim.list_slice(revised, N[p].i2 + 1, N[q].i1 - 1))) or nil
      end
      if text ~= "" then -- (rewritten at an end: the spacing there, as the revised text has it)
        local before, after = first_converted and between(y1 - 1, y1), last_converted and between(y2, y2 + 1)
        if before then
          local own = out:sub(1, a - 1):match("[%s~]*$")
          if own == "" then
            text = before .. text
          elseif before == "" then
            a = a - #own
          end
        end
        if after then
          local own = out:sub(b):match("^[%s~]*")
          if own == "" then
            text = text .. after
          elseif after == "" then
            b = b + #own
          end
        end
      else -- (a deletion: and a space next to it; a line it was all of: that line)
        if out:sub(a - 1, a - 1) == "\n" and (b > #out or out:sub(b, b) == "\n") then
          a = a - 1
        elseif b > #out and out:sub(a - 1, a - 1):match("%s") then
          a = a - #out:sub(1, a - 1):match("%s*$")
        elseif out:sub(b, b):match("%s") and (a == 1 or out:sub(a - 1, a - 1):match("[%s(]")) then
          b = b + 1
        elseif out:sub(a - 1, a - 1):match("%s") and out:sub(b, b):match("[%.,;:)%s]") then
          a = a - 1
        end
      end
      local same = out:sub(a, b - 1) == text -- (the source has it already: nothing to do, or list)
      out = out:sub(1, a - 1) .. text .. out:sub(b)
      local function short(t) -- (on one line, not too long)
        t = vim.trim(t:gsub("%s+", " "))
        return vim.fn.strchars(t) > 70 and vim.fn.strcharpart(t, 0, 69) .. "…" or t
      end
      local old, new = short(src:sub(edit.A, edit.B - 1)), short(text)
      local line_start = src:sub(1, edit.A - 1):match(".*\n()") or 1
      issues[#issues + 1] = not same and vim.tbl_extend("force", where, {
        lnum = where.lnum + select(2, src:sub(s, edit.A - 1):gsub("\n", "")),
        col = edit.A - line_start + 1,
        text = old == "" and "add: " .. new or new == "" and "delete: " .. old or ("change: %s → %s"):format(old, new),
        user_data = "edit",
      }) or nil
    end
  end
  return out, issues
end

-- ── Pairing paragraphs ──────────────────────────────────────────────────────

local function words(list)
  local bag, n = {}, 0
  for _, t in ipairs(list) do
    local w = (t.text or ""):lower()
    if w:match("%a%a%a") then
      bag[w], n = (bag[w] or 0) + 1, n + 1
    end
  end
  return bag, n
end
local function likeness(a, na, b, nb)
  local common = 0
  for w, k in pairs(a) do
    common = common + math.min(k, b[w] or 0)
  end
  return na + nb > 0 and 2 * common / (na + nb) or 0
end

-- ── Running it ──────────────────────────────────────────────────────────────

-- The source files with the document's edits: { path -> new text }, and the quickfix items.
-- The paragraphs of a document file, each its characters (read in a hidden buffer); nil if it
-- can't be read.
local loaded = {} -- (path -> { stamp, paragraphs }: a merge done again doesn't read it again)
local function load(path)
  local st = vim.uv.fs_stat(path)
  local stamp = st and ("%d:%d.%d"):format(st.size, st.mtime.sec, st.mtime.nsec)
  if stamp and loaded[path] and loaded[path][1] == stamp then
    return loaded[path][2]
  end
  local b = api.nvim_create_buf(false, true)
  local notify = vim.notify
  vim.notify = function() end -- (its errors: the caller's to report)
  pcall(neo().open, b, path)
  vim.notify = notify
  local out
  if neo().docs[b] then
    out = {}
    for _, block in ipairs(neo().blocks(b)) do
      out[#out + 1] = (versions(b, block))
    end
  end
  api.nvim_buf_delete(b, { force = true })
  loaded[path] = { stamp, out }
  return out
end

local function bag_of(list_of_chars)
  local all = {}
  for _, chars in ipairs(list_of_chars) do
    vim.list_extend(all, char_tokens(chars))
  end
  return words(all)
end

-- The words of a document file, read from its zip (no need to open it), and whether it is clean:
-- no tracked changes, as an export (comments added before sending it are fine).
local function glance(path)
  local ok, xml = pcall(require("neoffice.zip").read, path, path:match("%.odt$") and "content.xml" or "word/document.xml")
  if not ok or not xml then
    return nil
  end
  local clean = true
  for _, mark in ipairs({ "<w:ins ", "<w:del ", "<w:moveFrom ", "<w:moveTo ", "<w:rPrChange", "<text:changed-region" }) do
    clean = clean and not xml:find(mark, 1, true)
  end
  local tokens = {}
  for w in xml:gsub("<[^>]+>", " "):gmatch("[%a\128-\255]+") do
    tokens[#tokens + 1] = { text = w }
  end
  return tokens, clean
end

-- The documents (.docx, .odt) in the sources' folders and the folders in them: where an export is.
function M.documents_near(paths)
  local out = {}
  for _, dir in ipairs(vim.fn.uniq(vim.fn.sort(vim.tbl_map(vim.fs.dirname, paths)))) do
    for name, kind in vim.fs.dir(dir, { depth = 2 }) do
      local base = vim.fs.basename(name)
      if kind == "file" and (name:match("%.docx$") or name:match("%.odt$")) and not base:match("^~%$") and not base:match("^%.~lock") then
        out[#out + 1] = vim.fs.joinpath(dir, name)
      end
    end
  end
  return out
end

-- The document as it was sent, to tell its edits (of every round of review, tracked or not) from
-- what the export writes: the export itself, a .docx/.odt with no tracked changes in the sources' folders or below
-- them (as toword/), like this document, and of those the most like the sources (a collaborator's
-- copy kept there, edited, is less); or a copy named *.sent.docx beside the document. Returns its
-- path and paragraphs, or nil.
local function find_copy(buf, paths, source_bag, source_n)
  local doc = neo().docs[buf]
  local candidates, seen = {}, {}
  local function consider(path, sent)
    if not seen[path] and path ~= doc.path and #candidates < 200 then
      seen[path] = true
      candidates[#candidates + 1] = { path = path, sent = sent }
    end
  end
  for _, dir in ipairs(vim.fn.uniq(vim.fn.sort({ vim.fs.dirname(doc.path), vim.fn.getcwd(), unpack(vim.tbl_map(vim.fs.dirname, paths)) }))) do
    for name in vim.fs.dir(dir) do
      if name:match("%.sent%.docx$") or name:match("%.sent%.odt$") then
        consider(vim.fs.joinpath(dir, name), true)
      end
    end
  end
  for _, path in ipairs(M.documents_near(paths)) do
    consider(path, false)
  end
  local mine = {}
  for _, block in ipairs(neo().blocks(buf)) do
    mine[#mine + 1] = (versions(buf, block))
  end
  local bag, n = bag_of(mine)
  local best, best_score
  for _, c in ipairs(candidates) do
    local tokens, clean = glance(c.path)
    if tokens and (clean or c.sent) then
      local cbag, cn = words(tokens)
      if likeness(bag, n, cbag, cn) >= 0.5 then -- (this document)
        local score = (c.sent and 1 or 0) + likeness(source_bag, source_n, cbag, cn) -- (the export)
        if not best_score or score > best_score then
          best, best_score = c.path, score
        end
      end
    end
  end
  local paras = best and load(best)
  return paras and best or nil, paras
end

-- A source's text: of its buffer when it is loaded (with what you took from a merge and haven't
-- written yet), or else of its file.
local function source_text(path)
  for _, b in ipairs(api.nvim_list_bufs()) do
    if api.nvim_buf_is_loaded(b) and api.nvim_buf_get_name(b) == path then
      local text = table.concat(api.nvim_buf_get_lines(b, 0, -1, false), "\n")
      return text .. ((vim.bo[b].eol or vim.bo[b].fixeol) and "\n" or "")
    end
  end
  local f = assert(io.open(path), "can't read " .. path)
  local text = f:read("*a")
  f:close()
  return text
end

-- What the source would be with the document's edits: { [path] = text }, and the notes
-- (quickfix items). opts.sent: the copy of the document as sent (its path; false: none; nil:
-- looked for, see find_copy). Without it, the document's text as sent is its text with all
-- tracked changes rejected, so only tracked edits are found.
function M.compute(buf, paths, opts)
  buf, opts = buf == 0 and api.nvim_get_current_buf() or buf, opts or {}
  local doc = neo().docs[buf]
  local sources, all = {}, {}
  local bibs = {}
  for _, path in ipairs(paths) do
    local text = source_text(path)
    sources[path] = { text = text, edits = {} }
    vim.list_extend(all, regions(path, text))
    for bib in text:gmatch("\\addbibresource%b{}") do
      bibs[#bibs + 1] = vim.fs.joinpath(vim.fs.dirname(path), bib:sub(17, -2))
    end
  end
  for _, dir in ipairs(vim.tbl_map(vim.fs.dirname, paths)) do
    for name in vim.fs.dir(dir) do
      if name:match("%.bib$") then
        bibs[#bibs + 1] = vim.fs.joinpath(dir, name)
      end
    end
  end
  local bib = read_bib(vim.fn.uniq(vim.fn.sort(bibs)))
  local source_tokens = {}
  for _, path in ipairs(paths) do
    vim.list_extend(source_tokens, tex_tokens(sources[path].text, 1, #sources[path].text + 1))
  end
  local copy_path, copy = nil, nil
  if opts.sent then
    copy_path, copy = opts.sent, load(opts.sent)
    assert(copy, "can't read " .. opts.sent)
  elseif opts.sent == nil then
    local source_bag, source_n = words(source_tokens)
    copy_path, copy = find_copy(buf, paths, source_bag, source_n)
  end
  local copy_bags, copy_used = {}, {}
  for k, chars in ipairs(copy or {}) do
    copy_bags[k] = { words(char_tokens(chars)) }
  end
  for _, r in ipairs(all) do
    r.bag, r.n = words(tex_tokens(sources[r.path].text, r.s, r.e))
  end

  -- each paragraph of the document: its text as sent and as revised, and its source paragraph
  local items = {}
  for _, block in ipairs(neo().blocks(buf)) do
    local sent, revised = versions(buf, block)
    if copy then -- (the paragraph as it really was sent: edits made without tracking count too)
      local bag, n = words(char_tokens(#char_tokens(sent) > 0 and sent or revised))
      local found, score = nil, 0.5
      for k, b in ipairs(copy_bags) do
        local l = not copy_used[k] and b[2] > 0 and n > 0 and likeness(bag, n, b[1], b[2]) or 0
        if l >= score then
          found, score = k, l
        end
      end
      if found then
        sent, copy_used[found] = copy[found], true
      elseif n > 0 then -- (not in it: a paragraph added)
        sent = {}
      end
    end
    local item = { block = block, sent = sent, revised = revised, O = char_tokens(sent) }
    item.bag, item.n = words(item.O)
    items[#items + 1] = item
  end
  -- paired by their words: the longer paragraphs first; then the short ones (a heading, a title:
  -- few words, which another file may have too) only in the file of the paragraphs around them
  local used = {}
  local function pair(item, file, bag, n, least)
    local best, score = nil, 0
    for _, r in ipairs(all) do
      local l = (not file or r.path == file) and likeness(bag or item.bag, n or item.n, r.bag, r.n) or 0
      if l > score then
        best, score = r, l
      end
    end
    if best and score >= (least or 0.5) and not used[best] then
      used[best], item.best = true, best
    end
  end
  local SHORT = 8
  for _, item in ipairs(items) do
    if item.n >= SHORT then
      pair(item)
    end
  end
  for k, item in ipairs(items) do
    if item.n < SHORT then
      local file
      for d = 1, #items do -- (the nearest paired paragraph, before or after)
        local near = items[k - d] and items[k - d].n >= SHORT and items[k - d].best or items[k + d] and items[k + d].n >= SHORT and items[k + d].best
        if near then
          file = near.path
          break
        end
      end
      pair(item, file)
    end
  end
  for _, item in ipairs(items) do -- (a paragraph the document added, in the source already: a merge
    -- done again. Only when almost the same: a new paragraph is never put in place of another)
    if item.n == 0 then
      local bag, n = words(char_tokens(item.revised))
      if n > 0 then
        pair(item, nil, bag, n, 0.9)
      end
    end
  end

  local issues, previous = {}, nil
  local docpath = api.nvim_buf_get_name(buf)
  for _, item in ipairs(items) do
    local block, sent, revised, O, best = item.block, item.sent, item.revised, item.O, item.best
    local changed = table.concat(keys(O), " ") ~= table.concat(keys(char_tokens(revised)), " ")
    local where = best and { filename = best.path, lnum = best.line } or { filename = docpath, lnum = block.first }
    -- the comments on it
    for _, m in ipairs(api.nvim_buf_get_extmarks(buf, neo().ns.comment, { block.first - 1, 0 }, { block.last - 1, -1 }, {})) do
      local c = doc.comment[m[1]]
      if c and not c.resolved then -- (resolved ones were dealt with in the review)
        local thread = { c.author .. ": " .. c.text:gsub("\n", " ") }
        for _, r in ipairs(c.replies) do
          thread[#thread + 1] = r.author .. ": " .. r.text:gsub("\n", " ")
        end
        issues[#issues + 1] = vim.tbl_extend("force", where, { text = " " .. table.concat(thread, "  ↳ "), type = "I" })
      end
    end
    if best then
      previous = best
      -- (also when nothing is tracked: edits made without tracking, or accepted already, count)
      local text, found = merge_paragraph(sources[best.path].text, best.s, best.e, sent, revised, bib, where)
      if text ~= sources[best.path].text:sub(best.s, best.e - 1) then
        table.insert(sources[best.path].edits, { s = best.s, e = best.e, text = text })
      end
      vim.list_extend(issues, found)
    elseif #O == 0 and #char_tokens(revised) > 0 then -- a paragraph the document added
      local text = latex(revised, {}, bib, nil, issues, where)
      if previous then
        table.insert(sources[previous.path].edits, { s = previous.e, e = previous.e, text = "\n\n" .. text })
        issues[#issues + 1] = { filename = previous.path, lnum = previous.line, text = "new paragraph after this one: check where it goes" }
      else
        issues[#issues + 1] = vim.tbl_extend("force", where, { text = "new paragraph, not placed: " .. text:sub(1, 80) })
      end
    elseif changed then
      issues[#issues + 1] = { filename = docpath, lnum = block.first, text = "changed, but not found in the source (a table? generated by R?): " .. block.text:sub(1, 60) }
    end
  end

  for k, chars in ipairs(copy or {}) do -- paragraphs of the copy the document doesn't have: deleted
    local bag, n = copy_bags[k][1], copy_bags[k][2]
    if not copy_used[k] and n > 0 then
      local best, score = nil, 0.5
      for _, r in ipairs(all) do
        local l = not used[r] and likeness(bag, n, r.bag, r.n) or 0
        if l >= score then
          best, score = r, l
        end
      end
      if best then
        used[best] = true
        local text = sources[best.path].text
        local e = best.e
        local after = text:match("^\n[ \t]*\n", e) -- (and the blank line after it, or else before it)
        local before = not after and text:sub(1, best.s - 1):match("\n[ \t]*\n$")
        table.insert(sources[best.path].edits, { s = best.s - (before and #before or 0), e = after and e + #after or e, text = "" })
        issues[#issues + 1] = { filename = best.path, lnum = best.line, text = "paragraph deleted in the document: check it" }
      elseif table.concat(vim.tbl_map(function(c) return c.ch end, chars)):match("%a%a%a") then
        issues[#issues + 1] = { filename = docpath, lnum = 1, text = "a paragraph was deleted in the document, not found in the source: " .. table.concat(vim.tbl_map(function(c) return c.ch end, chars)):sub(1, 60) }
      end
    end
  end

  local out = {}
  for path, source in pairs(sources) do
    table.sort(source.edits, function(x, y) return x.s > y.s end)
    local text = source.text
    for _, edit in ipairs(source.edits) do
      text = text:sub(1, edit.s - 1) .. edit.text .. text:sub(edit.e)
    end
    out[path] = text
  end
  table.sort(issues, function(x, y) -- (those in the sources first, those only in the document last)
    if (x.filename == docpath) ~= (y.filename == docpath) then
      return y.filename == docpath
    elseif x.filename ~= y.filename then
      return x.filename < y.filename
    end
    return x.lnum < y.lnum or x.lnum == y.lnum and (x.col or 0) < (y.col or 0)
  end)
  return out, issues, copy_path
end

local proposals = {} -- (each source's proposal buffer, from the last merge)

-- <CR> in the quickfix list: to the item, in the window that shows its file already (your file, the
-- proposal, the document in its tab), so the diff stays as it is; in a new tab if none does.
function M.jump()
  local n = vim.fn.line(".")
  local item = vim.fn.getqflist()[n]
  if not item or not vim.fn.getqflist({ title = true }).title:find("^neoffice merge") or item.bufnr == 0 then
    return vim.cmd(n .. "cc") -- (another list now)
  end
  local here = vim.tbl_filter(function(w) return api.nvim_win_get_buf(w) == item.bufnr end, api.nvim_tabpage_list_wins(0))
  local win = here[1] or vim.fn.win_findbuf(item.bufnr)[1]
  if win then
    api.nvim_set_current_win(win)
  else
    vim.cmd("tab sbuffer " .. item.bufnr)
  end
  local last = api.nvim_buf_line_count(0)
  api.nvim_win_set_cursor(0, { math.min(math.max(item.lnum, 1), last), math.max(item.col - 1, 0) })
  vim.fn.setqflist({}, "a", { idx = n }) -- (the list's current item: this one, for :cnext)
end

-- The sources merged into before (a list of lists of paths, the last first), so that a document
-- opened anywhere (in Downloads, say) finds its manuscript.
M.memory = vim.fs.joinpath(vim.fn.stdpath("state"), "neoffice-merge.json")
local function remembered()
  local f = io.open(M.memory)
  local ok, list = pcall(vim.json.decode, f and f:read("*a") or "[]")
  if f then
    f:close()
  end
  return ok and type(list) == "table" and list or {}
end
local function remember(paths)
  local key = table.concat(vim.fn.sort(vim.deepcopy(paths)), "\n")
  local list = { paths }
  for _, old in ipairs(remembered()) do
    if #list < 20 and table.concat(vim.fn.sort(vim.deepcopy(old)), "\n") ~= key then
      list[#list + 1] = old
    end
  end
  vim.fn.mkdir(vim.fs.dirname(M.memory), "p")
  local f = io.open(M.memory, "w")
  if f then
    f:write(vim.json.encode(list))
    f:close()
  end
end
-- (of those, the one that has most of the document: its pairs of words, as single words are
-- shared by any two texts on a subject)
local function word_pairs(tokens)
  local bag, n, last = {}, 0, nil
  for _, t in ipairs(tokens) do
    local w = (t.text or ""):lower()
    if w:match("^%a+$") then
      if last then
        local key = last .. " " .. w
        bag[key], n = (bag[key] or 0) + 1, n + 1
      end
      last = w
    elseif not (t.text or ""):match("^%p$") then -- (a citation, a value...: a break)
      last = nil
    end
  end
  return bag, n
end
local function recall(buf)
  local mine = {}
  for _, block in ipairs(neo().blocks(buf)) do
    vim.list_extend(mine, char_tokens((versions(buf, block))))
  end
  local bag, n = word_pairs(mine)
  local function share(tokens) -- (how much of the document is in them)
    local theirs, common = word_pairs(tokens), 0
    for w, k in pairs(bag) do
      common = common + math.min(k, theirs[w] or 0)
    end
    return n > 0 and common / n or 0
  end
  local scores = {}
  for _, list in ipairs(remembered()) do
    if #list > 0 and vim.iter(list):all(function(p) return vim.uv.fs_stat(p) ~= nil end) then
      local tokens = {}
      for _, path in ipairs(list) do
        local text = source_text(path)
        vim.list_extend(tokens, tex_tokens(text, 1, #text + 1))
      end
      local score = share(tokens)
      for _, path in ipairs(M.documents_near(list)) do -- (its exports: their citations, values... written)
        local doc_tokens, clean = glance(path)
        if doc_tokens and clean and path ~= neo().docs[buf].path then
          score = math.max(score, share(doc_tokens))
        end
      end
      scores[#scores + 1] = { list = list, score = score }
    end
  end
  -- (the most like it, and of those as alike, the one with more files: appendices too)
  table.sort(scores, function(x, y)
    return x.score - y.score > 0.02 or math.abs(x.score - y.score) <= 0.02 and #x.list > #y.list
  end)
  local best = scores[1]
  local second = best and vim.iter(scores):find(function(x) -- (another manuscript: no file in common)
    return not vim.iter(x.list):any(function(p) return vim.tbl_contains(best.list, p) end)
  end)
  -- (clearly this one: else it asks)
  return best and best.score >= 0.6 and (not second or best.score - second.score >= 0.2) and best.list or nil
end

local function sources_in(dir) -- (its .Rnw files, or else its .tex files)
  for _, pattern in ipairs({ "*.[Rr]nw", "*.tex" }) do
    local found = vim.fn.glob(vim.fs.joinpath(dir, pattern), false, true)
    if #found > 0 then
      return found
    end
  end
  return {}
end

-- A citation to look up: its quickfix item goes to its text in the proposal.
local function place_citations(issues, path, proposal)
  for _, item in ipairs(issues) do
    if item.cite and item.filename == path then
      local word = vim.pesc(item.cite.name:match("[^%s,&]+") or "")
      local best
      for n, line in ipairs(api.nvim_buf_get_lines(proposal, 0, -1, false)) do
        local col = line:find(word .. "[^()]-" .. item.cite.year)
        if col and (not best or math.abs(n - item.lnum) < math.abs(best[1] - item.lnum)) then
          best = { n, col }
        end
      end
      if best then
        item.filename, item.bufnr, item.lnum, item.col = nil, proposal, best[1], best[2]
      end
    end
  end
end

local function summary(files, issues, copy)
  local edits = #vim.tbl_filter(function(i) return i.user_data == "edit" end, issues)
  return ("%d file(s) to review, %d edit(s), %d note(s); %s"):format(files, edits, #issues - edits,
    copy and "as sent: " .. vim.fn.fnamemodify(copy, ":~:.") or "no export found: tracked edits only")
end

-- While you review, the quickfix list follows the sources: an edit you take (do), or make yourself,
-- leaves it (a moment after each change, or :w).
local timer -- (of the last merge)
local saved_switchbuf -- (yours, while a merge tab is open)
local function follow(buf, paths, copy, qf)
  if timer and not timer:is_closing() then
    timer:close()
  end
  timer = assert(vim.uv.new_timer())
  local mine, group = timer, api.nvim_create_augroup("neoffice_merge", { clear = true })
  -- while a merge tab is open, <CR>, :cnext and the like go to the window that shows the file
  -- already (in whichever tab), or else to a new tab: so the diff stays. 'switchbuf' then, yours
  -- again when no merge tab is left
  if not saved_switchbuf then
    saved_switchbuf, vim.o.switchbuf = vim.o.switchbuf, "useopen,usetab,newtab"
  end
  api.nvim_create_autocmd({ "TabClosed", "BufWipeout" }, {
    group = api.nvim_create_augroup("neoffice_merge_tabs", { clear = true }), -- (kept when the list goes)
    callback = vim.schedule_wrap(function()
      local open = vim.iter(vim.tbl_values(proposals)):any(function(b)
        return api.nvim_buf_is_valid(b) and #vim.fn.win_findbuf(b) > 0
      end)
      if not open and saved_switchbuf then
        vim.o.switchbuf, saved_switchbuf = saved_switchbuf, nil
      end
    end),
  })
  local function refresh()
    if vim.fn.getqflist({ id = qf }).id ~= qf or not neo().docs[buf] then -- (its list gone, or the document closed)
      mine:close()
      return pcall(api.nvim_del_augroup_by_id, group)
    end
    local ok, out, issues = pcall(M.compute, buf, paths, { sent = copy or false })
    if not ok then
      return
    end
    local files = 0
    for _, path in ipairs(paths) do
      files = files + (out[path] ~= source_text(path) and 1 or 0)
      if proposals[path] and api.nvim_buf_is_valid(proposals[path]) then
        place_citations(issues, path, proposals[path])
      end
    end
    vim.fn.setqflist({}, "r", { id = qf, items = issues, title = "neoffice merge: " .. summary(files, issues, copy) })
  end
  for _, b in ipairs(api.nvim_list_bufs()) do -- (the sources open: those with a merge tab)
    if api.nvim_buf_is_loaded(b) and vim.tbl_contains(paths, api.nvim_buf_get_name(b)) then
      api.nvim_create_autocmd({ "TextChanged", "TextChangedI", "BufWritePost" }, {
        group = group,
        buffer = b,
        callback = function()
          if mine:is_closing() then
            return
          end
          mine:stop()
          mine:start(300, 0, vim.schedule_wrap(function()
            if not mine:is_closing() then
              refresh()
            end
          end))
        end,
      })
    end
  end
end

-- :Neoffice merge {files}: each changed file opens in a tab, the source beside what it would be
-- with the document's edits (Neovim's diff: ]c, do, dp); the comments and what needs a human go to
-- the quickfix list.
function M.run(paths)
  local buf = api.nvim_get_current_buf()
  if not neo().docs[buf] then
    return vim.notify("[neoffice] Not a document buffer", vim.log.levels.WARN)
  end
  local sent, given = nil, {} -- (a document among them: the document as sent; a folder: its sources)
  for _, p in ipairs(paths) do
    p = vim.fn.fnamemodify(vim.fs.normalize(p), ":p"):gsub("/$", "")
    if p:match("%.docx$") or p:match("%.odt$") then
      sent = p
    elseif vim.fn.isdirectory(p) == 1 then
      vim.list_extend(given, sources_in(p))
    else
      given[#given + 1] = p
    end
  end
  paths = #given > 0 and given or sources_in(vim.fn.getcwd()) -- (none: those of the current folder,
  if #paths == 0 then -- or of a manuscript merged into before, or ask)
    paths = recall(buf) or {}
  end
  if #paths == 0 then
    return vim.ui.input({ prompt = "Source of this document (.Rnw/.tex, or their folder): ", completion = "file" }, function(answer)
      if answer and vim.trim(answer) ~= "" then
        api.nvim_set_current_buf(buf)
        M.run(vim.list_extend({ vim.fn.expand(vim.trim(answer)) }, sent and { sent } or {}))
      end
    end)
  end
  local ok, out, issues, copy = pcall(M.compute, buf, paths, { sent = sent })
  if not ok then
    return vim.notify("[neoffice] Merge failed: " .. out, vim.log.levels.ERROR)
  end
  remember(paths)
  local opened = 0
  for _, path in ipairs(paths) do
    local now = source_text(path)
    -- the merge tab of this source, from before: refreshed (if it still shows the source)
    local proposal, source_win = proposals[path], nil
    if proposal and api.nvim_buf_is_valid(proposal) then
      local win = vim.fn.win_findbuf(proposal)[1]
      for _, w in ipairs(win and api.nvim_tabpage_list_wins(api.nvim_win_get_tabpage(win)) or {}) do
        source_win = api.nvim_buf_get_name(api.nvim_win_get_buf(w)) == path and w or source_win
      end
      if not source_win then
        api.nvim_buf_delete(proposal, { force = true })
      end
    end
    proposal = source_win and proposal or nil
    if out[path] and (out[path] ~= now or proposal) then
      local name = path .. " (with the document's edits)"
      if proposal then
        api.nvim_set_current_win(source_win)
      else
        for _, b in ipairs(api.nvim_list_bufs()) do -- (one left from an earlier merge, not shown)
          if api.nvim_buf_get_name(b) == name then
            api.nvim_buf_delete(b, { force = true })
          end
        end
        vim.cmd("tabnew " .. vim.fn.fnameescape(path))
        source_win = api.nvim_get_current_win()
        vim.cmd("diffthis")
        local ft = vim.bo.filetype
        vim.cmd("vertical leftabove new") -- (the proposal on the left, yours on the right, as fugitive does)
        proposal = api.nvim_get_current_buf()
        proposals[path] = proposal
        vim.bo[proposal].buftype, vim.bo[proposal].bufhidden, vim.bo[proposal].swapfile = "nofile", "wipe", false
        api.nvim_buf_set_name(proposal, name)
        vim.bo[proposal].filetype = ft
        vim.cmd("diffthis")
        vim.wo.wrap = true -- (after :diffthis, which turns it off)
        api.nvim_set_current_win(source_win)
        vim.wo.wrap = true
      end
      api.nvim_buf_set_lines(proposal, 0, -1, false, vim.split(out[path]:gsub("\n$", ""), "\n", { plain = true }))
      place_citations(issues, path, proposal)
      vim.cmd("diffupdate")
      opened = opened + (out[path] ~= now and 1 or 0)
    end
  end
  local text = summary(opened, issues, copy)
  vim.fn.setqflist({}, " ", { title = "neoffice merge: " .. text, items = issues })
  follow(buf, paths, copy, vim.fn.getqflist({ id = 0 }).id)
  if #issues > 0 then
    vim.cmd("botright copen")
    vim.keymap.set("n", "<CR>", M.jump, { buffer = true, desc = "Go to it (neoffice merge: where it is shown)" })
    vim.cmd("wincmd p")
  end
  local blind = not copy and #api.nvim_buf_get_extmarks(buf, neo().ns.change, 0, -1, { limit = 1 }) == 0
  vim.schedule(function() -- (after the redraw: seen)
    if blind then -- (nothing tells what changed: accepted already, or edited without tracking)
      return vim.notify("[neoffice] Merge: the document has no tracked changes, and the export it was made from wasn't found "
        .. "(a .docx with no tracked changes, in the sources' folder or below), so its edits can't be told apart. "
        .. "Give it: :Neoffice merge {export.docx}", vim.log.levels.WARN)
    end
    vim.notify("[neoffice] Merge: " .. text)
  end)
end

return M
