-- Tracked changes. Inserted text is highlighted; deleted text stays in the buffer, struck through
-- (a paragraph break in it shows as ¶). A change is
--   { type = "insert" | "delete" | "format", author, date, text (a deletion's, as in the document) }.
--
-- Reviewing is one rule: accepting keeps what you see (inserted text stays, deleted text goes),
-- rejecting goes back (inserted text goes, deleted text stays). It acts on the change at the
-- cursor, on the selected part of every change a visual selection touches, or on all of them.
--
-- Every state of the buffer remembers its changes, so u and <C-r> bring them back with the text.
-- While tracking is on, your own edits are recorded as changes too (see M.record).
local markup = require("neoffice.markup")
local api = vim.api
local M = {}

local function neo()
  return require("neoffice")
end

local function now()
  return os.date("%Y-%m-%dT%H:%M:%S")
end

local function before(a, b) -- positions { row, col }: a <= b
  return a[1] < b[1] or a[1] == b[1] and a[2] <= b[2]
end

local function same(a, b)
  return a[1] == b[1] and a[2] == b[2]
end

-- Each author gets a colour (of six: NeofficeInsert1..6, NeofficeDelete1..6), in order of appearance.
local authors = {}
local function color(author)
  if not authors[author] then
    authors[#authors + 1] = author
    authors[author] = (#authors - 1) % 6 + 1
  end
  return authors[author]
end

function M.style(ch)
  local n = color(ch.author)
  local hl = { insert = "NeofficeInsert" .. n, delete = "NeofficeDelete" .. n, format = "NeofficeFormat" }
  return { hl_group = hl[ch.type] }
end

-- The authors of a document's changes, sorted.
function M.authors(buf)
  local doc = neo().docs[buf or api.nvim_get_current_buf()]
  local names = {}
  for _, ch in pairs(doc and doc.change or {}) do
    names[ch.author] = true
  end
  names = vim.tbl_keys(names)
  table.sort(names)
  return names
end

local function marks(buf, start, stop, opts)
  return api.nvim_buf_get_extmarks(buf, neo().ns.change, start, stop, vim.tbl_extend("force", { details = true }, opts or {}))
end

local function range(m) -- start and end of a mark
  return { m[2], m[3] }, { m[4].end_row, m[4].end_col }
end

local function text(buf, s, e)
  return table.concat(api.nvim_buf_get_text(buf, s[1], s[2], e[1], e[2], {}), "\n")
end

local function mine(ch) -- one of your own changes not written to the document yet
  return ch.author == neo().config.author and not ch.node and not ch.region
end

-- ── Every state remembers its changes (for undo and redo) ───────────────────

local function capture(buf)
  local doc = neo().docs[buf]
  return vim.tbl_map(function(m)
    local ch = doc.change[m[1]]
    return { m[2], m[3], m[4].end_row, m[4].end_col, ch, ch.shown }
  end, marks(buf, 0, -1))
end

local function restore(buf, state)
  local doc = neo().docs[buf]
  api.nvim_buf_clear_namespace(buf, neo().ns.change, 0, -1)
  doc.change = {}
  for _, entry in ipairs(state) do
    neo().place(buf, "change", entry[5], entry[1], entry[2], entry[3], entry[4])
    entry[5].shown = entry[6]
  end
end

-- Deleted text: what is shown is what is deleted. If text was typed right before it (it would
-- start the deletion), the deletion starts after that text; if all of it was deleted, so is the change.
-- Your own insertions deleted entirely go, and your own deletions next to each other become one.
local function tidy(buf)
  local doc = neo().docs[buf]
  local own_deletion
  for _, m in ipairs(marks(buf, 0, -1)) do
    local ch = doc.change[m[1]]
    local s, e = range(m)
    local shown = text(buf, s, e)
    local typed = ch.shown and #shown > #ch.shown and shown:sub(-#ch.shown) == ch.shown
    if (ch.type == "delete" or mine(ch) and ch.type == "insert") and shown == "" then
      api.nvim_buf_del_extmark(buf, neo().ns.change, m[1])
      doc.change[m[1]] = nil
    elseif ch.type == "delete" and typed then -- (where its text starts: as many lines up from its end)
      local rows = vim.split(ch.shown, "\n", { plain = true })
      local row = e[1] - #rows + 1
      local width = #rows == 1 and e[2] or #api.nvim_buf_get_lines(buf, row, row + 1, true)[1]
      s = { row, width - #rows[1] }
      api.nvim_buf_set_extmark(buf, neo().ns.change, s[1], s[2], vim.tbl_extend("force", M.style(ch), { id = m[1], end_row = e[1], end_col = e[2], right_gravity = false }))
    elseif ch.type == "delete" and own_deletion and mine(ch) and same(own_deletion.e, s) then
      api.nvim_buf_del_extmark(buf, neo().ns.change, m[1])
      doc.change[m[1]] = nil
      local first = own_deletion
      api.nvim_buf_set_extmark(buf, neo().ns.change, first.s[1], first.s[2], vim.tbl_extend("force", M.style(first.ch), { id = first.id, end_row = e[1], end_col = e[2], right_gravity = false }))
      first.e, first.ch.shown = e, text(buf, first.s, e)
      goto continue
    else
      ch.shown = shown
    end
    own_deletion = ch.type == "delete" and mine(ch) and { id = m[1], s = s, e = e, ch = ch } or nil
    ::continue::
  end
end

-- ── Recording your edits ────────────────────────────────────────────────────

-- Remember the text and the changes, to tell what the next edit did (with tracking off too: what
-- you type is then kept out of the changes around it).
function M.snapshot(buf)
  local doc = neo().docs[buf]
  doc.window = nil
  local shadow = { lines = api.nvim_buf_get_lines(buf, 0, -1, false), inserted = {}, deletions = {} }
  for _, m in ipairs(marks(buf, 0, -1)) do
    local ch, s, e = doc.change[m[1]], range(m)
    if ch.type == "insert" and ch.author == neo().config.author then
      table.insert(shadow.inserted, { s, e })
    elseif ch.type == "delete" then
      table.insert(shadow.deletions, { s, e, ch = ch })
    end
  end
  doc.shadow = shadow
end

-- The characters of lines [first, last] as { row, col, ch, raw }, with "\n" between lines. An
-- image or equation is one character: its text as written (raw: not to be escaped again).
local function characters(lines, first, last)
  local out = {}
  for row = first, last do
    local _, chars = markup.parse(lines[row])
    for _, c in ipairs(chars) do
      local token = c.len and lines[row]:sub(c.col + 1, c.col + c.len)
      out[#out + 1] = { row - 1, c.col, token or c.ch, token and true }
    end
    local final = out[#out]
    if final and #(lines[row]:match("\\*$")) % 2 == 1 and final[2] == #lines[row] - 1 then
      final[4] = true -- (a "\" ending a line is a line break: a token too)
    end
    if row < last then
      out[#out + 1] = { row - 1, #lines[row], "\n" }
    end
  end
  return out
end

-- The characters of `lines` in the span that starts `lo` bytes from the start of the text and ends
-- `suffix` bytes before its end (a character partly in it, like an image's text, counts whole),
-- and the position where the span ends.
local function window(lines, lo, suffix)
  local starts, total = {}, 0
  for r, l in ipairs(lines) do
    starts[r], total = total, total + #l + 1
  end
  local hi = math.max(total - suffix, lo)
  local function position(offset)
    for r = #lines, 1, -1 do
      if starts[r] <= offset then
        return { r - 1, math.min(offset - starts[r], #lines[r]) }
      end
    end
    return { 0, 0 }
  end
  local first, last = position(lo), position(math.max(hi - 1, lo))
  local chars = vim.tbl_filter(function(c)
    local offset = starts[c[1] + 1] + c[2]
    return offset < hi and offset + #c[3] > lo
  end, characters(lines, first[1] + 1, math.min(last[1] + 2, #lines))) -- (+1: its line break)
  return chars, position(hi)
end

local function after(c) -- the position right after character c
  return c[3] == "\n" and { c[1] + 1, 0 } or { c[1], c[2] + #c[3] }
end

local function within(r, c) -- character c inside range r = { start, end }
  return before(r[1], c) and not before(r[2], c)
end

-- You deleted old characters [i, j], which were at `at`: text you had inserted yourself just goes;
-- the rest is put back where it was (line breaks too), struck through: deleted by you, or still
-- deleted by whoever deleted it. Deleting only line breaks (blank lines) isn't recorded.
-- Returns the position right after what was put back.
local function deleted(buf, old, chars, i, j, at)
  local doc = neo().docs[buf]
  local parts = {} -- { who = change (or "new"), text = ... }
  for k = i, j do
    local c, who = chars[k], "new"
    if c[4] then -- an image, equation, footnote, field or break: put back (they aren't tracked as deleted)
      who = "keep"
      vim.schedule(function()
        vim.notify("[neoffice] With tracking on, images, equations, footnotes, fields and breaks can't be"
          .. " deleted (they are put back): turn tracking off to delete them", vim.log.levels.WARN)
      end)
    elseif vim.iter(old.inserted):any(function(r) return within(r, c) end) then
      who = nil
    else
      local d = vim.iter(old.deletions):find(function(r) return within(r, c) end)
      who = d and d.ch or "new"
    end
    if who then -- (as it was written: escaped again, except images and equations)
      local last = parts[#parts]
      local piece = c[4] and c[3] or c[3] == "\n" and "\n" or markup.escape(c[3])
      if last and last.who == who then
        last.text = last.text .. piece
      else
        parts[#parts + 1] = { who = who, text = piece }
      end
    end
  end
  if not vim.iter(parts):any(function(part) return part.text:find("[^\n]") end) then
    return at
  end
  for _, m in ipairs(marks(buf, at, at)) do -- (deletions emptied by this edit: they come back below)
    local ch = doc.change[m[1]]
    if ch.type == "delete" and same({ m[2], m[3] }, { m[4].end_row, m[4].end_col }) then
      api.nvim_buf_del_extmark(buf, neo().ns.change, m[1])
      doc.change[m[1]] = nil
    end
  end
  local all = {}
  for _, part in ipairs(parts) do
    all[#all + 1] = part.text
  end
  pcall(vim.cmd, "undojoin") -- (one undo step with your edit)
  api.nvim_buf_set_text(buf, at[1], at[2], at[1], at[2], vim.split(table.concat(all), "\n"))
  local pos = at
  for _, part in ipairs(parts) do
    local lines = vim.split(part.text, "\n")
    local stop = { pos[1] + #lines - 1, (#lines == 1 and pos[2] or 0) + #lines[#lines] }
    if part.who ~= "keep" then
      local ch = part.who == "new" and { type = "delete", author = neo().config.author, date = now() } or part.who
      neo().place(buf, "change", ch, pos[1], pos[2], stop[1], stop[2])
      ch.shown = part.text
    end
    pos = stop
  end
  if api.nvim_get_current_buf() == buf and not api.nvim_get_mode().mode:match("^i") then
    local line = api.nvim_buf_get_lines(buf, pos[1], pos[1] + 1, true)[1]
    api.nvim_win_set_cursor(0, { pos[1] + 1, math.min(pos[2], math.max(#line - 1, 0)) }) -- (past it: x goes on)
  end
  return pos
end

-- Text you inserted, from s to e (current positions). It is never part of someone else's
-- insertion or of a deletion (nor, with tracking off, of any change): typed inside (or at the start
-- of) one, it splits it in two around it. With tracking on, it is an insertion of yours.
local function inserted(buf, s, e)
  local doc, me = neo().docs[buf], neo().config.author
  for _, m in ipairs(marks(buf, { s[1], 0 }, { e[1], -1 }, { overlap = true })) do
    local ch, ms, me_ = doc.change[m[1]], range(m)
    local foreign = ch.type == "delete" or ch.type == "insert" and (ch.author ~= me or not doc.tracking)
    if foreign and before(ms, e) and before(s, me_) and not same(s, me_) and not same(e, ms) then
      api.nvim_buf_del_extmark(buf, neo().ns.change, m[1])
      doc.change[m[1]] = nil
      local kept = before(ms, s) and not same(ms, s)
      if kept then
        neo().place(buf, "change", ch, ms[1], ms[2], s[1], s[2])
      end
      if before(e, me_) and not same(e, me_) then -- (the rest of theirs: a change of its own)
        local rest = kept and { type = ch.type, author = ch.author, date = ch.date } or ch
        neo().place(buf, "change", rest, e[1], e[2], me_[1], me_[2])
      end
    end
  end
  if not doc.tracking then -- (just text)
    return
  end
  for _, m in ipairs(marks(buf, { s[1], 0 }, { e[1], -1 }, { overlap = true })) do
    local ch, ms, me_ = doc.change[m[1]], range(m)
    if ch.author == me and ch.type == "insert" and before(ms, e) and before(s, me_) then
      -- inside or touching one of your insertions: it grows
      local from, to = before(ms, s) and ms or s, before(e, me_) and me_ or e
      return api.nvim_buf_set_extmark(buf, neo().ns.change, from[1], from[2], vim.tbl_extend("force", M.style(ch), { id = m[1], end_row = to[1], end_col = to[2], right_gravity = false }))
    end
  end
  neo().place(buf, "change", { type = "insert", author = me, date = now() }, s[1], s[2], e[1], e[2])
end

-- Compare the text with the snapshot and record what changed.
function M.record(buf)
  local doc = neo().docs[buf]
  local old = doc and doc.shadow
  if not old then
    return
  end
  local new = api.nvim_buf_get_lines(buf, 0, -1, false)
  local edits = {}
  -- Compare the characters before and after the edit (A, B), and record what changed; `stop` is
  -- where the end of B is, for what was deleted after its last character.
  local function compare(A, B, stop, whole)
    -- an image, equation, footnote... edited inside (alt text, width, LaTeX, note) is the same
    -- object, edited: not a deletion and an insertion (they aren't tracked: see deleted())
    local function same_object(a, b)
      return a and b and a[4] and b[4] and a[3]:sub(1, 1) == b[3]:sub(1, 1)
    end
    A, B = vim.list_slice(A), vim.list_slice(B)
    if same_object(A[1], B[1]) and A[1][1] == B[1][1] and A[1][2] == B[1][2] then
      table.remove(A, 1)
      table.remove(B, 1)
    end
    if same_object(A[#A], B[#B]) then
      table.remove(A)
      table.remove(B)
    end
    local function join(chars, from, to)
      return table.concat(vim.tbl_map(function(c) return c[3] end, vim.list_slice(chars, from, to)))
    end
    local function lines(chars) -- one character per line, for the diff
      return vim.tbl_map(function(c) return c[3] == "\n" and "\1" or c[3] end, chars)
    end
    -- a change inside a word replaces the whole word ("Last" -> "First", not "La" -> "Fir" keeping
    -- "st"; "Last" -> "Final", not "L"/"Fin" + "st"/"l")
    local hunks = {}
    if whole and join(A, 1, #A) ~= join(B, 1, #B) then -- (unless it is the same text again)
      hunks[1] = { 1, #A, 1, #B }
    end
    local function letter(c) -- (part of a word: not a space, line break, image...)
      return c and not c[4] and not c[3]:find("^%s")
    end
    for _, d in ipairs(whole and {} or neo().diff(lines(A), lines(B))) do
      local last = hunks[#hunks]
      local gap = last and join(A, last[1] + last[2], d[1] - 1)
      if gap and not gap:find("%s") then
        last[2], last[4] = d[1] + d[2] - last[1], d[3] + d[4] - last[3]
      else
        hunks[#hunks + 1] = d
      end
    end
    for k, d in ipairs(whole and {} or hunks) do -- (to the ends of its words: letters both share)
      local floor = k > 1 and hunks[k - 1][1] + hunks[k - 1][2] or 1
      while d[1] > floor and letter(A[d[1] - 1]) and letter(B[d[3] - 1]) and A[d[1] - 1][3] == B[d[3] - 1][3]
        and (letter(A[d[1]]) or letter(B[d[3]])) do
        d[1], d[2], d[3], d[4] = d[1] - 1, d[2] + 1, d[3] - 1, d[4] + 1
      end
      local ceiling = hunks[k + 1] and hunks[k + 1][1] or #A + 1
      while d[1] + d[2] < ceiling and letter(A[d[1] + d[2]]) and letter(B[d[3] + d[4]]) and A[d[1] + d[2]][3] == B[d[3] + d[4]][3]
        and (letter(A[d[1] + d[2] - 1]) or letter(B[d[3] + d[4] - 1])) do
        d[2], d[4] = d[2] + 1, d[4] + 1
      end
    end
    for _, d in ipairs(hunks) do
      local i, ci, j, cj = unpack(d)
      local at = B[j] and { B[j][1], B[j][2] } or stop
      -- (only line breaks inserted, a blank line or a split paragraph: not recorded, as for deleting them)
      local typed = cj > 0 and join(B, j, j + cj - 1):find("[^\n]")
      table.insert(edits, { A = A, i = i, ci = ci, at = at, s = typed and at, e = typed and after(B[j + cj - 1]) })
    end
  end
  -- exactly where Neovim says the text changed (see M.start): nothing else. (Not for lines added or
  -- removed at the very end: Neovim counts a line break after the last line, the characters don't.)
  if doc.window and doc.window[2] > 0 then
    local lo, suffix = unpack(doc.window)
    local A = window(old.lines, lo, suffix)
    local B, new_end = window(new, lo, suffix)
    compare(A, B, new_end, doc.window[3])
  else -- (or else, where a comparison of the lines finds it)
    for _, h in ipairs(neo().diff(old.lines, new)) do
      -- the characters of the changed lines, plus one unchanged line around them
      local a1, a2 = math.max(h[1] - 1, 1), math.min(h[1] + h[2], #old.lines)
      local b1, b2 = math.max(h[3] - 1, 1), math.min(h[3] + h[4], #new)
      compare(characters(old.lines, a1, a2), characters(new, b1, b2), { b2 - 1, #new[b2] })
    end
  end
  for k = #edits, 1, -1 do -- (back to front: putting deleted text back moves what follows)
    local edit = edits[k] -- (a replacement: the deleted text comes first, then the inserted)
    local put_back = doc.tracking and edit.ci > 0 -- (with tracking off, deleted text just goes)
    local stop = put_back and deleted(buf, old, edit.A, edit.i, edit.i + edit.ci - 1, edit.at) or edit.at
    if edit.s then -- (moved by what was put back before it)
      local rows, cols = stop[1] - edit.at[1], stop[2] - edit.at[2]
      local function shift(p)
        return { p[1] + rows, p[1] == edit.at[1] and p[2] + cols or p[2] }
      end
      inserted(buf, shift(edit.s), shift(edit.e))
    end
  end
end

-- After every edit: back at an earlier state (undo, redo)? Its changes as they were then.
-- Otherwise, record your edit (when tracking) and tidy up.
function M.after_edit(buf)
  local doc = neo().docs[buf]
  if not doc then
    return
  end
  local seq = vim.fn.undotree(buf).seq_cur
  if doc.states[seq] and seq ~= doc.seq then
    restore(buf, doc.states[seq])
  else
    M.record(buf)
    tidy(buf)
    seq = vim.fn.undotree(buf).seq_cur
  end
  doc.seq, doc.states[seq] = seq, capture(buf)
  M.snapshot(buf)
  M.show_info(buf)
end

-- Neovim tells where each edit is: the span of the text it changed, as how much is unchanged at the
-- start and at the end (doc.window = { bytes before, bytes after }, until the next snapshot). So an
-- edit is never taken for another that gives the same text (typing "is " after "is " is the same
-- text as typing it before), which comparing the text alone can't tell apart. w[3]: the edits so far
-- make one span (typing, cw, x, a paste...): what was there is replaced by what is there now, as a
-- whole; several separate edits (:s///g) are compared character by character, between them only.
local function watch(buf)
  if vim.b[buf].neoffice_watched then
    return
  end
  vim.b[buf].neoffice_watched = true
  api.nvim_buf_attach(buf, false, {
    on_bytes = function(_, b, _, _, _, start, _, _, old_len, _, _, new_len)
      local doc = neo().docs[b]
      if doc then
        local total = api.nvim_buf_get_offset(b, api.nvim_buf_line_count(b))
        local before, after_, w = start, total - start - new_len, doc.window
        if w then -- (one span still, if this edit touches the span so far: w[3])
          local span_end = total - new_len + old_len - w[2]
          local touches = start <= span_end and start + old_len >= w[1]
          doc.window = { math.min(w[1], before), math.min(w[2], after_), w[3] and touches }
        else
          doc.window = { before, after_, true }
        end
      end
    end,
    on_detach = function(_, b)
      vim.b[b].neoffice_watched = nil
    end,
  })
end

function M.start(buf) -- (when a document is opened)
  local doc = neo().docs[buf]
  watch(buf)
  tidy(buf)
  doc.seq, doc.states = vim.fn.undotree(buf).seq_cur, {}
  doc.states[doc.seq] = capture(buf)
  M.snapshot(buf)
end

-- arg: "on", "off" or nil (toggle). Only for this session: the document itself isn't changed
-- (tracking starts on when the document has it on).
function M.track(arg)
  local buf = api.nvim_get_current_buf()
  local doc = neo().docs[buf]
  if not doc then
    return vim.notify("[neoffice] Not a document buffer", vim.log.levels.WARN)
  end
  doc.tracking = arg == "on" or (arg ~= "off" and not doc.tracking)
  M.snapshot(buf)
  vim.notify("[neoffice] Tracking changes " .. (doc.tracking and "on" or "off"))
end

-- ── Reviewing ───────────────────────────────────────────────────────────────

local NS_INFO = api.nvim_create_namespace("neoffice_change_info")

-- Who made the change under the cursor, and when: at the end of the line.
function M.show_info(buf)
  api.nvim_buf_clear_namespace(buf, NS_INFO, 0, -1)
  local doc = neo().docs[buf]
  if not doc or api.nvim_get_current_buf() ~= buf then
    return
  end
  local row, col = unpack(api.nvim_win_get_cursor(0))
  local mode = api.nvim_get_mode().mode
  if (mode:match("^i") or mode:match("^R")) and col > 0 then -- (typing: the character just typed)
    local line = api.nvim_get_current_line():sub(1, col)
    col = col - #line:match("[^\128-\191][\128-\191]*$")
  end
  for _, m in ipairs(marks(buf, { row - 1, col }, { row - 1, col }, { overlap = true })) do
    local ch = doc.change[m[1]]
    local s, e = range(m)
    if ch and before(s, { row - 1, col }) and (before({ row - 1, col + 1 }, e) or same(s, e)) then
      local kind = ({ insert = "inserted", delete = "deleted", format = "formatting" })[ch.type]
      local info = ("  %s · %s · %s"):format(ch.author, ch.date:sub(1, 10), kind)
      return api.nvim_buf_set_extmark(buf, NS_INFO, row - 1, 0, { virt_text = { { info, "NeofficeChangeInfo" } }, virt_text_pos = "eol" })
    end
  end
end

-- The change under the cursor, or the nearest one on the cursor line.
local function at_cursor(buf)
  local row, col = unpack(api.nvim_win_get_cursor(0))
  local best, best_dist
  for _, m in ipairs(marks(buf, { row - 1, 0 }, { row - 1, -1 }, { overlap = true })) do
    local d = m[4]
    local dist = (m[2] == row - 1 and m[3] > col and m[3] - col) or (d.end_row == row - 1 and d.end_col <= col and col - d.end_col + 1) or 0
    if not best_dist or dist < best_dist then
      best, best_dist = m, dist
    end
  end
  return best
end

-- Accept or reject part [a, b] of a change (all of it for a formatting change). What is outside
-- that part stays a change. Returns whether the text changed.
local function apply(buf, m, action, a, b)
  local doc, ch = neo().docs[buf], neo().docs[buf].change[m[1]]
  local s, e = range(m)
  if ch.type == "format" then
    a, b = s, e
  end
  api.nvim_buf_del_extmark(buf, neo().ns.change, m[1])
  doc.change[m[1]] = nil
  local kept = false
  if not same(s, a) then
    neo().place(buf, "change", ch, s[1], s[2], a[1], a[2])
    kept = true
  end
  if not same(b, e) then -- (the rest of it, a change of its own)
    local rest = kept and { type = ch.type, author = ch.author, date = ch.date } or ch
    neo().place(buf, "change", rest, b[1], b[2], e[1], e[2])
  end
  if ch.type == "format" then
    local flags, props = ch[action] and ch[action]()
    if flags then -- (.docx only: .odt doesn't keep the formatting from before)
      neo().set_format(buf, s[1], s[2], e[2], flags, props)
      return true
    end
  elseif (action == "accept") == (ch.type == "delete") then -- an accepted deletion, a rejected insertion
    api.nvim_buf_set_text(buf, a[1], a[2], b[1], b[2], {})
    local lines = api.nvim_buf_get_lines(buf, 0, -1, false)
    if ch.type == "delete" and a[2] == 0 and lines[a[1] + 1] == "" and (b[1] > a[1] or b[2] > 0) then
      -- all of a paragraph was deleted: the paragraph goes too (with a blank line around it)
      local from, to = a[1], a[1] + 1
      if lines[a[1] + 2] == "" then
        to = to + 1
      elseif lines[a[1]] == "" then
        from = from - 1
      end
      api.nvim_buf_set_lines(buf, from, to, false, {})
    end
    return not same(a, b)
  elseif ch.type == "delete" then -- a rejected deletion: its text stays, with its paragraph breaks
    local shown = text(buf, a, b)
    if shown:find("¶") then
      api.nvim_buf_set_text(buf, a[1], a[2], b[1], b[2], vim.split(shown:gsub("¶", "\n"), "\n"))
      return true
    end
  end
  return false
end

-- action: "accept" or "reject"; all: every change (of `author`, if given); else the selected part
-- of the changes a visual selection touches (use_marks: the last selection, for
-- :'<,'>Neoffice accept), or the change at the cursor.
function M.resolve(action, all, use_marks, author)
  local buf = api.nvim_get_current_buf()
  if not neo().docs[buf] then
    return vim.notify("[neoffice] Not a document buffer", vim.log.levels.WARN)
  end
  local selection = not all and neo().selection(use_marks)
  local list = {}
  if all or selection then
    for _, m in ipairs(marks(buf, 0, -1)) do
      local s, e = range(m)
      if all then
        if not author or neo().docs[buf].change[m[1]].author == author then
          list[#list + 1] = { m, s, e }
        end
      elseif before(selection[1], e) and before(s, selection[2]) and (not same(selection[2], s) or same(s, e)) then
        list[#list + 1] = { m, before(selection[1], s) and s or selection[1], before(e, selection[2]) and e or selection[2] }
      end
    end
  else
    local m = at_cursor(buf)
    list = m and { { m, range(m) } } or {}
  end
  if #list == 0 then
    return vim.notify("[neoffice] No tracked change here")
  end
  vim.o.undolevels = vim.o.undolevels -- (each review is an undo step of its own)
  local changed = false
  for k = #list, 1, -1 do -- back to front, so edits don't move the changes still to do
    changed = apply(buf, list[k][1], action, list[k][2], list[k][3]) or changed
  end
  if not changed then -- (still a step of its own, so that u undoes it)
    api.nvim_buf_set_text(buf, 0, 0, 0, 0, { " " })
    api.nvim_buf_set_text(buf, 0, 0, 0, 1, {})
  end
  vim.bo[buf].modified = true
  M.snapshot(buf) -- (reviewing isn't recorded as a new change)
  M.after_edit(buf) -- (this state remembers its changes now)
end

function M.jump(dir)
  local row, col = unpack(api.nvim_win_get_cursor(0))
  local all = marks(0, 0, -1)
  for k = dir > 0 and 1 or #all, dir > 0 and #all or 1, dir do
    local m = all[k]
    if dir * (m[2] - (row - 1)) > 0 or (m[2] == row - 1 and dir * (m[3] - col) > 0) then
      return api.nvim_win_set_cursor(0, { m[2] + 1, m[3] })
    end
  end
  vim.notify("[neoffice] No more tracked changes")
end

function M.list(author) -- (all of them, or those of one author)
  local buf = api.nvim_get_current_buf()
  local doc = neo().docs[buf]
  if not doc then
    return
  end
  local items = {}
  for _, m in ipairs(marks(buf, 0, -1)) do
    local ch, s, e = doc.change[m[1]], range(m)
    items[#items + 1] = (not author or ch.author == author) and {
      bufnr = buf, lnum = s[1] + 1, col = s[2] + 1,
      text = ("[%s] %s, %s: %s"):format(ch.type, ch.author, ch.date:sub(1, 10), text(buf, s, e):gsub("\n", " ")),
    } or nil
  end
  vim.fn.setloclist(0, items)
  vim.cmd(#items > 0 and "lopen" or "echo 'No tracked changes'")
end

return M
