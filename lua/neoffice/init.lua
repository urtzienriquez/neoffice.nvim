-- neoffice: edit .docx/.odt files as plain text with Markdown-style formatting marks.
--
-- Paragraphs are laid out as in Markdown, separated by blank lines (see M.paragraphs).
-- Everything that isn't plain text is kept in extmarks, which Neovim moves along as you edit:
--   run      original run properties (font, colour, ...) of a stretch of text
--   object   an element we don't edit: an image (shown as ![alt](path)), field, bookmark, ...
--   comment  a comment thread anchored to a range
--   change   a tracked insertion (range) or deletion (point, deleted text shown inline)
-- On save, each paragraph is cut into pieces wherever the formatting or the marks change. Only
-- paragraphs whose pieces differ from when the document was opened (or last saved) are written
-- back, by the format module (docx.lua / odt.lua), and only the parts of the file that changed.
-- .doc files are converted to .docx with LibreOffice on open, and back on save.
local xml = require("neoffice.xml")
local zip = require("neoffice.zip")
local markup = require("neoffice.markup")
local api = vim.api
local M = {}

M.config = {
  author = vim.env.USER or "neoffice", -- name on your comments and tracked changes
  track_changes = nil, -- record your edits as tracked changes: true/false, nil = as the document says
  conceal = false, -- hide the formatting marks (except on the cursor line)
  images = true, -- show images in the buffer (needs Neovim's vim.ui.img and a terminal that supports it)
  image_width = 60, -- at most this many columns wide
  image_height = 20, -- and this many lines high
  panel_keys = { -- in the comments panel (false disables one)
    reply = "r", resolve = "<CR>", delete = "d", jump = "o", close = "q", next = "]]", prev = "[[",
  },
}

M.ns = {}
for _, kind in ipairs({ "run", "object", "comment", "change", "empty" }) do
  M.ns[kind] = api.nvim_create_namespace("neoffice_" .. kind)
end
local NS_DECOR = api.nvim_create_namespace("neoffice_decorations")

-- buf -> { path, base (neoffice's copy, see open()), is_doc, stamp, synced, fmt, ctx,
--          paras (nodes), texts, sigs, tracking,
--          run = {}, object = {}, comment = {}, change = {} (extmark id -> data) }
M.docs = {}

function M.setup(opts)
  M.config = vim.tbl_deep_extend("force", M.config, opts or {})
end

-- Create an extmark of `kind` holding `data`. Marks start with left gravity so a range can never
-- end before it starts; runs also grow at their end, so text typed after a word takes its format.
function M.place(buf, kind, data, r1, c1, r2, c2)
  local opts = { end_row = r2, end_col = c2, right_gravity = false, end_right_gravity = kind == "run" }
  if kind == "comment" or kind == "change" then
    opts = vim.tbl_extend("force", opts, require("neoffice." .. kind .. "s").style(data))
  end
  local id = api.nvim_buf_set_extmark(buf, M.ns[kind], r1, c1, opts)
  M.docs[buf][kind][id] = data
  return id
end

-- ── Paragraphs ──────────────────────────────────────────────────────────────

-- The paragraphs of the buffer, as in Markdown: lines separated by a blank line, a paragraph's
-- lines joined by spaces. A heading, title or list item starts a paragraph of its own (list items
-- can follow each other without blank lines). Further blank lines are empty paragraphs (as are
-- blank lines at the start or end). Each is { first, last, owns (last row whose marks belong to it), text }.
-- A table row line holds a paragraph for each paragraph of each cell, with { col, stop } (its text's
-- byte columns) and cell = { table, row, col } (the rows where its table and row start, its column).
-- alone: rows that are a paragraph of their own (an empty paragraph you typed in: see M.blocks).
function M.paragraphs(lines, alone)
  alone = alone or {}
  local out, blanks, table_start = {}, {}, nil
  local function empty_paragraphs(from)
    for k = from, #blanks do
      out[#out + 1] = { first = blanks[k], last = blanks[k], text = "" }
    end
    blanks = {}
  end
  for row, line in ipairs(lines) do
    local prev = out[#out]
    local cells = line:match("^|") and markup.cells(line)
    table_start = line:match("^|") and (#blanks == 0 and table_start or row) or nil
    if line:match("^%s*$") then
      blanks[#blanks + 1] = row
    elseif line:match("^|") then -- a table row (or the line of dashes under its first row)
      empty_paragraphs(prev and 2 or 1)
      for c, cell in ipairs(cells or {}) do
        for _, para in ipairs(cell) do
          out[#out + 1] = {
            first = row, last = row, owns = row, text = para.text, col = para.col, stop = para.col + #para.text,
            cell = { table = table_start, row = row, col = c },
          }
        end
      end
    elseif prev and prev.text ~= "" and not prev.cell and not prev.alone and not alone[row] and #blanks == 0 and not next((markup.block(line))) then -- continues it
      local hard = #(prev.text:match("\\*$")) % 2 == 1 -- (after a "\": a line break, else a space)
      prev.last, prev.text = row, prev.text .. (hard and "\n" or " ") .. line
    else
      empty_paragraphs(prev and 2 or 1) -- (the first blank line after a paragraph separates)
      out[#out + 1] = { first = row, last = row, text = line, alone = alone[row] }
    end
  end
  empty_paragraphs(1)
  local next_first = #lines + 1
  for k = #out, 1, -1 do
    out[k].owns = out[k].cell and out[k].first or next_first - 1
    next_first = out[k].first
  end
  return out
end

-- The paragraphs of a buffer. A line that was an empty paragraph (a blank line of its own) stays a
-- paragraph of its own when you type in it, even with no blank line before or after it, as in a
-- word processor (it has a mark at its start; deleting the line or joining it to another drops it).
function M.blocks(buf)
  local alone = {}
  for _, m in ipairs(api.nvim_buf_get_extmarks(buf, M.ns.empty, 0, -1, { details = true })) do
    if m[3] == 0 and not m[4].invalid then
      alone[m[2] + 1] = true
    end
  end
  return M.paragraphs(api.nvim_buf_get_lines(buf, 0, -1, false), alone)
end

local function covers(mark, row, col)
  local d = mark[4]
  return (mark[2] < row or mark[2] == row and mark[3] <= col) and (d.end_row > row or d.end_row == row and d.end_col > col)
end

local SAME = { "props", "link", "insert", unpack(markup.FLAGS) } -- (what makes characters one piece)

-- Cut a paragraph into pieces: text with uniform formatting, objects, images, and comment/change
-- markers. Reads the buffer and its marks only.
local function pieces(buf, para)
  local doc = M.docs[buf]
  local block, chars = markup.parse(para.text)
  block.cell = para.cell and true or nil
  -- where each character is in the buffer (the paragraph's lines were joined with spaces)
  local lines = api.nvim_buf_get_lines(buf, para.first - 1, para.last, true)
  local row, start = 1, -(para.col or 0)
  for _, ch in ipairs(chars) do
    while row < #lines and ch.col > start + #lines[row] do
      start, row = start + #lines[row] + 1, row + 1
    end
    ch.row, ch.col = para.first - 2 + row, ch.col - start
  end
  local function inside(r, c) -- (a table cell: only what is in its columns)
    return not para.cell or (c >= para.col and c <= para.stop)
  end
  local function marks(kind)
    return api.nvim_buf_get_extmarks(buf, M.ns[kind], { para.first - 1, 0 }, { para.owns - 1, -1 }, { details = true, overlap = true })
  end
  local runs, objects = marks("run"), marks("object")
  local changes = vim.tbl_map(function(m)
    return vim.list_extend(m, { doc.change[m[1]] })
  end, marks("change"))

  -- markers, sorted by position; at the same spot, ranges end (1) before others (2) and starts (3)
  local events = {}
  local function event(r, c, rank, piece)
    if r >= para.first - 1 and r <= para.owns - 1 and inside(r, c) then
      local at = #chars + 1
      for k, ch in ipairs(chars) do
        if ch.row > r or ch.row == r and ch.col >= c then
          at = k
          break
        end
      end
      events[#events + 1] = { at, rank, #events, piece }
    end
  end
  for _, m in ipairs(marks("comment")) do
    local c, d = doc.comment[m[1]], m[4]
    local point = m[2] == d.end_row and m[3] == d.end_col
    event(m[2], m[3], 3, { comment_start = c, point = point })
    if not point then
      event(d.end_row, d.end_col, 1, { comment_end = c })
    end
  end
  -- a deletion is the text it covers (shown struck through; ¶ for a paragraph break)
  -- (images, equations and other objects in deleted text stay in the document until it is accepted)
  local deletions = vim.tbl_filter(function(m) return m[5].type == "delete" end, changes)
  local function deleted_by(ch)
    return not (ch.image or ch.math or ch.footnote or ch.field or ch.ch == "￼") and vim.iter(deletions):find(function(m) return covers(m, ch.row, ch.col) end)
  end
  for _, m in ipairs(changes) do
    if m[5].type == "delete" then
      -- (written with the text, below)
    else
      event(m[2], m[3], 3, { insert_start = m[5] })
      event(m[4].end_row, m[4].end_col, 1, { insert_end = m[5] })
    end
  end
  for _, m in ipairs(objects) do
    if not doc.object[m[1]].visible then
      event(m[2], m[3], 2, { object = doc.object[m[1]].node })
    end
  end
  table.sort(events, function(x, y)
    return x[1] ~= y[1] and x[1] < y[1] or x[1] == y[1] and (x[2] ~= y[2] and x[2] < y[2] or x[2] == y[2] and x[3] < y[3])
  end)

  local out, e, parts = {}, 1, {} -- (parts: how many pieces of each deletion so far)
  local props = runs[1] and doc.run[runs[1][1]] -- text before the first run takes its properties
  for k = 1, #chars + 1 do
    while events[e] and events[e][1] == k do
      out[#out + 1], e = events[e][4], e + 1
    end
    local ch = chars[k]
    if not ch then
      break
    end
    local deletion = deleted_by(ch)
    if deletion then -- deleted text: the deletion's part here (continuing, unless something came between)
      local last = out[#out]
      if last and last.deletion == deletion[5] then
        last.deleted = last.deleted .. (ch.ch == "¶" and "\n" or ch.ch)
      else
        parts[deletion[5]] = (parts[deletion[5]] or (deletion[2] < para.first - 1 and 1 or 0)) + 1
        out[#out + 1] = { deletion = deletion[5], deleted = ch.ch == "¶" and "\n" or ch.ch, part = parts[deletion[5]] - 1 }
      end
      goto next_char
    end
    local mark = (ch.ch == "￼" or ch.image or ch.math or ch.footnote or ch.field) and vim.iter(objects):find(function(m)
      return doc.object[m[1]].visible and m[2] == ch.row and m[3] == ch.col
    end)
    local object = mark and doc.object[mark[1]]
    if ch.image then -- unchanged, edited (another path or alt text) or new
      local same = object and object.image and object.image.path == ch.image.path and object.image.alt == ch.image.alt
        and object.image.width == ch.image.width
      assert(not ch.image.width or markup.cm(ch.image.width, 1), "image width: write it in cm, mm, in or % (of the text's width), as in {width=12cm} or {width=50%}")
      local where = { mark = mark and mark[1], row = ch.row, col = ch.col, end_col = ch.col + ch.len }
      out[#out + 1] = same and { object = object.node } or { image = vim.tbl_extend("force", ch.image, where, { original = object }) }
    elseif ch.math then -- an equation: unchanged, or (re)written from its LaTeX
      local same = object and object.math and object.math.tex == ch.math.tex and object.math.display == ch.math.display
      local where = { mark = mark and mark[1], row = ch.row, col = ch.col, end_col = ch.col + ch.len }
      out[#out + 1] = same and { object = object.node } or { math = vim.tbl_extend("force", ch.math, where) }
    elseif ch.footnote then -- a footnote: unchanged, edited or new
      local same = object and object.footnote and object.footnote.text == ch.footnote.text
      local where = { mark = mark and mark[1], row = ch.row, col = ch.col, end_col = ch.col + ch.len }
      out[#out + 1] = same and { object = object.node } or { footnote = vim.tbl_extend("force", ch.footnote, where, { original = object }) }
    elseif ch.field then -- a field (citation, date...): kept as it is; its text can't be edited
      if object and object.field then
        out[#out + 1] = { object = object.node, kept = object.field.text ~= ch.field.text and ch.field.text or nil }
      else -- (typed: just text)
        out[#out + 1] = { text = "⟦" .. ch.field.text .. "⟧", props = props or nil }
      end
    elseif ch.page_break then
      out[#out + 1] = { page_break = true }
    elseif ch.ch == "\n" then
      out[#out + 1] = { line_break = true }
    elseif object then
      out[#out + 1] = { object = object.node }
    else
      -- the first run covering this character (text typed between two runs belongs to the first);
      -- text outside any run keeps the properties of the character before it
      local run = vim.iter(runs):find(function(m) return covers(m, ch.row, ch.col) end)
      if run then
        props = doc.run[run[1]]
      end
      local piece = { text = ch.ch, props = props or nil, link = ch.link }
      for _, flag in ipairs(markup.FLAGS) do
        piece[flag] = ch[flag]
      end
      for _, m in ipairs(changes) do
        if m[5].type == "insert" and covers(m, ch.row, ch.col) then
          piece.insert = m[5]
        end
      end
      local last = out[#out]
      if last and last.text and vim.iter(SAME):all(function(key) return last[key] == piece[key] end) then
        last.text = last.text .. ch.ch
      else
        out[#out + 1] = piece
      end
    end
    ::next_char::
  end
  -- an insertion (deletion) running past the end of the paragraph inserted (deleted) its break too
  local stop = { para.last - 1, para.stop or #lines[#lines] }
  for _, m in ipairs(changes) do
    local d = m[4]
    local before_end = m[2] < stop[1] or m[2] == stop[1] and m[3] <= stop[2]
    if m[5].type ~= "format" and before_end and (d.end_row > stop[1] or d.end_row == stop[1] and d.end_col > stop[2]) then
      out[#out + 1] = { paragraph_mark = m[5] }
    end
  end
  return block, out
end

-- A fingerprint of a paragraph's pieces: equal fingerprints mean the paragraph's XML is still right.
local function signature(block, list)
  local out = { vim.inspect(block, { newline = "", indent = "" }) }
  for _, piece in ipairs(list) do
    local keys = vim.tbl_keys(piece)
    table.sort(keys)
    for _, key in ipairs(keys) do
      local v = piece[key]
      if type(v) == "table" then -- (identity, plus what can change inside comments, changes, images)
        local replies = table.concat(vim.tbl_map(tostring, v.replies or {}), ",")
        v = table.concat({ tostring(v), replies, tostring(v.resolved), v.text or "", v.path or "", v.alt or "", v.width or "", v.tex or "", tostring(v.display) }, "/")
      end
      out[#out + 1] = key .. "=" .. tostring(v)
    end
    out[#out + 1] = ";"
  end
  return table.concat(out, " ")
end

-- ── Open ────────────────────────────────────────────────────────────────────

local function setup_buffer(buf)
  vim.bo[buf].filetype = "neoffice"
  pcall(vim.treesitter.start, buf) -- markdown (see plugin/neoffice.lua): highlights the marks
  vim.bo[buf].syntax = "neoffice" -- adds <u> (see syntax/neoffice.vim)
  local function window()
    vim.wo.wrap, vim.wo.linebreak, vim.wo.breakindent = true, true, true
    if M.config.conceal then
      vim.wo.conceallevel = 2
    end
  end
  local group = api.nvim_create_augroup("neoffice_" .. buf, {})
  api.nvim_create_autocmd("BufWinEnter", { group = group, buffer = buf, callback = window })
  api.nvim_create_autocmd("BufWipeout", { group = group, buffer = buf, callback = function() M.docs[buf] = nil end })
  api.nvim_create_autocmd({ "TextChanged", "TextChangedI" }, {
    group = group,
    buffer = buf,
    callback = function()
      require("neoffice.changes").after_edit(buf)
      M.decorate(buf)
    end,
  })
  api.nvim_create_autocmd({ "CursorMoved", "CursorMovedI", "TextChanged" }, {
    group = group,
    buffer = buf,
    callback = function() require("neoffice.changes").show_info(buf) end,
  })
  if api.nvim_get_current_buf() == buf then
    window()
  end
  require("neoffice.images").attach(buf, group)
  M.decorate(buf)
end

-- Convert a file with LibreOffice (to = "docx", "doc:MS Word 97", ...) into `dir`; returns the new path.
local function convert(path, to, dir)
  local profile = "-env:UserInstallation=file://" .. vim.fn.stdpath("cache") .. "/neoffice-libreoffice"
  local r = vim.system({ "soffice", profile, "--headless", "--convert-to", to, "--outdir", dir, path }):wait()
  local out = dir .. "/" .. vim.fn.fnamemodify(path, ":t:r") .. "." .. to:match("^%w+")
  assert(r.code == 0 and vim.uv.fs_stat(out), "LibreOffice (soffice) could not convert " .. path)
  return out
end

-- A copy of the plugin's empty document (templates/), to start a new one from.
local function template(ext)
  local here = vim.fs.dirname(debug.getinfo(1, "S").source:sub(2)) -- (lua/neoffice)
  local copy = vim.fn.tempname() .. "." .. ext
  assert(vim.uv.fs_copyfile(vim.fs.joinpath(here, "..", "..", "templates", "empty." .. ext), copy))
  return copy
end

-- The table cell a paragraph is in: { tc, tr, tbl } (cell, row and table elements), or nil.
local CELL = { ["w:tc"] = true, ["table:table-cell"] = true }
local ROW = { ["w:tr"] = true, ["table:table-row"] = true }
local TABLE = { ["w:tbl"] = true, ["table:table"] = true }
local function up(node, tags)
  repeat
    node = node.parent
  until not node or tags[node.tag]
  return node
end
function M.cell_of(node)
  local tc = up(node, CELL)
  local tr = tc and up(tc, ROW)
  return tr and { tc = tc, tr = tr, tbl = up(tr, TABLE) } or nil
end

-- Size and modification time of a file (nil if it doesn't exist), to notice other programs' writes.
local function stamp(path)
  local s = vim.uv.fs_stat(path)
  return s and ("%d:%d.%d"):format(s.size, s.mtime.sec, s.mtime.nsec)
end

-- neoffice works on its own copy of the document (the base: a .docx for .doc files, the template
-- for a new one), and copies it to the file when writing. So saving never depends on the file
-- staying where it was, and a file deleted meanwhile is simply written again.
local function open(buf, path)
  local ext = path:match("%.(%w+)$"):lower()
  local new = not vim.uv.fs_stat(path) -- a new document: created on the first :w
  local base
  if ext == "doc" then
    base = new and template("docx") or convert(path, "docx", vim.fn.tempname())
  elseif new then
    base = template(ext)
  else
    base = vim.fn.tempname() .. "." .. ext
    assert(vim.uv.fs_copyfile(path, base))
  end
  local fmt = require("neoffice." .. (ext == "doc" and "docx" or ext))
  local ctx, paras = fmt.load(base)
  ctx.dirty, ctx.files = {}, {} -- parts to write on save (name -> true, or content for new files)
  local doc = { path = path, base = base, is_doc = ext == "doc", fmt = fmt, ctx = ctx, paras = {}, run = {}, object = {}, comment = {}, change = {} }
  doc.stamp, doc.synced = stamp(path), not new -- (synced: the file has the base's content)
  doc.tracking = M.config.track_changes
  if doc.tracking == nil then
    doc.tracking = ctx.tracking or false
  end
  M.docs[buf] = doc
  for _, ns in pairs(M.ns) do
    api.nvim_buf_clear_namespace(buf, ns, 0, -1)
  end

  -- numbered list items show their numbers (counted per list and level)
  local counters = {}
  for _, para in ipairs(paras) do
    local list = para.block.list
    if list then
      local count = counters[list.id or ""] or {}
      counters[list.id or ""] = count
      count[list.level] = (count[list.level] or 0) + 1
      for deeper = list.level + 1, #count do
        count[deeper] = nil
      end
      list.n = count[list.level]
    end
  end

  local texts, cols, last_text, cells = {}, {}, 0, {}
  for k, para in ipairs(paras) do
    cells[k] = M.cell_of(para.node)
    para.block.cell = cells[k] and true or nil
    for _, s in ipairs(para.spans) do
      if cells[k] and (s.line_break or s.page_break) then -- (in a table cell: kept, as ￼)
        s.line_break, s.page_break, s.visible = nil, nil, true
      elseif s.line_break or s.page_break then
        s.object = nil
      end
      if s.deletion then -- (shown in the text, struck through; a paragraph break as ¶¶)
        s.text = s.deletion.text:gsub("\n", "¶¶")
      elseif s.text then
        s.text = s.text:gsub("\n", " ")
      elseif s.object and s.visible and not (s.image or s.math or s.footnote or s.field) then
        s.text, s.raw = "￼", true
      end
    end
    texts[k], cols[k] = markup.render(para.spans, para.block)
    last_text = (texts[k] ~= "" or cells[k]) and k or last_text
    doc.paras[k] = para.node
  end
  doc.texts = texts

  -- Lay paragraphs out as in Markdown (see M.paragraphs): a blank line after each paragraph that
  -- has text and is followed by more text (but not between list items); an empty paragraph is a
  -- blank line of its own.
  -- A table is a pipe table: a line per row, "|" between cells and "<br>" between the paragraphs
  -- of a cell, and a line of dashes after its first row.
  local lines, marks, started, rows, columns, empty_rows = {}, {}, {}, {}, 0, {}
  for k, para in ipairs(paras) do
    local r, at
    local cell, before, after = cells[k], cells[k - 1], cells[k + 1]
    if cell then
      if before and before.tr == cell.tr then
        lines[#lines] = lines[#lines] .. (before.tc == cell.tc and " <br> " or " | ")
        columns = columns + (before.tc == cell.tc and 0 or 1)
      else
        lines[#lines + 1], columns = "| ", 1
        rows[cell.tbl] = (rows[cell.tbl] or 0) + 1
      end
      r = #lines - 1
      local offset = #lines[#lines]
      lines[#lines] = lines[#lines] .. texts[k]
      at = function(col) return r, offset + col end
      if not (after and after.tr == cell.tr) then -- the row ends
        lines[#lines] = lines[#lines] .. " |"
        if rows[cell.tbl] == 1 then
          lines[#lines + 1] = "|" .. ("---|"):rep(columns)
        end
        if not (after and after.tbl == cell.tbl) and k < last_text then
          lines[#lines + 1] = ""
        end
      end
    else
      r = #lines -- (a paragraph with line breaks takes several lines)
      local starts, offset = {}, 0
      for _, line in ipairs(vim.split(texts[k], "\n", { plain = true })) do
        lines[#lines + 1], starts[#starts + 1], offset = line, offset, offset + #line + 1
      end
      at = function(col) -- the position of a byte column of the paragraph's text
        local row = #starts
        while starts[row] > col do
          row = row - 1
        end
        return r + row - 1, col - starts[row]
      end
      local next_list = paras[k + 1] and paras[k + 1].block.list
      local tight = para.block.list and next_list and next_list.id == para.block.list.id and texts[k + 1] ~= ""
      if texts[k] ~= "" and k < last_text and not tight then
        lines[#lines + 1] = ""
      end
      if texts[k] == "" then
        empty_rows[#empty_rows + 1] = r
      end
    end
    local seen_text = false
    for s_index, s in ipairs(para.spans) do
      local ra, ca = at(cols[k][s_index][1])
      local rb, cb = at(cols[k][s_index][2])
      local range_start = s.comment_start or s.insert_start
      local range_end = s.comment_end or s.insert_end
      if range_end and not seen_text and started[range_end] and started[range_end][1] < r then
        ra, ca = r, 0 -- (a range from an earlier paragraph ends where this one's line starts, before any "- ")
      end
      seen_text = seen_text or (s.text and s.text ~= "")
      local data = s.object and { node = s.object, visible = s.visible, image = s.image, math = s.math, footnote = s.footnote, field = s.field }
      if s.deletion then
        marks[#marks + 1] = { "change", s.deletion, ra, ca, rb, cb }
      elseif s.object then
        marks[#marks + 1] = { "object", data, ra, ca, rb, cb }
      elseif s.line_break or s.page_break then
        -- (written in the text: nothing to remember)
      elseif s.text then
        marks[#marks + 1] = { "run", s.props or false, ra, ca, rb, cb }
      elseif range_start then
        started[range_start] = { ra, ca }
      elseif range_end then
        local from = started[range_end] or { ra, ca }
        started[range_end] = nil
        marks[#marks + 1] = { s.comment_end and "comment" or "change", range_end, from[1], from[2], ra, ca }
      end
    end
  end
  for data, from in pairs(started) do -- ranges that never ended become points
    marks[#marks + 1] = { data.replies and "comment" or "change", data, from[1], from[2], from[1], from[2] }
  end

  local undolevels = vim.bo[buf].undolevels
  vim.bo[buf].undolevels = -1
  api.nvim_buf_set_lines(buf, 0, -1, false, lines)
  vim.bo[buf].undolevels = undolevels
  for _, mark in ipairs(marks) do
    M.place(buf, unpack(mark))
  end
  local alone = {}
  for _, r in ipairs(empty_rows) do
    api.nvim_buf_set_extmark(buf, M.ns.empty, r, 0, { invalidate = true, right_gravity = false })
    alone[r + 1] = true
  end
  doc.sigs = {}
  for k, block in ipairs(M.paragraphs(lines, alone)) do
    assert(block.text == texts[k], "neoffice: paragraph layout mismatch") -- (see markup.render)
    doc.sigs[k] = signature(pieces(buf, block))
  end
  vim.bo[buf].modified = false
  setup_buffer(buf)
  require("neoffice.changes").start(buf)
  if new then
    vim.notify("[neoffice] New document: it is created when you write it")
  end
end

function M.open(buf, path)
  path = vim.fn.fnamemodify(path, ":p")
  local ok, err = pcall(open, buf, path)
  if not ok then -- an empty or half-read buffer must never be written over the document
    M.docs[buf] = nil
    for _, ns in pairs(M.ns) do
      api.nvim_buf_clear_namespace(buf, ns, 0, -1)
    end
    api.nvim_buf_set_lines(buf, 0, -1, false, {})
    vim.bo[buf].modified, vim.bo[buf].modifiable, vim.bo[buf].readonly = false, false, true
    vim.notify("[neoffice] Could not read " .. path .. ": " .. err, vim.log.levels.ERROR)
  end
end

-- The visual selection as { start, end } positions (end excluded), leaving visual mode; nil in
-- normal mode. use_marks: the last selection instead (for commands like :'<,'>Neoffice accept).
function M.selection(use_marks)
  local mode = use_marks and vim.fn.visualmode() or vim.fn.mode()
  if not mode:match("^[vV\22]") then
    return nil
  end
  local s, e = vim.fn.getpos(use_marks and "'<" or "v"), vim.fn.getpos(use_marks and "'>" or ".")
  if s[2] > e[2] or (s[2] == e[2] and s[3] > e[3]) then
    s, e = e, s
  end
  local last = vim.fn.getline(e[2])
  local stop = { e[2] - 1, math.min(e[3] - 1 + #vim.fn.matchstr(last, ".", e[3] - 1), #last) }
  local first_col = s[3] - 1
  if mode == "V" then -- (whole lines, with their line break)
    first_col = 0
    stop = e[2] < vim.fn.line("$") and { e[2], 0 } or { e[2] - 1, #last }
  end
  if not use_marks then
    api.nvim_feedkeys(vim.keycode("<Esc>"), "nx", false)
  end
  return { { s[2] - 1, first_col }, stop }
end

-- Decorations that follow the text: titles highlighted (NeofficeTitle, above tree-sitter) and,
-- when the document numbers its headings, their numbers (1, 1.1, ...) shown before them.
function M.decorate(buf)
  local doc = M.docs[buf]
  if not doc then
    return
  end
  api.nvim_buf_clear_namespace(buf, NS_DECOR, 0, -1)
  local lines, count = api.nvim_buf_get_lines(buf, 0, -1, false), {}
  for _, m in ipairs(api.nvim_buf_get_extmarks(buf, M.ns.change, 0, -1, { details = true })) do
    local ch = doc.change[m[1]]
    for row = m[2], m[4].end_row - 1 do -- (a deleted line break: shown as a struck ¶ at the end of its line)
      if ch and ch.type == "delete" then
        local hl = require("neoffice.changes").style(ch).hl_group
        api.nvim_buf_set_extmark(buf, NS_DECOR, row, #lines[row + 1], { virt_text = { { "¶", hl } }, virt_text_pos = "inline" })
      end
    end
  end
  for _, para in ipairs(M.blocks(buf)) do
    local block, text_start = markup.block(lines[para.first])
    -- the alignment mark (above Markdown's highlighting, which takes ">..." for a quote)
    local own, own_start, _, align_at = markup.block(para.text)
    if own.align then
      local col = (para.col or 0) + align_at - 1
      api.nvim_buf_set_extmark(buf, NS_DECOR, para.first - 1, col, {
        end_col = (para.col or 0) + own_start - 2, hl_group = "NeofficeAlign", priority = 250,
      })
    end
    if block.title then
      api.nvim_buf_set_extmark(buf, NS_DECOR, para.first - 1, 0, {
        end_row = para.last - 1, end_col = #lines[para.last], hl_group = "NeofficeTitle", priority = 200,
      })
    elseif block.heading and doc.ctx.numbered then
      count[block.heading] = (count[block.heading] or 0) + 1
      for deeper = block.heading + 1, 9 do
        count[deeper] = nil
      end
      local number = {}
      for level = 1, block.heading do
        number[level] = count[level] or 0
      end
      api.nvim_buf_set_extmark(buf, NS_DECOR, para.first - 1, text_start - 1, {
        virt_text = { { table.concat(number, ".") .. " ", "NeofficeNumber" } }, virt_text_pos = "inline",
      })
    end
  end
end

-- :Neoffice numbers [on|off]: whether the document numbers its headings (a change to the document).
function M.numbers(arg)
  local buf = api.nvim_get_current_buf()
  local doc = M.docs[buf]
  if not doc then
    return vim.notify("[neoffice] Not a document buffer", vim.log.levels.WARN)
  end
  local on = arg == "on" or (arg ~= "off" and not doc.ctx.numbered)
  if on ~= doc.ctx.numbered then
    doc.fmt.number_headings(doc.ctx, on)
    vim.bo[buf].modified = true
  end
  M.decorate(buf)
  vim.notify("[neoffice] Heading numbers " .. (on and "on" or "off"))
end

-- Give columns [c1, c2) of a line this bold/italic/underline (rewriting its marks) and these
-- run properties. Used to reject a tracked formatting change.
function M.set_format(buf, row, c1, c2, flags, props)
  local line = api.nvim_buf_get_lines(buf, row, row + 1, true)[1]
  local block, chars = markup.parse(line)
  local spans = vim.tbl_map(function(c)
    local inside = c.col >= c1 and c.col < c2
    local span = { text = c.ch, link = c.link, image = c.image, math = c.math }
    for _, flag in ipairs(markup.FLAGS) do
      span[flag] = c[flag]
      if inside and flags[flag] ~= nil then
        span[flag] = flags[flag]
      end
    end
    return span
  end, chars)
  local new = markup.render(spans, block)
  -- replace only what differs, so the marks around it stay where they are
  local p, s = 0, 0
  while p < math.min(#line, #new) and line:byte(p + 1) == new:byte(p + 1) do
    p = p + 1
  end
  while s < math.min(#line, #new) - p and line:byte(#line - s) == new:byte(#new - s) do
    s = s + 1
  end
  api.nvim_buf_set_text(buf, row, p, row, #line - s, { new:sub(p + 1, #new - s) })
  M.place(buf, "run", props, row, p, row, #new - s)
end

-- ── Save ────────────────────────────────────────────────────────────────────

-- Line diff hunks { old start, old count, new start, new count }, starts adjusted for empty sides.
function M.diff(old, new)
  local text = function(list) -- (one line per entry: a paragraph's line breaks don't count)
    list = vim.tbl_map(function(t) return (t:gsub("\n", "\r")) end, list)
    return #list == 0 and "" or table.concat(list, "\n") .. "\n"
  end
  return vim.tbl_map(function(h)
    return { h[1] + (h[2] == 0 and 1 or 0), h[2], h[3] + (h[4] == 0 and 1 or 0), h[4] }
  end, (vim.text.diff or vim.diff)(text(old), text(new), { result_type = "indices" }))
end

-- Map each current paragraph to the one it continues (or nil for new ones) with a diff of their texts.
local function match_lines(old, new)
  local old_of, deleted, a, b = {}, {}, 1, 1
  local hunks = M.diff(old, new)
  table.insert(hunks, { #old + 1, 0, #new + 1, 0 }) -- so lines after the last hunk get mapped too
  for _, h in ipairs(hunks) do
    local sa, ca, sb, cb = unpack(h)
    while b < sb do -- unchanged lines
      old_of[b], a, b = a, a + 1, b + 1
    end
    for k = 0, math.max(ca, cb) - 1 do -- changed lines pair up in order
      if k < cb then
        old_of[sb + k] = k < ca and sa + k or nil
      else
        deleted[#deleted + 1] = sa + k
      end
    end
    a, b = sa + ca, sb + cb
  end
  return old_of, deleted
end

local PARAGRAPHS = { ["w:p"] = true, ["text:p"] = true, ["text:h"] = true }
local CONTENT = vim.tbl_extend("force", PARAGRAPHS, { ["w:tbl"] = true, ["text:list"] = true, ["text:list-item"] = true, ["table:table"] = true })

local function empty(el)
  return not vim.iter(el.children):any(function(c) return CONTENT[c.tag] end)
end

-- Remove deleted paragraphs. An empty list item or list goes too; a table cell keeps an empty
-- paragraph (it needs one), unless every paragraph of its table was deleted: then the table goes.
local function remove_paragraphs(fmt, nodes)
  local cells = {}
  for _, node in ipairs(nodes) do
    local parent = node.parent
    if parent then -- (unless it went with its list or table already)
      xml.remove(node)
      while parent.parent and (parent.tag == "text:list-item" or parent.tag == "text:list") and empty(parent) do
        local up = parent.parent
        xml.remove(parent)
        parent = up
      end
      if parent.tag == "w:tc" or parent.tag == "table:table-cell" then
        cells[parent] = node
      end
    end
  end
  for cell, template in pairs(cells) do
    local tbl = cell.parent
    while tbl and tbl.tag ~= "w:tbl" and tbl.tag ~= "table:table" do
      tbl = tbl.parent
    end
    if tbl and tbl.parent and #xml.collect(tbl, PARAGRAPHS) == 0 then
      xml.remove(tbl)
    elseif empty(cell) and (not tbl or tbl.parent) then
      xml.append(cell, fmt.new_para(template))
    end
  end
end

-- The bytes of an image: a file inside the document (as shown when it was opened) or on disk
-- (~ or relative to the document's folder).
local function read_image(doc, path)
  local data
  if doc.ctx.entries[path] then
    data = zip.read(doc.base, path)
  else
    local file = vim.fs.normalize(path)
    if not file:match("^/") then
      file = vim.fs.joinpath(vim.fs.dirname(doc.path), file)
    end
    local f = io.open(file, "rb")
    data = f and f:read("*a")
    if f then
      f:close()
    end
  end
  assert(data, "image not found: " .. path)
  assert(require("neoffice.images").info(data), "not a PNG, JPEG or GIF image: " .. path)
  return data
end

function M.save(buf, target)
  local doc = M.docs[buf]
  if not doc then
    return vim.notify("[neoffice] This document wasn't opened correctly: not saving it", vim.log.levels.ERROR)
  end
  target = vim.fn.fnamemodify(target or doc.path, ":p")
  local ok, err = pcall(function()
    local now = stamp(doc.path)
    if target == doc.path and now and doc.stamp and now ~= doc.stamp and vim.v.cmdbang ~= 1 then
      error("the file was changed by another program since it was read or written (:w! writes over it)")
    end
    local blocks = M.blocks(buf)
    if #doc.paras == 0 and #blocks == 1 and blocks[1].text == "" then -- (still no paragraph)
      blocks = {}
    end
    local texts = vim.tbl_map(function(p) return p.text end, blocks)
    local old_of, deleted = match_lines(doc.texts, texts)

    -- 1. Work out every paragraph first: if anything is wrong (a missing image, an equation that
    -- can't be converted), nothing changes.
    local plans, equations = {}, {}
    for k, block in ipairs(blocks) do
      local kind, list = pieces(buf, block)
      for _, piece in ipairs(list) do
        if piece.image then
          piece.image.bytes = read_image(doc, piece.image.path)
        end
      end
      for _, piece in ipairs(list) do
        if piece.math then
          equations[#equations + 1] = piece.math
        end
      end
      plans[k] = { old = old_of[k], block = kind, pieces = list, sig = signature(kind, list), cell = block.cell }
    end

    if #equations > 0 then
      local nodes = doc.fmt.equations and doc.fmt.equations(doc.ctx, equations)
      assert(nodes, "equations ($...$) can only be written in .docx documents (write \\$ for a dollar sign)")
      for k, eq in ipairs(equations) do
        eq.node = nodes[k]
      end
    end

    -- 2. Rewrite only the paragraphs that changed.
    local ctx, paras, prev, plain, kept = doc.ctx, {}, nil, nil, {}
    local layout = require("neoffice.tables").start(doc.fmt, ctx, vim.tbl_map(function(plan)
      return { cell = plan.cell, node = plan.old and doc.paras[plan.old] }
    end, plans), M.cell_of)
    local function outside(node) -- (the outermost table it is in, or itself)
      local cell = M.cell_of(node)
      return cell and outside(cell.tbl) or node
    end
    for k, plan in ipairs(plans) do
      local node = plan.old and doc.paras[plan.old]
      local placed = layout:node({ cell = plan.cell, node = node }, prev)
      if node and (placed and placed ~= node or not placed and M.cell_of(node)) then
        deleted[#deleted + 1], plan.old, node = plan.old, nil, nil -- (moved into or out of a table cell)
      end
      node = placed or node
      local anchor = prev and outside(prev) or doc.paras[1] and outside(doc.paras[1])
      if not node and not anchor then -- the document had no paragraph at all
        node = doc.fmt.first_para(ctx)
      elseif not node then -- a new paragraph, formatted like the one before it
        node = doc.fmt.new_para(plain or prev or doc.paras[1])
        xml.append(anchor.parent, node, xml.index(anchor) + (prev and 1 or 0))
      end
      plain = not plan.cell and node or plain
      if not plan.old or plan.sig ~= doc.sigs[plan.old] then
        doc.fmt.fill(ctx, node, plan.block, plan.pieces)
        ctx.dirty[doc.fmt.MAIN] = true
        for _, piece in ipairs(plan.pieces) do -- images written now: they are the document's from now on
          local new = piece.image or piece.math or piece.footnote
          if piece.kept then
            kept[#kept + 1] = piece.kept
          end
          if new then
            if new.mark then
              api.nvim_buf_del_extmark(buf, M.ns.object, new.mark)
              doc.object[new.mark] = nil
            end
            local data = { node = new.node, visible = true }
            data.image = piece.image and { path = new.path, alt = new.alt, width = new.width }
            data.math = piece.math and { tex = new.tex, display = new.display }
            data.footnote = piece.footnote and { text = new.text, id = new.id }
            M.place(buf, "object", data, new.row, new.col, new.row, new.end_col)
          end
        end
      end
      paras[k], prev = node, node
    end
    layout:finish()
    if #deleted > 0 then
      remove_paragraphs(doc.fmt, vim.tbl_map(function(old) return doc.paras[old] end, deleted))
      ctx.dirty[doc.fmt.MAIN] = true
    end
    doc.paras, doc.texts, doc.sigs = paras, texts, {}
    for k, block in ipairs(blocks) do
      doc.sigs[k] = signature(pieces(buf, block))
    end

    -- 3. Write the parts that changed (nothing at all if nothing did).
    local live = { comments = vim.tbl_values(doc.comment), changes = {} }
    for _, ch in pairs(doc.change) do
      live.changes[ch] = true
    end
    local entries = doc.fmt.save(ctx, live)
    if next(entries) then -- the base gets the changes...
      zip.write(doc.base, entries)
      doc.synced = false
    end
    ctx.dirty, ctx.files, ctx.comments_dirty = {}, {}, false
    if target ~= doc.path or not doc.synced or not vim.uv.fs_stat(target) then -- ...and goes to the file
      local file = doc.is_doc and convert(doc.base, "doc:MS Word 97", vim.fn.tempname()) or doc.base
      assert(vim.uv.fs_copyfile(file, target .. ".neoffice~"))
      assert(os.rename(target .. ".neoffice~", target))
      if target == doc.path then
        doc.synced, doc.stamp = true, stamp(target)
      end
    end
    if #kept > 0 then
      vim.schedule(function()
        vim.notify("[neoffice] Fields can't be edited (Word or Zotero would overwrite it); kept as they were: "
          .. table.concat(kept, ", "), vim.log.levels.WARN)
      end)
    end
  end)
  if not ok then
    return vim.notify("[neoffice] Save failed, the file was not changed: " .. err, vim.log.levels.ERROR)
  end
  if target == doc.path then
    vim.bo[buf].modified = false
  end
  vim.notify("[neoffice] Saved " .. vim.fn.fnamemodify(target, ":~:."))
end

return M
