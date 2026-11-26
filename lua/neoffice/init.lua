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
local floats = require("neoffice.floats")
local api = vim.api
local M = {}

M.config = {
  author = vim.env.USER or "neoffice", -- name on your comments and tracked changes
  track_changes = nil, -- record your edits as tracked changes: true/false, nil = as the document says
  conceal = false, -- hide the formatting marks (except on the cursor line)
  images = true, -- show images in the buffer (needs Neovim's vim.ui.img and a terminal that supports it)
  lsp = true, -- a language server for documents: headings as symbols (gO), folding, hover on comments
  comment_text = false, -- each comment's author and text at the end of its line (:Neoffice comments text)
  image_width = 60, -- at most this many columns wide
  image_height = 20, -- and this many lines high
  -- captions ("Table: ...", "Figure: ..."): their words, and what goes between the number and the text
  captions = { table = "Table", figure = { "Figure", "Fig." }, separator = ". " },
  -- how tables look, unless the line above one says otherwise ("{width = full}"; :help neoffice-table-format)
  tables = { rules = "booktabs", width = "auto", height = "compact", position = "center" },
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
-- byte columns) and cell = { table, row, col } (the rows where its table and row start, its column),
-- with its column's alignment, from the table's line of dashes (align: see markup.table_format),
-- and the options on the line above the table (format: "width = full, height = tall", see tables.lua), if it has one.
-- The paragraphs of a box (between "::: {width = 6cm, float = right}" and ":::": see floats.lua) have
-- box = { start (the row of its first line), attrs (the text in its braces) }.
-- alone: rows that are a paragraph of their own (an empty paragraph you typed in: see M.blocks).
function M.paragraphs(lines, alone)
  alone = alone or {}
  local out, blanks, table_start, formats, options = {}, {}, nil, {}, {}
  local box, box_of_row, cut = nil, {}, false -- (cut: a box's line came just before: no paragraph goes on across it)
  local function empty_paragraphs(from)
    for k = from, #blanks do
      out[#out + 1] = { first = blanks[k], last = blanks[k], text = "" }
    end
    blanks = {}
  end
  for row, line in ipairs(lines) do
    local prev = out[#out]
    local fence, fence_attrs = markup.box_line(line)
    if fence then -- a box opens or closes (inside one, ":::" closes it; with braces, another opens)
      local was = box
      if box then
        empty_paragraphs(prev and 2 or 1)
        box = nil
      end
      if not was or fence_attrs then
        empty_paragraphs(prev and 2 or 1)
        box = { start = row, attrs = fence_attrs or "" }
      end
      table_start, cut = nil, true
      goto continue
    end
    box_of_row[row] = box
    local option_line = markup.table_options(line)
    if option_line and (lines[row + 1] or ""):match("^|") then -- (the options of the table below it)
      empty_paragraphs(prev and 2 or 1)
      options[row + 1], table_start = option_line, nil
      goto continue
    end
    table_start = line:match("^|") and (#blanks == 0 and table_start or row) or nil
    local cells = table_start and markup.cells(line, row == table_start + 1) -- (dashes: under the first row)
    if cells and #cells == 0 then -- (a "|" alone: text)
      cells, table_start = false, nil
    end
    if line:match("^%s*$") then
      blanks[#blanks + 1] = row
    elseif cells ~= false and table_start then -- a table row (or the line of dashes under its first row)
      empty_paragraphs(prev and 2 or 1)
      formats[table_start] = not cells and markup.table_format(line) or formats[table_start]
      for c, cell in ipairs(cells or {}) do
        for _, para in ipairs(cell) do
          out[#out + 1] = {
            first = row, last = row, owns = row, text = para.text, col = para.col, stop = para.col + #para.text,
            cell = { table = table_start, row = row, col = c },
          }
        end
      end
    elseif prev and prev.text ~= "" and not prev.cell and not prev.alone and not alone[row] and #blanks == 0 and not cut and not next((markup.block(line))) then -- continues it
      local hard = #(prev.text:match("\\*$")) % 2 == 1 -- (after a "\": a line break, else a space)
      local floating = not hard and markup.floats_only(lines[row - 1]) -- (its floating images, on a line of their own: nothing between)
      if floating then
        prev.joins = prev.joins or {}
        prev.joins[row - prev.first] = 0
      end
      prev.last, prev.text = row, prev.text .. (hard and "\n" or floating and "" or " ") .. line
    else
      empty_paragraphs(prev and 2 or 1) -- (the first blank line after a paragraph separates)
      out[#out + 1] = { first = row, last = row, text = line, alone = alone[row] }
    end
    if not line:match("^%s*$") then
      cut = false
    end
    ::continue::
  end
  empty_paragraphs(1)
  local next_first = #lines + 1
  for k = #out, 1, -1 do
    local cell = out[k].cell
    out[k].box = box_of_row[out[k].first]
    out[k].owns = cell and out[k].first or next_first - 1
    next_first = out[k].first
    local format = cell and formats[cell.table]
    if format then
      cell.align = format.align[cell.col]
    end
    if cell then
      cell.format = options[cell.table]
    end
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

-- How many characters join a paragraph's line `i` (of its lines) to the next in its text: one (a
-- space or a line break), or none (after a line of floating images: see M.paragraphs).
function M.join(para, i)
  return para.joins and para.joins[i] or 1
end

-- Cut a paragraph into pieces: text with uniform formatting, objects, images, and comment/change
-- markers. Reads the buffer and its marks only.
local function pieces(buf, para)
  local doc = M.docs[buf]
  local block, chars = markup.parse(para.text)
  block.cell = para.cell and true or nil
  if para.cell and not block.align and para.cell.align ~= "left" then -- (its column's, unless it has its own)
    block.align = para.cell.align
  end
  -- where each character is in the buffer (the paragraph's lines were joined with spaces)
  local lines = api.nvim_buf_get_lines(buf, para.first - 1, para.last, true)
  local function position(offset) -- (of a byte of the paragraph's text, ending something: row, column)
    local row, start = 1, -(para.col or 0)
    while row < #lines and offset > start + #lines[row] do
      start, row = start + #lines[row] + M.join(para, row), row + 1
    end
    return para.first - 2 + row, offset - start
  end
  local row, start = 1, -(para.col or 0)
  for _, ch in ipairs(chars) do
    if ch.len then -- (an image, equation... may go on in the next line)
      ch.end_row, ch.end_col = position(ch.col + ch.len)
    end
    -- (a character right after a line joined with nothing is on the next line)
    while row < #lines and (ch.col > start + #lines[row] or ch.col == start + #lines[row] and M.join(para, row) == 0) do
      start, row = start + #lines[row] + M.join(para, row), row + 1
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
      local lo, hi = 1, #chars + 1 -- (the first character at or after r, c)
      while lo < hi do
        local mid = math.floor((lo + hi) / 2)
        local ch = chars[mid]
        if ch.row > r or ch.row == r and ch.col >= c then
          hi = mid
        else
          lo = mid + 1
        end
      end
      events[#events + 1] = { lo, rank, #events, piece }
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
  -- (characters come in order, and so do marks (by their start): the first mark that covers a
  -- character is found by stepping past those that ended before it)
  local function walker(list)
    local first = 1
    return function(ch)
      while list[first] and not (list[first][4].end_row > ch.row or list[first][4].end_row == ch.row and list[first][4].end_col > ch.col) do
        first = first + 1
      end
      for k = first, #list do
        local m = list[k]
        if m[2] > ch.row or m[2] == ch.row and m[3] > ch.col then
          return nil
        elseif covers(m, ch.row, ch.col) then
          return m
        end
      end
    end
  end
  local deletions = {}
  for _, m in ipairs(changes) do
    if m[5].type == "delete" then
      deletions[#deletions + 1] = m
    end
  end
  local deletion_at, run_at = walker(deletions), walker(runs)
  local function deleted_by(ch)
    return #deletions > 0 and not (ch.image or ch.math or ch.footnote or ch.field or ch.ch == "￼") and deletion_at(ch)
  end
  for _, m in ipairs(changes) do
    if m[5].type == "delete" and m[5].no_text then
      event(m[2], m[3], 2, { deletion = m[5], deleted = "", part = 0 })
    elseif m[5].type == "delete" then
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
      local last, char = out[#out], ch.ch == "¶" and not ch.esc and "\n" or ch.ch
      if not (last and last.deletion == deletion[5]) then
        parts[deletion[5]] = (parts[deletion[5]] or (deletion[2] < para.first - 1 and 1 or 0)) + 1
        last = { deletion = deletion[5], deleted = "", part = parts[deletion[5]] - 1, segs = {} }
        out[#out + 1] = last
      end
      last.deleted = last.deleted .. char
      -- (its formatting: the deleted text is written formatted, to come back as it was)
      local run = run_at(ch)
      local seg, fmt = last.segs[#last.segs], { props = run and doc.run[run[1]] or nil }
      for _, flag in ipairs(markup.FLAGS) do
        fmt[flag] = ch[flag]
      end
      local same = seg and seg.props == fmt.props
      for _, flag in ipairs(markup.FLAGS) do
        same = same and seg[flag] == fmt[flag]
      end
      if same then
        seg.text = seg.text .. char
      else
        fmt.text = char
        last.segs[#last.segs + 1] = fmt
      end
      goto next_char
    end
    local mark
    if ch.ch == "￼" or ch.image or ch.math or ch.footnote or ch.field then
      for _, m in ipairs(objects) do
        if m[2] == ch.row and m[3] == ch.col and doc.object[m[1]].visible then
          mark = m
          break
        end
      end
    end
    local object = mark and doc.object[mark[1]]
    if ch.image then -- unchanged, edited (another path or alt text) or new
      local same = object and object.image and object.image.path == ch.image.path and object.image.alt == ch.image.alt
        and floats.key(object.image) == floats.key(ch.image)
      assert(not ch.image.bad, ch.image.bad) -- (braces after it that aren't its attributes)
      assert(not ch.image.width or markup.cm(ch.image.width, 1), "image width: write it in cm, mm, in or % (of the text's width), as in {width = 12cm} or {width = 50%}")
      local where = { mark = mark and mark[1], row = ch.row, col = ch.col, end_row = ch.end_row, end_col = ch.end_col }
      out[#out + 1] = same and { object = object.node } or { image = vim.tbl_extend("force", ch.image, where, { original = object }) }
    elseif ch.math then -- an equation: unchanged, or (re)written from its LaTeX
      local same = object and object.math and object.math.tex == ch.math.tex and object.math.display == ch.math.display
      local where = { mark = mark and mark[1], row = ch.row, col = ch.col, end_row = ch.end_row, end_col = ch.end_col }
      out[#out + 1] = same and { object = object.node } or { math = vim.tbl_extend("force", ch.math, where) }
    elseif ch.footnote then -- a footnote: unchanged, edited or new
      local same = object and object.footnote and object.footnote.text == ch.footnote.text
      local where = { mark = mark and mark[1], row = ch.row, col = ch.col, end_row = ch.end_row, end_col = ch.end_col }
      out[#out + 1] = same and { object = object.node, note = object.footnote } -- (note: in case it must be written again)
        or { footnote = vim.tbl_extend("force", ch.footnote, where, { original = object }) }
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
      local run = run_at(ch)
      if run then
        props = doc.run[run[1]]
      end
      local insert
      for _, m in ipairs(changes) do
        if m[5].type == "insert" and covers(m, ch.row, ch.col) then
          insert = m[5]
        end
      end
      -- the same formatting as the text before it: one piece (its text is collected, joined below)
      local last = out[#out]
      local same = last and type(last.text) == "table" and last.props == (props or nil) and last.link == ch.link
        and last.insert == insert
      for _, flag in ipairs(markup.FLAGS) do
        same = same and last[flag] == ch[flag]
      end
      if same then
        last.text[#last.text + 1] = ch.ch
      else
        local piece = { text = { ch.ch }, props = props or nil, link = ch.link, insert = insert }
        for _, flag in ipairs(markup.FLAGS) do
          piece[flag] = ch[flag]
        end
        out[#out + 1] = piece
      end
    end
    ::next_char::
  end
  for _, piece in ipairs(out) do
    if type(piece.text) == "table" then
      piece.text = table.concat(piece.text)
    end
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
  local l = block.list
  local out = { table.concat({ tostring(block.heading), tostring(block.title), tostring(block.quote), tostring(block.align), tostring(block.caption),
    tostring(block.cell), l and table.concat({ l.kind, l.level, tostring(l.n), tostring(l.id) }, ",") or "" }, "|") }
  for _, piece in ipairs(list) do
    local keys = vim.tbl_keys(piece)
    table.sort(keys)
    for _, key in ipairs(keys) do
      local v = piece[key]
      if key == "segs" then -- (the deleted text's formatting)
        v = table.concat(vim.tbl_map(function(seg)
          return tostring(seg.props) .. table.concat(vim.tbl_map(function(f) return tostring(seg[f]) end, markup.FLAGS)) .. seg.text
        end, v), "/")
      elseif type(v) == "table" then -- (identity, plus what can change inside comments, changes, images)
        local replies = table.concat(vim.tbl_map(tostring, v.replies or {}), ",")
        v = table.concat({ tostring(v), replies, tostring(v.resolved), v.text or "", v.path or "", v.alt or "", v.path and floats.key(v) or "", v.tex or "", tostring(v.display) }, "/")
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
  require("neoffice.lsp").start(buf)
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
      M.decorate_soon(buf)
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
local converted = {} -- (a .doc file, as it was -> its .docx, this session)
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

-- Captions: a paragraph that starts with a caption's word, its number (a numbering field of that
-- name, or, in a caption style, the number written) and a separator ("Table 1. ", "Figure 2: ")
-- shows as "Table: ..." (block.caption = "table" | "figure"): the word, number and separator are
-- written again when it is (see M.save), the separator as it was (block.caption_sep).
local SEPARATORS = { ": ", ". ", " | ", " – ", " — ", " - ", ":", ".", " " }
local function separator(text)
  for _, s in ipairs(SEPARATORS) do
    if text:sub(1, #s) == s then
      return s
    end
  end
  return ""
end
local function find_captions(paras)
  for _, para in ipairs(paras) do
    local block, spans = para.block, para.spans
    if not (block.heading or block.title or block.quote or block.list) then
      -- the text before the first thing that isn't text, and that thing
      local lead, k = "", 1
      while spans[k] and spans[k].text and not spans[k].deletion and not spans[k].object do
        lead, k = lead .. spans[k].text, k + 1
      end
      local word, rest = lead:match("^%s*(%S+)%s+(.*)$")
      local kind = word and markup.caption_kind(word)
      local field = spans[k] and spans[k].field
      local cut -- (how many bytes of text to take away from the start, the field included or not)
      if kind and rest == "" and field and field.seq and markup.caption_kind(field.seq) == kind then
        -- "Table " + its number + the separator, in the text after
        local after, j = "", k + 1
        while spans[j] and spans[j].text and not spans[j].object and not spans[j].deletion do
          after, j = after .. spans[j].text, j + 1
        end
        local sep = separator(after)
        block.caption, block.caption_sep = kind, sep ~= "" and sep or nil
        block.caption_label, block.caption_seq = word, field.seq
        for _ = 1, k do
          table.remove(spans, 1)
        end
        cut = #sep
      elseif kind and block.caption_style and rest:match("^%d+") then -- "Table 1. ", written
        local number = rest:match("^%d+")
        local sep = separator(rest:sub(#number + 1))
        if sep ~= "" or #rest == #number then
          block.caption, block.caption_sep = kind, sep ~= "" and sep or nil
          block.caption_label, block.caption_number = word, number -- (numbered by hand: kept so)
          cut = #lead - #rest + #number + #sep
        end
      end
      while cut and cut > 0 and spans[1] and spans[1].text and not spans[1].object do -- (from the text spans)
        local n = math.min(cut, #spans[1].text)
        spans[1].text, cut = spans[1].text:sub(n + 1), cut - n
        if spans[1].text == "" then
          table.remove(spans, 1)
        end
      end
    end
  end
end

-- A floating image (anchored to its paragraph: see floats.lua) shows at the start of its
-- paragraph, wherever its anchor is in the text (its place is measured from the paragraph, so it
-- doesn't move); written, it goes there. (Not one inside a tracked change or a comment's range.)
local function hoist_floats(paras)
  for _, para in ipairs(paras) do
    local depth, moved = 0, {}
    for k, s in ipairs(para.spans) do
      depth = depth + ((s.insert_start or s.comment_start) and 1 or 0) - ((s.insert_end or s.comment_end) and 1 or 0)
      if k > 1 and depth == 0 and s.image and floats.floating(s.image) then
        moved[#moved + 1] = k
      end
    end
    for n = #moved, 1, -1 do
      table.insert(para.spans, 1, table.remove(para.spans, moved[n] + (#moved - n)))
    end
  end
end

-- neoffice works on its own copy of the document (the base: a .docx for .doc files, the template
-- for a new one), and copies it to the file when writing. So saving never depends on the file
-- staying where it was, and a file deleted meanwhile is simply written again.
local function open(buf, path)
  local ext = path:match("%.(%w+)$"):lower()
  local new = not vim.uv.fs_stat(path) -- a new document: created on the first :w
  local base
  if ext == "doc" then -- (converted by LibreOffice: once a session, while the file stays as it was)
    local s = not new and vim.uv.fs_stat(path)
    local key = s and ("%s:%d:%d.%d"):format(path, s.size, s.mtime.sec, s.mtime.nsec)
    if new then
      base = template("docx")
    else
      if not (converted[key] and vim.uv.fs_stat(converted[key])) then
        converted[key] = convert(path, "docx", vim.fn.tempname())
      end
      base = vim.fn.tempname() .. ".docx"
      assert(vim.uv.fs_copyfile(converted[key], base))
    end
  elseif new then
    base = template(ext)
  else
    base = vim.fn.tempname() .. "." .. ext
    assert(vim.uv.fs_copyfile(path, base))
  end
  local fmt = require("neoffice." .. (ext == "doc" and "docx" or ext))
  local ctx, paras = fmt.load(base)
  find_captions(paras)
  hoist_floats(paras)
  ctx.dirty, ctx.files = {}, {} -- parts to write on save (name -> true, or content for new files)
  local doc = { path = path, base = base, is_doc = ext == "doc", fmt = fmt, ctx = ctx, paras = {}, run = {}, object = {}, comment = {}, change = {}, caption_seps = {}, caption_as = {}, boxes = {} }
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

  -- Each table's format, shown in its line of dashes (see markup.table_format): its attributes,
  -- and a column's alignment when all its paragraphs have it (then not shown in each).
  local texts, cols, last_text, cells, formats, column = {}, {}, 0, {}, {}, {}
  do
    local aligns = {}
    for k, para in ipairs(paras) do
      local cell, before = M.cell_of(para.node), nil
      cells[k], before = cell, cells[k - 1]
      if cell then
        column[k] = before and before.tr == cell.tr and column[k - 1] + (before.tc == cell.tc and 0 or 1) or 1
        aligns[cell.tbl] = aligns[cell.tbl] or {}
        local have, align = aligns[cell.tbl][column[k]], para.block.align or "left"
        aligns[cell.tbl][column[k]] = (have == nil or have == align) and align or false
      end
    end
    for k, para in ipairs(paras) do
      local cell = cells[k]
      if cell and not formats[cell.tbl] then
        local have = doc.fmt.table_format and doc.fmt.table_format(doc.ctx, cell.tbl) or {}
        formats[cell.tbl] = { align = {}, options = require("neoffice.tables").line(have) }
        for col, align in pairs(aligns[cell.tbl]) do
          formats[cell.tbl].align[col] = (align == "center" or align == "right") and align or nil
        end
      end
      if cell and formats[cell.tbl].align[column[k]] then
        para.block.align = nil
      end
    end
  end
  for k, para in ipairs(paras) do
    para.block.cell = cells[k] and true or nil
    if cells[k] then -- (spaces at the edges of a cell's text: not shown, as the table can't show them)
      local text = vim.tbl_filter(function(sp) return sp.text and not sp.deletion and not sp.object end, para.spans)
      for _, sp in ipairs(text) do
        sp.text = sp.text:gsub("^%s+", "")
        if sp.text ~= "" then
          break
        end
      end
      for k2 = #text, 1, -1 do
        text[k2].text = text[k2].text:gsub("%s+$", "")
        if text[k2].text ~= "" then
          break
        end
      end
    end
    for _, s in ipairs(para.spans) do
      if cells[k] and (s.line_break or s.page_break) then -- (in a table cell: kept, as ￼)
        s.line_break, s.page_break, s.visible = nil, nil, true
      elseif s.line_break or s.page_break then
        s.object = nil
      end
      if s.deletion then -- (shown in the text, struck through; a paragraph break as ¶¶)
        s.text = s.deletion.text:gsub("\n", "¶¶")
        s.deletion.no_text = s.text == "" or nil -- (a deleted image, tab...: a point, kept as it is)
      elseif s.text then
        s.text = s.text:gsub("\n", " ")
      elseif s.object and s.visible and not (s.image or s.math or s.footnote or s.field) then
        s.text, s.raw = "￼", true
      end
    end
    texts[k], cols[k] = markup.render(para.spans, para.block)
    last_text = (texts[k] ~= "" or cells[k]) and k or last_text
    doc.paras[k] = para.node
    doc.caption_seps[k] = para.block.caption_sep
    doc.caption_as[k] = para.block.caption and { label = para.block.caption_label, seq = para.block.caption_seq, number = para.block.caption_number } or nil
    if para.box then -- (the document's boxes, and how they are placed: see M.save)
      doc.boxes[para.box.node] = floats.key(para.box.attrs)
    end
  end
  doc.texts = texts

  -- Aligned tables: each cell's text padded to its column's width (see tables.text_widths).
  local pad_before, pad_after = {}, {}
  do
    local tables, by_table = require("neoffice.tables"), {}
    for k in ipairs(paras) do
      local cell, before = cells[k], cells[k - 1]
      if cell then
        local t = by_table[cell.tbl] or { rows = {}, cells = {} }
        by_table[cell.tbl] = t
        if not (before and before.tr == cell.tr) then
          t.rows[#t.rows + 1] = {}
        end
        local row = t.rows[#t.rows]
        if before and before.tc == cell.tc then -- (another paragraph of the cell)
          row[#row] = row[#row] .. " <br> " .. texts[k]
          t.cells[#t.cells].last = k
        else
          row[#row + 1] = texts[k]
          t.cells[#t.cells + 1] = { first = k, last = k, row = row, col = #row }
        end
      end
    end
    for tbl, t in pairs(by_table) do
      local widths = tables.text_widths(t.rows)
      formats[tbl].widths = widths
      for _, c in ipairs(t.cells) do
        pad_before[c.first], pad_after[c.last] = tables.padding(c.row[c.col], widths and widths[c.col], formats[tbl].align[c.col])
      end
    end
  end

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
        rows[cell.tbl] = (rows[cell.tbl] or 0) + 1
        if rows[cell.tbl] == 1 and formats[cell.tbl].options ~= "" then -- (its options, above it)
          lines[#lines + 1] = "{" .. formats[cell.tbl].options .. "}"
        end
        lines[#lines + 1], columns = "| ", 1
      end
      r = #lines - 1
      lines[#lines] = lines[#lines] .. (pad_before[k] or "")
      local offset = #lines[#lines]
      lines[#lines] = lines[#lines] .. texts[k] .. (pad_after[k] or "")
      at = function(col) return r, offset + col end
      if not (after and after.tr == cell.tr) then -- the row ends
        lines[#lines] = lines[#lines] .. " |"
        if rows[cell.tbl] == 1 then
          lines[#lines + 1] = markup.format_line(formats[cell.tbl], columns, formats[cell.tbl].widths)
        end
        if not (after and after.tbl == cell.tbl) and k < last_text then
          lines[#lines + 1] = ""
        end
      end
    else
      local box = para.box -- (a box's paragraphs: between its lines, "::: {...}" and ":::")
      if box and (not paras[k - 1] or paras[k - 1].box ~= box) then
        local attrs = floats.render(box.attrs)
        lines[#lines + 1] = attrs ~= "" and "::: {" .. attrs .. "}" or ":::"
      end
      local ends_box = box and (not paras[k + 1] or paras[k + 1].box ~= box)
      r = #lines -- (a paragraph with line breaks takes several lines)
      -- its floating images first, on a line of their own, joined to the text with nothing (see
      -- M.paragraphs), unless the text would then read as something else
      local text, parts, joins = texts[k], {}, {}
      local floating = 0
      while para.spans[floating + 1] and para.spans[floating + 1].image and floats.floating(para.spans[floating + 1].image) do
        floating = floating + 1
      end
      local split = floating > 0 and cols[k][floating] and cols[k][floating][2]
      local rest = split and text:sub(split + 1)
      if rest and not rest:match("^%s*$") and not rest:match("^%s*\n") and not next((markup.block(rest))) and not rest:match("^|")
        and not markup.box_line(rest) and not markup.table_options(rest) then
        parts[1], joins[1], text = text:sub(1, split), 0, rest
      end
      for _, line in ipairs(vim.split(text, "\n", { plain = true })) do
        parts[#parts + 1] = line
      end
      local starts, offset = {}, 0
      for i, line in ipairs(parts) do
        lines[#lines + 1], starts[#starts + 1], offset = line, offset, offset + #line + (joins[i] or 1)
      end
      at = function(col, ends) -- the position of a byte column of the paragraph's text
        local row = #starts
        -- (what ends where a line joined with nothing ends: on that line)
        while starts[row] > col or ends and row > 1 and starts[row] == col and joins[row - 1] == 0 do
          row = row - 1
        end
        return r + row - 1, col - starts[row]
      end
      local next_list = paras[k + 1] and paras[k + 1].block.list
      local tight = para.block.list and next_list and next_list.id == para.block.list.id and texts[k + 1] ~= ""
      if texts[k] ~= "" and k < last_text and not tight and not ends_box then
        lines[#lines + 1] = ""
      end
      if texts[k] == "" then
        empty_rows[#empty_rows + 1] = r
      end
      if ends_box then
        lines[#lines + 1] = ":::"
        if k < last_text then
          lines[#lines + 1] = ""
        end
      end
    end
    local seen_text = false
    for s_index, s in ipairs(para.spans) do
      local ra, ca = at(cols[k][s_index][1])
      local rb, cb = at(cols[k][s_index][2], true)
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
  local block = mode == "\22" and { math.min(s[3], e[3]) - 1, math.max(s[3], e[3]) } or nil -- (its columns)
  return { { s[2] - 1, first_col }, stop, block = block }
end

-- Decorations that follow the text: titles highlighted (NeofficeTitle, above tree-sitter) and,
-- when the document numbers its headings, their numbers (1, 1.1, ...) shown before them.
local decorating = {}
function M.decorate_soon(buf) -- (once typing pauses: decorations are for the whole buffer)
  if not decorating[buf] then
    decorating[buf] = true
    vim.defer_fn(function()
      decorating[buf] = nil
      if api.nvim_buf_is_valid(buf) then
        M.decorate(buf)
      end
    end, 50)
  end
end

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
  for r, line in ipairs(lines) do -- a table's options (above it)
    if line:sub(1, 1) == "{" and markup.table_options(line) and (lines[r + 1] or ""):match("^|") then
      api.nvim_buf_set_extmark(buf, NS_DECOR, r - 1, 0, { end_col = #line, hl_group = "NeofficeTableOptions", priority = 150 })
    end
  end
  for _, para in ipairs(M.blocks(buf)) do
    if para.text:find("==", 1, true) then
      M.decorate_marks(buf, para, lines)
    end
    local block, text_start = markup.block(lines[para.first])
    -- the alignment mark (above Markdown's highlighting, which takes ">..." for a quote)
    local own, own_start, _, align_at = markup.block(para.text)
    if own.align then
      local col = (para.col or 0) + align_at - 1
      api.nvim_buf_set_extmark(buf, NS_DECOR, para.first - 1, col, {
        end_col = (para.col or 0) + own_start - 2, hl_group = "NeofficeAlign", priority = 250,
      })
    end
    if block.caption and not para.cell then -- its number (as the document shows it): "Table 1: ..."
      local key = "caption." .. block.caption
      count[key] = (count[key] or 0) + 1
      api.nvim_buf_set_extmark(buf, NS_DECOR, para.first - 1, #markup.captions()[block.caption], {
        virt_text = { { " " .. count[key], "NeofficeNumber" } }, virt_text_pos = "inline",
      })
      api.nvim_buf_set_extmark(buf, NS_DECOR, para.first - 1, 0, {
        end_col = text_start - 1, hl_group = "NeofficeCaption", priority = 200,
      })
    elseif block.title then
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

-- Highlighted text shown in its colour.
local function highlight_group(color)
  local hex = markup.COLORS[color] or color:gsub("^#", "")
  local name = "NeofficeHighlight_" .. hex
  if vim.tbl_isempty(api.nvim_get_hl(0, { name = name })) then
    local r, g, b = tonumber(hex:sub(1, 2), 16), tonumber(hex:sub(3, 4), 16), tonumber(hex:sub(5, 6), 16)
    local light = 0.299 * r + 0.587 * g + 0.114 * b > 140
    api.nvim_set_hl(0, name, { bg = "#" .. hex, fg = light and "#000000" or "#ffffff", default = true })
  end
  return name
end
function M.decorate_marks(buf, para, lines)
  local _, chars = markup.parse(para.text)
  local function position(offset) -- (the paragraph's lines were joined by one character)
    if para.cell then
      return para.first - 1, para.col + offset
    end
    local row = para.first
    while row < para.last and offset > #lines[row] do
      offset, row = offset - #lines[row] - M.join(para, row - para.first + 1), row + 1
    end
    return row - 1, offset
  end
  local function mark(from, to, group) -- (chars from..to)
    local r1, c1 = position(chars[from].col)
    local r2, c2 = position(chars[to].col + (chars[to].len or #chars[to].ch))
    api.nvim_buf_set_extmark(buf, NS_DECOR, r1, c1, { end_row = r2, end_col = c2, hl_group = group, priority = 150 })
  end
  for _, key in ipairs({ "h" }) do -- (underlines: by the syntax, syntax/neoffice.vim)
    local from
    for k = 1, #chars + 1 do
      local c, prev = chars[k], chars[k - 1]
      if from and not (c and c[key] == prev[key]) then
        mark(from, k - 1, highlight_group(prev[key]))
        from = nil
      end
      if c and c[key] and not from then
        from = k
      end
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
-- The spans that render characters of markup.parse (all they read as), with `flags` changed for
-- those `inside(c)`.
local function spans_of(chars, inside, flags)
  return vim.tbl_map(function(c)
    local span = { text = c.ch ~= "\n" and c.ch or nil, link = c.link, line_break = c.ch == "\n" or nil,
      image = c.image, math = c.math, footnote = c.footnote, field = c.field, page_break = c.page_break }
    for _, flag in ipairs(markup.FLAGS) do
      span[flag] = c[flag]
      if flags and inside(c) and flags[flag] ~= nil then
        span[flag] = flags[flag]
      end
    end
    return span
  end, chars)
end

-- Line `row` rewritten as `new`: only the characters that differ are replaced (a mark moved past a
-- word: the mark, not the word), so the marks around and between them stay where they are.
-- Returns the columns, in the new line, from the first difference to the last.
local function rewrite(buf, row, new)
  local line = api.nvim_buf_get_lines(buf, row, row + 1, true)[1]
  local function chars(s) -- (its characters, and where each starts)
    local list, at = {}, {}
    for pos, ch in s:gmatch("()([^\128-\191][\128-\191]*)") do
      list[#list + 1], at[#at + 1] = ch, pos - 1
    end
    at[#list + 1] = #s
    return list, at
  end
  local old, old_at = chars(line)
  local list = chars(new)
  local hunks = vim.diff(table.concat(old, "\n") .. "\n", table.concat(list, "\n") .. "\n", { result_type = "indices" })
  for k = #hunks, 1, -1 do -- (from the end: the columns before stay right)
    local a1, an, b1, bn = unpack(hunks[k])
    local c1 = an > 0 and old_at[a1] or old_at[a1 + 1]
    local c2 = an > 0 and old_at[a1 + an] or c1
    api.nvim_buf_set_text(buf, row, c1, row, c2, { table.concat(list, "", b1, b1 + bn - 1) })
  end
  local p, s = 0, 0
  while p < math.min(#line, #new) and line:byte(p + 1) == new:byte(p + 1) do
    p = p + 1
  end
  while s < math.min(#line, #new) - p and line:byte(#line - s) == new:byte(#new - s) do
    s = s + 1
  end
  return p, #new - s
end

function M.set_format(buf, row, c1, c2, flags, props)
  local line = api.nvim_buf_get_lines(buf, row, row + 1, true)[1]
  local block, chars = markup.parse(line)
  local new = markup.render(spans_of(chars, function(c) return c.col >= c1 and c.col < c2 end, flags), block)
  local p, e = rewrite(buf, row, new)
  M.place(buf, "run", props, row, p, row, e)
end

-- Delete the text from a to b ({ row, col }, b excluded) keeping the rest as it reads: a mark it
-- shares with the text around (the "^{" of "^{3" in "^{34}", when "4}" goes) stays.
function M.cut(buf, a, b)
  local expected
  if a[1] == b[1] then
    local _, chars = markup.parse(api.nvim_buf_get_lines(buf, a[1], a[1] + 1, true)[1])
    expected = vim.tbl_filter(function(c) return c.col < a[2] or c.col >= b[2] end, chars)
  end
  api.nvim_buf_set_text(buf, a[1], a[2], b[1], b[2], {})
  if not expected then
    return
  end
  local block, now = markup.parse(api.nvim_buf_get_lines(buf, a[1], a[1] + 1, true)[1])
  local function reads(c) -- (all of a character that the text shows)
    local key = { c.ch, c.link or "" }
    for _, flag in ipairs(markup.FLAGS) do
      key[#key + 1] = c[flag] and flag .. tostring(c[flag]) or ""
    end
    return table.concat(key, "\0")
  end
  local same = #now == #expected
  for k = 1, same and #now or 0 do
    same = same and reads(now[k]) == reads(expected[k])
  end
  if not same then
    rewrite(buf, a[1], markup.render(spans_of(expected), block))
  end
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
-- Pair changed paragraphs (old[sa..], new[sb..]) with the old ones they most look like, in order
-- (so an edited paragraph keeps its own properties when the one before it was deleted). Returns
-- new index -> old index.
local function pair(old, new, sa, ca, sb, cb)
  local out = {}
  if ca == cb or ca == 0 or cb == 0 or ca * cb > 2500 then -- (as many, or too many: in order)
    for k = 0, math.min(ca, cb) - 1 do
      out[sb + k] = sa + k
    end
    return out
  end
  local function like(x, y) -- (how many characters they share at their start and end)
    local n, m = 0, 0
    while n < #x and n < #y and x:byte(n + 1) == y:byte(n + 1) do
      n = n + 1
    end
    while m < #x - n and m < #y - n and x:byte(#x - m) == y:byte(#y - m) do
      m = m + 1
    end
    return n + m
  end
  local best = {} -- best[i][j]: the best pairing of the first i old and j new paragraphs
  for i = 0, ca do
    best[i] = { [0] = 0 }
  end
  for j = 1, cb do
    best[0][j] = 0
  end
  for i = 1, ca do
    for j = 1, cb do
      best[i][j] = math.max(best[i - 1][j], best[i][j - 1], best[i - 1][j - 1] + 1 + like(old[sa + i - 1], new[sb + j - 1]))
    end
  end
  local i, j = ca, cb
  while i > 0 and j > 0 do
    if best[i][j] == best[i - 1][j] then
      i = i - 1
    elseif best[i][j] == best[i][j - 1] then
      j = j - 1
    else
      out[sb + j - 1], i, j = sa + i - 1, i - 1, j - 1
    end
  end
  return out
end

local function match_lines(old, new)
  local old_of, deleted, a, b = {}, {}, 1, 1
  local hunks = M.diff(old, new)
  table.insert(hunks, { #old + 1, 0, #new + 1, 0 }) -- so lines after the last hunk get mapped too
  for _, h in ipairs(hunks) do
    local sa, ca, sb, cb = unpack(h)
    while b < sb do -- unchanged lines
      old_of[b], a, b = a, a + 1, b + 1
    end
    for j, i in pairs(pair(old, new, sa, ca, sb, cb)) do
      old_of[j] = i
    end
    local used = {}
    for j = sb, sb + cb - 1 do
      used[old_of[j] or 0] = true
    end
    for i = sa, sa + ca - 1 do
      if not used[i] then
        deleted[#deleted + 1] = i
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
    -- a section break (Word's: in the paragraph's properties) isn't shown, so it isn't deleted with
    -- its paragraph: the paragraph before it ends the section instead (unless it ends one already)
    local section = node.parent and xml.child(xml.child(node, "w:pPr") or {}, "w:sectPr")
    if section then
      local k, before = xml.index(node) - 1, nil
      while k > 0 and not before do
        local c = node.parent.children[k]
        before = c.tag == "w:p" and not vim.tbl_contains(nodes, c) and c or nil
        k = k - 1
      end
      local ppr = before and (xml.child(before, "w:pPr") or xml.append(before, xml.elem("w:pPr"), 1))
      if ppr and not xml.child(ppr, "w:sectPr") then
        xml.remove(section)
        local after = xml.child(ppr, "w:pPrChange") -- (sectPr comes last in pPr, but for this)
        xml.append(ppr, section, after and xml.index(after) or nil)
      end
    end
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
      xml.append(cell, fmt.new_para(template, ctx))
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

-- Everything a save may change in a document's XML, and in what neoffice notes about it (ctx, and
-- the comments and changes: `notes`), to put back if the save fails: each node's children and
-- attributes, and the fields of ctx (one level deep) and of the notes.
local function checkpoint(ctx, notes)
  local nodes, seen = {}, {}
  local function walk(n)
    if seen[n] then
      return
    end
    seen[n] = true
    local attrs = {}
    for k, kv in ipairs(n.attrs or {}) do
      attrs[k] = { kv[1], kv[2] }
    end
    local children = {}
    for k, c in ipairs(n.children) do
      children[k] = c
    end
    nodes[#nodes + 1] = { n, children, n.attrs and attrs }
    for _, c in ipairs(n.children) do
      if c.children then
        walk(c)
      end
    end
  end
  local fields = {}
  for key, value in pairs(ctx) do
    if type(value) == "table" and value.children then -- (an XML tree)
      walk(value)
    elseif type(value) == "table" then -- (a map of parts, styles, ids...)
      fields[key] = {}
      for k, v in pairs(value) do
        fields[key][k] = v
        if type(v) == "table" and v.children and not v.parent then
          walk(v)
        end
      end
    end
  end
  local scalars = vim.tbl_extend("force", {}, ctx)
  local saved = {}
  for _, t in ipairs(notes) do
    saved[t] = vim.tbl_extend("force", {}, t)
  end
  return function()
    for t, copy in pairs(saved) do
      for key in pairs(t) do
        t[key] = nil
      end
      for key, value in pairs(copy) do
        t[key] = value
      end
    end
    for _, entry in ipairs(nodes) do
      local n = entry[1]
      n.children, n.attrs = entry[2], entry[3]
      for _, c in ipairs(n.children) do
        c.parent = n
      end
    end
    for key in pairs(ctx) do
      ctx[key] = nil
    end
    for key, value in pairs(scalars) do
      ctx[key] = fields[key] or value
    end
  end
end

-- Tables typeset now (new ones, and those whose options line changed) get their columns of
-- numbers aligned right, unless their line of dashes says how: { [the row of a table's first
-- line] = { [column] = true } } (rows: of the buffer, 1-based), and its paragraphs (blocks) in
-- `blocks_of`. (On save they are aligned so, and the line of dashes says so: "---:"; formatting
-- does the same.)
function M.number_columns(buf, blocks, old_of)
  local doc = M.docs[buf]
  if not doc then
    return {}, {}
  end
  if not blocks then
    blocks = M.blocks(buf)
    old_of = match_lines(doc.texts, vim.tbl_map(function(p) return p.text end, blocks))
  end
  local tables, by_table, out, blocks_of = require("neoffice.tables"), {}, {}, {}
  for k, block in ipairs(blocks) do
    if block.cell then
      local t = by_table[block.cell.table] or { list = {} }
      by_table[block.cell.table] = t
      t.list[#t.list + 1] = block
      local node = old_of[k] and doc.paras[old_of[k]]
      t.existing = t.existing or (node and M.cell_of(node) or {}).tbl
    end
  end
  for start, t in pairs(by_table) do
    local retyped = not t.existing
    if t.existing and doc.fmt.table_format then
      local ok, want = pcall(tables.parse, t.list[1].cell.format)
      retyped = ok and tables.line(vim.tbl_extend("force", tables.defaults(), want)) ~= tables.line(doc.fmt.table_format(doc.ctx, t.existing))
    end
    local numbers, said, columns = {}, {}, 0
    for _, block in ipairs(retyped and t.list or {}) do
      local col = block.cell.col
      columns, said[col] = math.max(columns, col), said[col] or block.cell.align ~= nil
      if block.cell.row ~= start then -- (not the first row: the header)
        local is = tables.is_number(block.text:sub(select(2, markup.block(block.text))))
        if is ~= nil then
          numbers[col] = is and numbers[col] ~= false
        end
      end
    end
    local right = {}
    for col, yes in pairs(numbers) do
      right[col] = yes and not said[col] or nil
    end
    if next(right) then
      out[start], blocks_of[start] = right, t.list
    end
  end
  return out, blocks_of
end
local function align_numbers(buf, blocks, old_of)
  for start, right in pairs((M.number_columns(buf, blocks, old_of))) do
    local line = api.nvim_buf_get_lines(buf, start, start + 1, true)[1]
    if markup.table_format(line) then -- (a table without its line of dashes is formatted first)
      for _, block in ipairs(blocks) do
        if block.cell and block.cell.table == start and right[block.cell.col] then
          block.cell.align = "right"
        end
      end
      api.nvim_buf_set_lines(buf, start, start + 1, true, { markup.realign(line, vim.tbl_map(function() return "right" end, right)) })
    end
  end
end

function M.save(buf, target)
  local doc = M.docs[buf]
  if not doc then
    return vim.notify("[neoffice] This document wasn't opened correctly: not saving it", vim.log.levels.ERROR)
  end
  target = vim.fn.fnamemodify(target or doc.path, ":p")
  local ok, err = pcall(function()
    local ext = function(path) return (path:match("%.(%w+)$") or ""):lower() end
    if ext(target) ~= ext(doc.path) then
      error(("a .%s document can't be written as .%s (another format)"):format(ext(doc.path), ext(target)))
    end
    if target ~= doc.path and vim.fn.fnamemodify(api.nvim_buf_get_name(buf), ":p") == target then
      doc.path, doc.stamp, doc.synced = target, stamp(target), false -- (:saveas: the document's file now)
    end
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
    align_numbers(buf, blocks, old_of)

    -- 1. Work out every paragraph first: if anything is wrong (a missing image, an equation that
    -- can't be converted), nothing changes.
    local plans, equations = {}, {}
    for k, block in ipairs(blocks) do
      local kind, list = pieces(buf, block)
      for _, piece in ipairs(list) do
        local original = piece.image and piece.image.original and piece.image.original.image
        if piece.image and not (original and original.path == piece.image.path) then -- (a new file)
          piece.image.bytes = read_image(doc, piece.image.path)
        end
      end
      for _, piece in ipairs(list) do
        if piece.math then
          equations[#equations + 1] = piece.math
        end
      end
      plans[k] = { old = old_of[k], block = kind, pieces = list, sig = signature(kind, list), cell = block.cell, box = block.box }
      if block.box and not block.box.parsed then -- its box's attributes (an error if they aren't right)
        local attrs, bad = floats.parse_box(block.box.attrs)
        assert(attrs, bad)
        block.box.parsed = attrs
      end
      for _, piece in ipairs(block.box and list or {}) do -- (a new image in a box: as wide as the box)
        if piece.image and not piece.image.original and not piece.image.width then
          piece.image.width = block.box.parsed.width or "50%"
        end
      end
    end

    if #equations > 0 then
      local nodes = doc.fmt.equations and doc.fmt.equations(doc.ctx, equations)
      assert(nodes, "equations ($...$) can only be written in .docx documents (write \\$ for a dollar sign)")
      for k, eq in ipairs(equations) do
        eq.node = nodes[k]
      end
    end

    -- 2. Rewrite only the paragraphs that changed. (If anything fails before the file is written,
    -- the document's XML is put back as it was: see checkpoint.)
    local ctx, paras, prev, plain, kept, placements, seps, box_keys, written_as = doc.ctx, {}, nil, nil, {}, {}, {}, {}, {}
    local notes = vim.list_extend(vim.tbl_values(doc.change), vim.tbl_values(doc.comment))
    for _, c in pairs(doc.comment) do
      vim.list_extend(notes, c.replies)
    end
    local restore = checkpoint(ctx, notes)
    local written, failure = pcall(function()
    local layout = require("neoffice.tables").start(doc.fmt, ctx, vim.tbl_map(function(plan)
      return { cell = plan.cell, node = plan.old and doc.paras[plan.old] }
    end, plans), M.cell_of)
    local function outside(node) -- (the outermost table it is in, or itself; or a block content
      -- control (a table of contents...) it ends: what comes after it goes after it)
      local cell = M.cell_of(node)
      if cell then
        return outside(cell.tbl)
      end
      local content = node.parent
      local last = content and vim.iter(content.children):rev():find(function(c) return c.tag end)
      if content and content.tag == "w:sdtContent" and last == node and content.parent then
        return outside(content.parent)
      end
      return node
    end
    -- boxes: the buffer's (their paragraphs' box tables) and the document's (their runs or frames)
    local box_of = doc.fmt.box_of or function() end
    local boxes, claimed, made = {}, {}, {}
    for _, plan in ipairs(plans) do
      local run = plan.box and plan.old and box_of(doc.paras[plan.old])
      if run and not boxes[plan.box] and not claimed[run] then
        boxes[plan.box], claimed[run] = run, true
      end
    end
    local outer -- (the last paragraph that isn't in a box)
    for k, plan in ipairs(plans) do
      local node = plan.old and doc.paras[plan.old]
      local prev_box = prev and box_of(prev)
      local placed = layout:node({ cell = plan.cell, node = node, edited = not plan.old or plan.sig ~= doc.sigs[plan.old] }, prev_box and outer or prev)
      if node and (placed and placed ~= node or not placed and M.cell_of(node)) then
        deleted[#deleted + 1], plan.old, node = plan.old, nil, nil -- (moved into or out of a table cell)
      end
      node = placed or node
      if node and box_of(node) ~= (plan.box and boxes[plan.box]) then -- (moved into or out of a box)
        deleted[#deleted + 1], plan.old, node = plan.old, nil, nil
      end
      local anchor = prev and outside(prev) or doc.paras[1] and outside(doc.paras[1])
      if not node and plan.box then -- a new paragraph in a box (a new box, maybe)
        local run = boxes[plan.box]
        if not run then
          local new = doc.fmt.new_box(ctx, plan.box.parsed)
          run, boxes[plan.box], made[new.node] = new.node, new.node, true
        end
        node = doc.fmt.new_para(prev_box == run and prev or plain or prev or doc.paras[1], ctx)
        if prev_box == run then
          xml.append(prev.parent, node, xml.index(prev) + 1)
        else -- (its first: before the box's other paragraphs)
          local content = doc.fmt.box_container(run)
          local first = vim.iter(content.children):find(function(c) return c.tag end)
          xml.append(content, node, first and xml.index(first) or nil)
        end
      elseif not node and prev_box then -- a new paragraph after a box: before the paragraph it is in
        node = doc.fmt.new_para(plain or doc.paras[1], ctx)
        if prev_box.parent then
          xml.append(prev_box.parent.parent, node, xml.index(prev_box.parent))
        elseif outer then
          local after = outside(outer)
          xml.append(after.parent, node, xml.index(after) + 1)
        else
          local first = doc.paras[1] and outside(doc.paras[1])
          xml.append(first.parent, node, xml.index(first))
        end
      elseif not node and not anchor then -- the document had no paragraph at all
        node = doc.fmt.first_para(ctx)
      elseif not node then -- a new paragraph, formatted like the one before it
        node = doc.fmt.new_para(plain or prev or doc.paras[1], ctx)
        xml.append(anchor.parent, node, xml.index(anchor) + (prev and 1 or 0))
      end
      plain = not plan.cell and not plan.box and node or plain
      outer = not plan.box and node or outer
      seps[k] = plan.old and doc.caption_seps[plan.old] or nil
      written_as[k] = plan.old and doc.caption_as[plan.old] or nil
      if not plan.old or plan.sig ~= doc.sigs[plan.old] then
        plan.written = true
        local list = plan.pieces
        local kind = plan.block.caption
        if kind and doc.fmt.caption_number then -- a caption: its word, number and separator first
          -- (a caption that was one: its word, and its number as it was: a field, or written by hand)
          local as = written_as[k] or {}
          local label, plainly = as.label or markup.captions()[kind], { b = false, i = false, u = false, s = false, h = false, sup = false, sub = false }
          local first = vim.iter(plan.pieces):find(function(p) return p.text end)
          plainly.insert = first and first.insert -- (a tracked insertion, as its text)
          seps[k] = seps[k] or (M.config.captions or {}).separator or ". "
          local number = as.number and vim.tbl_extend("force", plainly, { text = as.number })
            or { object = doc.fmt.caption_number(ctx, as.seq or markup.captions()[kind]), visible = true }
          list = vim.list_extend({ vim.tbl_extend("force", plainly, { text = label .. " " }), number, vim.tbl_extend("force", plainly, { text = seps[k] }) }, plan.pieces)
          written_as[k] = { label = as.label, seq = as.seq, number = as.number }
        end
        doc.fmt.fill(ctx, node, plan.block, list)
        if kind and doc.fmt.keep_with_next then -- (a table's caption with its table; a figure with its caption)
          local keep = kind == "table" and node or kind == "figure" and prev and not M.cell_of(prev) and prev
          if keep then
            doc.fmt.keep_with_next(ctx, keep)
          end
        end
        ctx.dirty[doc.fmt.MAIN] = true
        for _, piece in ipairs(plan.pieces) do -- images written now: they are the document's from now on
          local new = piece.image or piece.math or piece.footnote
          if piece.kept then
            kept[#kept + 1] = piece.kept
          end
          if new then -- (its mark, once the file is written)
            local data = { node = new.node, visible = true }
            if piece.image then -- (its path, alt text and attributes, as it now is)
              data.image = { path = new.path, alt = new.alt }
              for _, key in ipairs(floats.KEYS) do
                data.image[key] = new[key]
              end
            end
            data.math = piece.math and { tex = new.tex, display = new.display }
            data.footnote = piece.footnote and { text = new.text, id = new.id }
            placements[#placements + 1] = { new = new, data = data }
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
    -- each box in the paragraph after it (the first that isn't in a box or a table), placed as its
    -- line says; those no longer in the buffer go
    for k, plan in ipairs(plans) do
      local next_plan = plans[k + 1]
      if plan.box and not (next_plan and next_plan.box == plan.box) then -- (its last paragraph)
        local run = boxes[plan.box]
        local j = k + 1
        while plans[j] and (plans[j].box or plans[j].cell) do
          j = j + 1
        end
        local at = plans[j] and paras[j] or doc.fmt.first_para(ctx)
        doc.fmt.attach_box(ctx, run, at)
        local key = floats.key(plan.box.parsed)
        local written = made[run] or doc.boxes[run] ~= key
        for i = k, 1, -1 do
          if plans[i].box ~= plan.box then
            break
          end
          written = written or plans[i].written
        end
        if made[run] or doc.boxes[run] ~= key then
          doc.fmt.place_box(ctx, run, plan.box.parsed)
        end
        if written then
          doc.fmt.fit_box(ctx, run)
          doc.fmt.box_changed(run)
          ctx.dirty[doc.fmt.MAIN] = true
        end
        box_keys[run] = key
      end
    end
    for run in pairs(doc.boxes) do
      if not box_keys[run] and run.parent then
        xml.remove(run)
        ctx.dirty[doc.fmt.MAIN] = true
      end
    end
    if ctx.dirty[doc.fmt.MAIN] and doc.fmt.number_captions then -- (captions added, removed, moved)
      doc.fmt.number_captions(ctx)
    end

    -- 3. Write the parts that changed (nothing at all if nothing did).
    local live = { comments = vim.tbl_values(doc.comment), changes = {}, mark_decisions = doc.mark_decisions }
    for _, ch in pairs(doc.change) do
      live.changes[ch] = true
    end
    local entries = doc.fmt.save(ctx, live)
    if next(entries) then -- the base gets the changes...
      zip.write(doc.base, entries)
      doc.synced = false
    end
    end)
    if not written then
      restore()
      error(failure, 0)
    end
    ctx.dirty, ctx.files, ctx.comments_dirty = {}, {}, false
    -- (the document has them now: what neoffice remembers of it follows)
    for _, p in ipairs(placements) do
      if p.new.mark then
        api.nvim_buf_del_extmark(buf, M.ns.object, p.new.mark)
        doc.object[p.new.mark] = nil
      end
      M.place(buf, "object", p.data, p.new.row, p.new.col, p.new.end_row, p.new.end_col)
    end
    doc.paras, doc.texts, doc.sigs, doc.caption_seps, doc.boxes, doc.caption_as = paras, texts, {}, seps, box_keys, written_as
    for k, block in ipairs(blocks) do -- (those written may have new marks: new images, equations...)
      doc.sigs[k] = plans[k].written and signature(pieces(buf, block)) or plans[k].sig
    end
    local alone = {} -- (empty paragraphs are paragraphs of their own, as when opened: see M.blocks)
    for _, m in ipairs(api.nvim_buf_get_extmarks(buf, M.ns.empty, 0, -1, {})) do
      alone[m[2]] = true
    end
    for _, block in ipairs(blocks) do
      if block.text == "" and not block.cell and not alone[block.first - 1] then
        api.nvim_buf_set_extmark(buf, M.ns.empty, block.first - 1, 0, { invalidate = true, right_gravity = false })
      end
    end
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
