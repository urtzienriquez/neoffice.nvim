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
local pformat = require("neoffice.paras") -- (paragraphs' formats)

-- Whether a paragraph is a plain one (not a heading, title, quote, list item or caption): those
-- have a format of their own, set by the lines of formats (see paras.lua).
local function plain(block)
  return not (block.heading or block.title or block.quote or block.list or block.caption or block.vspace or block.toc)
end
local api = vim.api
local M = {}

M.config = {
  author = vim.env.USER or "neoffice", -- name on your comments and tracked changes
  track_changes = nil, -- record your edits as tracked changes: true/false, nil = as the document says
  conceal = false, -- hide the formatting marks (except on the cursor line)
  show_tabs = true, -- show tabs as a faint →, and non-breaking spaces as a faint · (they look like spaces otherwise, and aren't)
  styles = true, -- the document's paragraph styles, at its top ("---" / "styles:" / "  Normal: {...}" / "---")
  page = true, -- its page settings: at its top ("page: {size = A4, margins = 2.5cm}"), and where a section starts ("\\newpage {orientation = landscape}")
  images = true, -- show images in the buffer (needs Neovim's vim.ui.img and a terminal that supports it)
  lsp = true, -- a language server for documents: headings as symbols (gO), folding, hover on comments
  comment_text = false, -- each comment's author and text at the end of its line (:Neoffice comments text)
  -- where .doc files converted by LibreOffice are kept, by their content, so one opens at once the
  -- next time (false: converted each time; nil: Neovim's cache directory)
  doc_cache = nil,
  -- PDF exports (:Neoffice pdf; :help neoffice-pdf): next to the document, by LibreOffice
  -- auto: after each save, in every document (:Neoffice pdf auto turns it on or off in one);
  -- open: open the PDF the first time it is written (viewer: a command, else the system's)
  pdf = { auto = false, open = true, viewer = nil, quiet = false, timeout = 120 },
  image_width = 60, -- at most this many columns wide
  image_height = 20, -- and this many lines high
  -- captions ("Table: ...", "Figure: ..."): their words, and what goes between the number and the text
  captions = { table = "Table", figure = { "Figure", "Fig." }, separator = ". " },
  -- how tables look, unless the line above one says otherwise ("{width = full}"; :help neoffice-table-format)
  tables = { rules = "booktabs", width = "auto", height = "compact", position = "center" },
  -- neoffice's keys (false drops one, or all: keymaps = false; :help neoffice-keymaps): in documents,
  -- only g? (the rest are yours to set); in the comments panel, two keys a family, as fugitive's
  keymaps = {
    help = "g?", -- in documents: this help
    comments = {
      go = "<CR>", reply = "aa", edit = "ee", resolve = "rr", delete = "dd", text = "tt",
      next = "]]", prev = "[[", undo = "u", redo = "<C-r>", close = { "gq", "q" }, help = "g?",
    },
  },
}

M.ns = {}
for _, kind in ipairs({ "run", "object", "comment", "change", "empty", "section" }) do
  M.ns[kind] = api.nvim_create_namespace("neoffice_" .. kind)
end
local NS_DECOR = api.nvim_create_namespace("neoffice_decorations")

-- buf -> { path, base (neoffice's copy, see open()), is_doc, stamp, synced, fmt, ctx,
--          paras (nodes), texts, sigs, tracking,
--          run = {}, object = {}, comment = {}, change = {} (extmark id -> data) }
M.docs = {}

function M.setup(opts)
  opts = vim.deepcopy(opts or {})
  if opts.panel_keys then -- (the comments panel's keys, as they were set before keymaps)
    local km = type(opts.keymaps) == "table" and opts.keymaps or {}
    km.comments = km.comments or {}
    for old, key in pairs(opts.panel_keys) do
      local new = old == "jump" and "go" or old
      if km.comments[new] == nil then
        km.comments[new] = key
      end
    end
    opts.keymaps, opts.panel_keys = km, nil
  end
  local keymaps = opts.keymaps
  M.config = vim.tbl_deep_extend("force", M.config, opts)
  if keymaps == false then
    M.config.keymaps = false
  end
end

-- :help at a tag, across the whole width (at the bottom, or the top: as 'splitbelow' says), as
-- zotero.nvim and fugitive open theirs.
function M.help(tag)
  vim.cmd((vim.o.splitbelow and "botright" or "topleft") .. " help " .. (tag or "neoffice"))
end

-- Create an extmark of `kind` holding `data`. Marks start with left gravity so a range can never
-- end before it starts; runs also grow at their end, so text typed after a word takes its format.
function M.place(buf, kind, data, r1, c1, r2, c2)
  if M.docs[buf] and M.docs[buf].sigs and not M.docs[buf].opening then -- (a change but not to the text: after the fingerprints)
    M.finish_sigs(buf)
    M.docs[buf].all_dirty = true
  end
  local opts = { end_row = r2, end_col = c2, right_gravity = false, end_right_gravity = kind == "run" }
  if kind == "comment" or kind == "change" then
    opts = vim.tbl_extend("force", opts, require("neoffice." .. kind .. "s").style(data))
  end
  local id = api.nvim_buf_set_extmark(buf, M.ns[kind], r1, c1, opts)
  M.docs[buf][kind][id] = data
  if kind == "comment" and M.docs[buf].ctx then -- (shown: one deleted from the buffer goes from the document)
    M.docs[buf].ctx.shown_comments = M.docs[buf].ctx.shown_comments or {}
    M.docs[buf].ctx.shown_comments[data] = true
  end
  if kind == "change" then -- (see changes.lua: changed)
    M.docs[buf].change_version = (M.docs[buf].change_version or 0) + 1
  end
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
  local dashes_seen = {} -- (the tables whose line of dashes was found: see markup.is_dashes)
  local box, box_of_row, cut = nil, {}, false -- (cut: a box's line came just before: no paragraph goes on across it)
  local format, fenced = { attrs = {} }, false -- (the paragraphs' format: see paras.lua; fenced: a format's line came last)
  local pending -- (format's lines for the paragraph right below them only, no "+": { lines, row, last })
  local function empty_after() -- (how many blank lines after a paragraph, or a format's line, separate)
    if box and #out == box.count then -- (right after a box's line: nothing in it to separate from)
      return 1
    end
    return (out[#out] or fenced) and 2 or 1
  end
  local function empty_paragraphs(from)
    for k = from, #blanks do
      -- (around: the blank lines it is among and the lines on either side, what makes it one: see touched)
      out[#out + 1] = { first = blanks[k], last = blanks[k], text = "", format = format, around = { blanks[1] - 1, blanks[#blanks] + 1 } }
    end
    blanks = {}
  end
  local header = markup.header(lines) -- (the document's styles, at its top: no paragraphs)
  local section -- (a section's line, for the paragraph after it: see pages.lua)
  for row, line in ipairs(lines) do
    if section and #out > section.count then
      out[section.count + 1].section, section = section, nil
    end
    local prev = out[#out]
    if header and row <= header.last then
      fenced = row == header.last or fenced -- (a blank line after it separates)
      goto continue
    end
    local fence, fence_attrs = markup.box_line(line)
    if fence then -- a box opens or closes (inside one, ":::" closes it; with braces, another opens)
      local was = box
      if box then
        empty_paragraphs(empty_after())
        box = nil
      end
      if not was or fence_attrs then
        empty_paragraphs(empty_after())
        box = { start = row, attrs = fence_attrs or "", count = #out }
      end
      table_start, cut = nil, true
      goto continue
    end
    box_of_row[row] = box
    local section_line = not box and line:find("^\\newpage") and markup.section_line(line)
    if section_line then -- a section starts after it: what it changes
      empty_paragraphs(empty_after())
      if section then
        section_line.bad = section_line.bad or ("line %d: two sections' lines with no paragraph between them (a section needs one)"):format(section.row)
      end
      section_line.row, section_line.count = row, #out
      section, cut, fenced, table_start = section_line, true, true, nil
      goto continue
    end
    local para_line = not box and line:find("^[<%-{]") and not (lines[row + 1] or ""):match("^|") and markup.para_line(line)
    if para_line and para_line.bad and line:sub(1, 1) == "{" and pcall(require("neoffice.tables").parse, markup.table_options(line) or "") then
      empty_paragraphs(empty_after()) -- (a table's options with no table under it, deleted: nothing,
      cut, fenced, table_start = true, true, nil -- neither text nor a format)
      goto continue
    end
    if para_line then -- the format of the paragraphs after it ("+"), or of the next one: what it changes (see paras.lua)
      empty_paragraphs(empty_after())
      if para_line.sticky then
        local whole = markup.apply_para_line(format, para_line)
        whole.bad, whole.text = para_line.bad, para_line.text
        format = whole
        if pending then -- (between a paragraph's own line and it: still right above it)
          pending.last = row
        end
      else -- (two in a row for one paragraph: both, the second over the first)
        pending = { lines = vim.list_extend(pending and pending.lines or {}, { para_line }), row = pending and pending.row or row, last = row }
      end
      cut, fenced = true, true
      table_start = nil
      goto continue
    end
    local option_line = markup.table_options(line)
    if option_line and (lines[row + 1] or ""):match("^|") then -- (the options of the table below it)
      empty_paragraphs(empty_after())
      options[row + 1], table_start = option_line, nil
      goto continue
    end
    table_start = line:match("^|") and (#blanks == 0 and table_start or row) or nil
    local dashes = table_start and markup.is_dashes(lines, table_start, row, dashes_seen[table_start]) -- (see markup.is_dashes)
    local cells = table_start and not dashes and markup.cells(line) or nil
    if dashes then
      dashes_seen[table_start] = true
    end
    if cells and #cells == 0 then -- (a "|" alone: text)
      cells, table_start = false, nil
    end
    if line:match("^%s*$") then
      blanks[#blanks + 1] = row
      if pending then -- (a paragraph's own format's line, with no paragraph right below it: for none)
        out.stray_format, pending = out.stray_format or pending.row, nil
      end
    elseif cells ~= false and table_start then -- a table row (or the line of dashes under its first row)
      empty_paragraphs(empty_after())
      formats[table_start] = not cells and markup.table_format(line) or formats[table_start]
      for c, cell in ipairs(cells or {}) do
        for _, para in ipairs(cell) do
          out[#out + 1] = {
            first = row, last = row, owns = row, text = para.text, col = para.col, stop = para.col + #para.text,
            cell = { table = table_start, row = row, col = c },
          }
        end
      end
    elseif prev and prev.text ~= "" and not prev.cell and not prev.alone and not prev.vspace and not alone[row] and #blanks == 0 and not cut and not next((markup.block(line)))
      -- (a line of floating images right before a list item, a heading...: that one's, see below)
      and not (markup.floats_only(line) and lines[row + 1] and next((markup.block(lines[row + 1])))) then -- continues it
      local hard = #(prev.text:match("\\*$")) % 2 == 1 -- (after a "\": a line break, else a space)
      -- (its floating images, on a line of their own, before or after its text: nothing between; after
      -- it, also with text typed after them)
      local floating = not hard and (markup.floats_only(lines[row - 1]) or markup.floats_only(line, true))
      if floating then
        prev.joins = prev.joins or {}
        prev.joins[row - prev.first] = 0
      end
      prev.last, prev.text = row, prev.text .. (hard and "\n" or floating and "" or " ") .. line
    elseif prev and prev.first == row - 1 and not prev.cell and not prev.alone and not prev.vspace and not alone[row] and #blanks == 0 and not cut
      and markup.floats_only(lines[row - 1]) and not next((markup.block(lines[row - 1]))) and next((markup.block(line))) then
      -- a paragraph's floating images on their own line, and its marks ("<=> ") with its text after
      local _, start = markup.block(line)
      local prefix = line:sub(1, start - 1)
      prev.lead = { prefix = #prefix, floats = #prev.text }
      prev.joins = { [1] = 0 }
      prev.last, prev.text = row, prefix .. prev.text .. line:sub(start)
    else
      empty_paragraphs(empty_after()) -- (the first blank line after a paragraph separates)
      local kind = markup.block(line) -- (vertical space, a table of contents: a line of its own, the next not joined to it)
      out[#out + 1] = { first = row, last = row, text = line, alone = alone[row], vspace = (kind.vspace or kind.toc) and true or nil }
    end
    if out[#out] and out[#out].first == row then
      out[#out].format, out[#out].run = format, format -- (run: the format of the paragraphs around it, from "+" lines)
      if pending and pending.last ~= row - 1 then -- (not right below it: for none)
        out.stray_format, pending = out.stray_format or pending.row, nil
      end
      if pending then -- (a format's line for it only: over its run's)
        local own = format
        for _, f in ipairs(pending.lines) do
          local whole = markup.apply_para_line(own, f)
          whole.bad, whole.text = own.bad or f.bad, f.text
          own = whole
        end
        local kind = (markup.block(line))
        if out[#out].cell or box_of_row[row] or kind.heading or kind.title or kind.quote or kind.list or kind.caption or kind.vspace then
          own.bad = own.bad or ("line %d: a format's line without \"+\" is for the paragraph right after it, and that is no plain paragraph (a heading, a list item, a table, a box...)"):format(pending.row)
        end
        out[#out].format, pending = own, nil
      end
    end
    if not line:match("^%s*$") then
      cut, fenced = false, false
    end
    ::continue::
  end
  empty_paragraphs(1)
  if pending then
    out.stray_format = out.stray_format or pending.row
  end
  if section and #out > section.count then
    out[section.count + 1].section, section = section, nil
  end
  if box then -- (a box never closed: what follows it would all go into it; said when saving)
    out.unclosed_box = box.start
  end
  if section then -- (nothing after it: said when saving)
    section.bad = "nothing after a section's line: a section needs a paragraph (delete the line to join it to the one before)"
    if out[#out] then
      out[#out].section_after = section
    else
      out.section_after = section
    end
  end
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

-- Where a byte of a paragraph's text is in the buffer: row (1-based), column (0-based).
function M.position_in(buf, para, offset)
  if para.cell then
    return para.first, (para.col or 0) + offset
  end
  local lines = api.nvim_buf_get_lines(buf, para.first - 1, para.last, false)
  offset = M.buffer_offset(para, offset)
  local row = 1
  while row < #lines and offset > #lines[row] do
    offset, row = offset - #lines[row] - M.join(para, row), row + 1
  end
  return para.first + row - 1, offset
end

-- (a paragraph edited: a mark that now closes nothing, where the formatting it closed is gone, as
-- when a word is deleted with the "**" in it; or a "^{" or "_{" no longer closed, so what follows is
-- a superscript or subscript: the save stops, and says where; see markup.orphan_marks)
local function check_marks(buf, block, was)
  local function count(text, list)
    local n = {}
    for _, o in ipairs(list or markup.orphan_marks(text, true)) do
      n[o.mark] = (n[o.mark] or 0) + 1
    end
    return n
  end
  local now = markup.orphan_marks(block.text, true)
  if #now == 0 then
    return
  end
  local before, flags_was, flags_now = count(was), {}, {}
  for _, pair in ipairs({ { was, flags_was }, { block.text, flags_now } }) do
    for _, c in ipairs(select(2, markup.parse_shared(pair[1]))) do
      for _, flag in ipairs(markup.FLAGS) do
        if c[flag] then
          pair[2][flag] = (pair[2][flag] or 0) + 1
        end
      end
    end
  end
  local seen = {}
  for _, o in ipairs(now) do
    seen[o.mark] = (seen[o.mark] or 0) + 1
    local new = seen[o.mark] > (before[o.mark] or 0)
    local lost = (flags_now[o.flag] or 0) < (flags_was[o.flag] or 0)
    if new and (o.spreads or lost) then
      local row, col = M.position_in(buf, block, o.col)
      local escaped = o.mark:gsub(".", "\\%0")
      assert(false, o.spreads
        and ("line %d, column %d: this %s isn't closed any more, so all that follows it would be a %s; close it with }"):format(row, col + 1, o.mark, o.flag == "sup" and "superscript" or "subscript")
        or ("line %d, column %d: this %s closes nothing any more (its other one was deleted), so it would be written as characters and the %s it closed lost; put the other one back, or write %s for the characters"):format(
          row, col + 1, o.mark, ({ b = "bold", i = "italics", s = "strikethrough", h = "highlight", u = "underline" })[o.flag] or "formatting", escaped))
    end
  end
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

-- Where a byte of a paragraph's text is in its lines, as if they were one after the other (joined
-- as M.join says). The same, but for a paragraph whose floating images are on a line of their own
-- and its marks ("<=> ") on the line after them (para.lead = { prefix, floats: their lengths}):
-- its text is marks, images, the rest; its lines, images, then marks and the rest.
function M.buffer_offset(para, offset, ends)
  local lead = para.lead
  if not lead then
    return offset
  elseif ends and offset == lead.prefix + lead.floats then -- (the end of its images: their line's)
    return lead.floats
  elseif offset < lead.prefix then
    return lead.floats + offset
  elseif offset < lead.prefix + lead.floats then
    return offset - lead.prefix
  end
  return offset
end

-- Cut a paragraph into pieces: text with uniform formatting, objects, images, and comment/change
-- markers. Reads the buffer and its marks only.
-- The run properties for the text of a new paragraph (which has none): those of most of the text
-- of the nearest paragraph above it (else below) of the same kind (a heading's of a heading's...),
-- as a paragraph typed in Word or LibreOffice goes on with the formatting of the one before it
-- (its font and size, often set on the text rather than in a style). nil if there is none.
local function neighbour_props(buf, para, block)
  local doc = M.docs[buf]
  local function kind(b)
    return table.concat({ tostring(b.heading or false), tostring(b.title or false), tostring(b.list and true or false),
      tostring(b.quote or false), tostring(b.caption or false) }, "|")
  end
  local want = kind(block)
  local last = api.nvim_buf_line_count(buf) - 1
  -- (a list item or a quote with none of its kind near: as the plain paragraphs, as a form's text)
  local wants = { want }
  if block.list or block.quote then
    wants[2] = kind({})
  end
  for _, wanted in ipairs(wants) do
    want = wanted
    for _, dir in ipairs({ -1, 1 }) do
      local from = dir < 0 and { para.first - 2, -1 } or { para.owns, 0 }
      while from[1] >= 0 and from[1] <= last do
        local found = api.nvim_buf_get_extmarks(buf, M.ns.run, from, dir < 0 and 0 or -1, { limit = 1 })[1]
        if not found then
          break
        end
        local row = found[2]
        local line = api.nvim_buf_get_lines(buf, row, row + 1, true)[1]
        if kind((markup.block(line))) == want then
          local best, size
          for _, m in ipairs(api.nvim_buf_get_extmarks(buf, M.ns.run, { row, 0 }, { row, -1 }, { details = true })) do
            local len = (m[4].end_row == row and m[4].end_col or #line) - m[3]
            if doc.run[m[1]] and (not size or len > size) then
              best, size = doc.run[m[1]], len
            end
          end
          if best then
            return best
          end
        end
        from = dir < 0 and { row - 1, -1 } or { row + 1, 0 }
      end
    end
  end
end

-- The run properties for the text of a new paragraph in a table cell (a row pasted, a cell typed):
-- those of most of the cell in its column in the nearest row above it (else below) that has text,
-- else of any cell in the nearest rows: as its row and column look.
local function cell_props(buf, para)
  local doc = M.docs[buf]
  local last = api.nvim_buf_line_count(buf) - 1
  local here = para.first - 1
  for _, any in ipairs({ false, true }) do
    for _, dir in ipairs({ -1, 1 }) do
      local r = here + dir
      while r >= 0 and r <= last do
        local line = api.nvim_buf_get_lines(buf, r, r + 1, true)[1]
        if not line:match("^|") then
          break
        end
        local spans = not markup.table_format(line) and markup.cell_spans(line) or {}
        local best, size
        for c, sp in ipairs(spans) do
          if any or c == para.cell.col then
            for _, m in ipairs(api.nvim_buf_get_extmarks(buf, M.ns.run, { r, sp.s - 1 }, { r, sp.e }, { details = true, overlap = true })) do
              local s = math.max(m[3], sp.s - 1)
              local e = m[4].end_row == r and math.min(m[4].end_col, sp.e) or sp.e
              if doc.run[m[1]] and e > s and (not size or e - s > size) then
                best, size = doc.run[m[1]], e - s
              end
            end
          end
        end
        if best then
          return best
        end
        r = r + dir
      end
    end
  end
end

-- The font and size of most of a table's text (each its run's own, or false: its style's), as the
-- buffer's marks say: what a cell's text has where it says nothing ("{size = 7pt}" in a cell says it).
-- nil for a key if no text in it has a run (a new table).
local function table_fonts(buf, para)
  local doc = M.docs[buf]
  local key = api.nvim_buf_get_changedtick(buf) .. ":" .. para.cell.table
  doc.table_fonts = doc.table_fonts or {}
  if doc.table_fonts.tick ~= api.nvim_buf_get_changedtick(buf) then
    doc.table_fonts = { tick = api.nvim_buf_get_changedtick(buf) }
  end
  if doc.table_fonts[key] then
    return doc.table_fonts[key]
  end
  local own_of, known = doc.fmt.run_font or function() return {} end, {}
  local weight, order = {}, {}
  for _, k in ipairs(markup.SPAN_KEYS) do
    weight[k], order[k] = {}, {}
  end
  local last, r = api.nvim_buf_line_count(buf) - 1, para.cell.table - 1
  while r <= last do
    local line = api.nvim_buf_get_lines(buf, r, r + 1, true)[1]
    if not line:match("^|") then
      break
    end
    for _, sp in ipairs(not markup.table_format(line) and markup.cell_spans(line) or {}) do
      for _, m in ipairs(api.nvim_buf_get_extmarks(buf, M.ns.run, { r, sp.s - 1 }, { r, sp.e }, { details = true, overlap = true })) do
        local props = doc.run[m[1]]
        local s = math.max(m[3], sp.s - 1)
        local e = m[4].end_row == r and math.min(m[4].end_col, sp.e) or sp.e
        if props and e > s then
          known[props] = known[props] or own_of(doc.ctx, props)
          for _, k in ipairs(markup.SPAN_KEYS) do
            local v = known[props][k] or false
            local measure = v and (pformat.measure(k, v) or v) or false
            if k == "color" and measure == "000000" then -- (black: as none of its own, automatic)
              v, measure = false, false
            end
            if not weight[k][measure] then
              order[k][#order[k] + 1], weight[k][measure] = measure, { n = 0, v = v }
            end
            weight[k][measure].n = weight[k][measure].n + e - s
          end
        end
      end
    end
    r = r + 1
  end
  local out = {}
  for _, k in ipairs(markup.SPAN_KEYS) do
    local best
    for _, m in ipairs(order[k]) do
      if best == nil or weight[k][m].n > weight[k][best].n then -- (false, its style's, is a value)
        best = m
      end
    end
    if best ~= nil then
      out[k] = weight[k][best].v
    end
  end
  doc.table_fonts[key] = out
  return out
end

-- Where each comment starts and ends, as a save writes it: where a paragraph can take it. A start
-- (or an end) on a table's "|", in the spaces around a cell's text, on its line of dashes or on a
-- line that is no paragraph's goes to the nearest place that is one's: a start forward, an end
-- back; one that would then end before it starts is at a point. { [mark id] = { r1, c1, r2, c2 } }
local function settle_comments(buf, blocks)
  local slots = {} -- (in order: what each paragraph takes, from a position to a position)
  for _, b in ipairs(blocks) do
    if b.cell then
      slots[#slots + 1] = { b.first - 1, b.col, b.first - 1, b.stop }
    else
      local last = api.nvim_buf_get_lines(buf, b.owns - 1, b.owns, false)[1] or ""
      slots[#slots + 1] = { b.first - 1, 0, b.owns - 1, #last }
    end
  end
  local function before(r1, c1, r2, c2) -- (r1, c1 strictly before r2, c2)
    return r1 < r2 or r1 == r2 and c1 < c2
  end
  local out = {}
  if #slots == 0 then
    return out
  end
  for _, m in ipairs(api.nvim_buf_get_extmarks(buf, M.ns.comment, 0, -1, { details = true })) do
    local r1, c1, r2, c2 = m[2], m[3], m[4].end_row or m[2], m[4].end_col or m[3]
    local point = r1 == r2 and c1 == c2
    local s1, t1 -- (the start: the first slot that ends at or after it)
    for _, sl in ipairs(slots) do
      if not before(sl[3], sl[4], r1, c1) then
        s1, t1 = before(r1, c1, sl[1], sl[2]) and sl[1] or r1, before(r1, c1, sl[1], sl[2]) and sl[2] or c1
        break
      end
    end
    if not s1 then
      s1, t1 = slots[#slots][3], slots[#slots][4]
    end
    local s2, t2 = s1, t1
    if not point then -- (the end: the last slot that starts at or before it)
      s2, t2 = nil, nil
      for k = #slots, 1, -1 do
        local sl = slots[k]
        if not before(r2, c2, sl[1], sl[2]) then
          s2, t2 = before(sl[3], sl[4], r2, c2) and sl[3] or r2, before(sl[3], sl[4], r2, c2) and sl[4] or c2
          break
        end
      end
      if not s2 or before(s2, t2, s1, t1) then
        s2, t2 = s1, t1
      end
    end
    if s1 ~= r1 or t1 ~= c1 or s2 ~= r2 or t2 ~= c2 then
      out[m[1]] = { s1, t1, s2, t2 }
    end
  end
  return out
end

-- (new: a paragraph the document hasn't yet, whose text takes the formatting of the one next to it)
local function pieces(buf, para, new)
  local doc = M.docs[buf]
  local block, chars = markup.parse(para.text)
  block.cell = para.cell and true or nil
  if block.list and not para.cell then -- (no blank line before it: it goes on with the list above it)
    local above = para.first > 1 and api.nvim_buf_get_lines(buf, para.first - 2, para.first - 1, true)[1] or nil
    block.list.continues = above and not above:match("^%s*$") or nil
  end
  if para.cell and not block.align and para.cell.align ~= "left" then -- (its column's, unless it has its own)
    block.align = para.cell.align
  end
  local cell_own -- (a cell's: its text's font, size and colour of its own, against its table's; nothing else)
  if para.cell then
    assert(not block.para_bad, block.para_bad)
    for k in pairs(block.para or {}) do
      assert(vim.tbl_contains(markup.SPAN_KEYS, k), ("in a table cell, only the font, size and colour of its text can be given: %s"):format(k))
    end
    cell_own = block.para
  end
  if plain(block) and not para.cell and not para.box then -- its format: its run's, and what it says (see paras.lua)
    local format = pformat.apply(para.format or { attrs = {} }, block.align, block.para)
    block.align, block.para = format.align, format.attrs
  else
    -- (its alignment as it shows: "left" too, against a style that isn't, as a centred title's: <<-)
    assert(not (para.box and block.para and block.para.style),
      ("line %d: a style is for paragraphs of the text, not one in a table cell or a box"):format(para.first))
    block.para = nil -- (not a plain paragraph's: its format isn't neoffice's to change)
  end
  -- where each character is in the buffer (the paragraph's lines were joined with spaces)
  local lines = api.nvim_buf_get_lines(buf, para.first - 1, para.last, true)
  local function position(offset) -- (of a byte of the paragraph's text, ending something: row, column)
    offset = M.buffer_offset(para, offset, true)
    local row, start = 1, -(para.col or 0)
    while row < #lines and offset > start + #lines[row] do
      start, row = start + #lines[row] + M.join(para, row), row + 1
    end
    return para.first - 2 + row, offset - start
  end
  if not para.lead and #lines == 1 then -- (on one line, nothing before its text: where each is, directly)
    local row0, shift = para.first - 1, para.col or 0
    for _, ch in ipairs(chars) do
      if ch.len then
        ch.end_row, ch.end_col = row0, ch.col + ch.len + shift
      end
      ch.row, ch.col = row0, ch.col + shift
    end
  else
    local row, start = 1, -(para.col or 0)
    for _, ch in ipairs(chars) do
      if ch.len then -- (an image, equation... may go on in the next line)
        ch.end_row, ch.end_col = position(ch.col + ch.len)
      end
      ch.col = M.buffer_offset(para, ch.col)
      -- (a character right after a line joined with nothing is on the next line)
      while row < #lines and (ch.col > start + #lines[row] or ch.col == start + #lines[row] and M.join(para, row) == 0) do
        start, row = start + #lines[row] + M.join(para, row), row + 1
      end
      ch.row, ch.col = para.first - 2 + row, ch.col - start
    end
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
      local before = chars[lo - 1]
      if piece.comment_start and before and before.len and (before.end_row > r or before.end_row == r and before.end_col > c) then
        lo = lo - 1 -- (a comment starting inside a footnote's text, an image...: on it, not after it)
      end
      events[#events + 1] = { lo, rank, #events, piece }
      return events[#events]
    end
  end
  local comment_marks, listed = marks("comment"), {}
  for _, m in ipairs(comment_marks) do
    listed[m[1]] = true
  end
  for id, at in pairs(doc.comment_at or {}) do -- (and those settled here from where no paragraph is: see settle_comments)
    if not listed[id] and (at[1] <= para.owns - 1 and at[3] >= para.first - 1) then
      local m = api.nvim_buf_get_extmark_by_id(buf, M.ns.comment, id, { details = true })
      if m[1] then
        comment_marks[#comment_marks + 1] = { id, m[1], m[2], m[3] }
      end
    end
  end
  for _, m in ipairs(comment_marks) do
    local c, d = doc.comment[m[1]], m[4]
    local at = doc.comment_at and doc.comment_at[m[1]] or { m[2], m[3], d.end_row, d.end_col } -- (see settle_comments)
    local point = at[1] == at[3] and at[2] == at[4]
    local start = event(at[1], at[2], 3, { comment_start = c, point = point })
    if not point then
      local stop = event(at[3], at[4], 1, { comment_end = c })
      if start and stop and stop[1] <= start[1] then -- (nothing between them: at a point, never ended before it starts)
        start[4].point = true
        events[#events] = nil
      end
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
    elseif para.cell then -- (in a table's cell: what of it is in the cell, a row inserted whole too)
      local row, col, stop = para.first - 1, para.col or 0, para.stop or para.col or 0
      local sr, sc, er, ec = m[2], m[3], m[4].end_row, m[4].end_col
      if sr < row or sr == row and sc < col then
        sr, sc = row, col
      end
      if er > row or er == row and ec > stop then
        er, ec = row, stop
      end
      if sr == row and er == row and sc < ec then
        event(sr, sc, 3, { insert_start = m[5] })
        event(er, ec, 1, { insert_end = m[5] })
      end
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

  -- fields (citations...) whose text was typed again (a line retyped: cc, dd and p): the paragraph's
  -- field that shows that text and is at no field's place any more, rather than text that looks like it
  -- (lost: its mark is empty, its text deleted; this paragraph's first, then any, as when a paragraph
  -- was moved with dd and p; its mark then goes to its new place)
  -- (an image the same way: the one of the same file inside the document)
  local used_image = {} -- (image marks already matched: each once)
  local function lost_field(ch)
    local everywhere = api.nvim_buf_get_extmarks(buf, M.ns.object, 0, -1, { details = true })
    for _, list in ipairs({ objects, everywhere }) do
      for _, m in ipairs(list) do
        local object = doc.object[m[1]]
        local empty = m[4].end_row == m[2] and m[4].end_col == m[3]
        local same = object and (ch.field and object.field and object.field.text == ch.field.text
          or ch.image and object.image and object.image.path == ch.image.path and doc.ctx.entries[ch.image.path])
        if empty and object and object.visible and same then
          api.nvim_buf_set_extmark(buf, M.ns.object, ch.row, ch.col, {
            id = m[1], end_row = ch.end_row, end_col = ch.end_col, right_gravity = false, end_right_gravity = false,
          })
          return object
        end
      end
    end
  end

  -- the font and size its text is to have where it says nothing: its format's (a plain paragraph's),
  -- else what most of its text has now (false: none of its own, its style's)
  -- (a new paragraph with no runs of its own: its text as the one next to it, in a table its row's
  -- and column's; what its font and size are counted from too)
  local inherited
  if new and not runs[1] and #chars > 0 then
    -- (a new table's first cells: as the text near it)
    inherited = para.cell and cell_props(buf, para) or neighbour_props(buf, para, para.cell and {} or block)
  end
  local fonts = {}
  local own_of, known = doc.fmt.run_font or function() return {} end, {}
  local function own(props) -- (once for each run's properties)
    known[props] = known[props] or own_of(doc.ctx, props)
    return known[props]
  end
  local in_table = para.cell and table_fonts(buf, para)
  for _, key in ipairs(markup.SPAN_KEYS) do
    if block.para then -- (none: its style's)
      fonts[key] = block.para[key] or false
    elseif cell_own and cell_own[key] then -- (a cell's own)
      fonts[key] = cell_own[key]
    elseif in_table and in_table[key] ~= nil and (not new or runs[1]) then -- (else its table's; a pasted row's: its column's, see cell_props)
      fonts[key] = in_table[key]
    else
      local weight, order, at, props_now = {}, {}, walker(runs), runs[1] and doc.run[runs[1][1]] or inherited
      for _, ch in ipairs(chars) do
        if ch.ch ~= "" and ch.ch ~= "\n" and not ch[key] and not (key == "color" and ch.link) then
          local run = at(ch)
          local p = run and doc.run[run[1]]
          props_now = p and not (doc.link_props[p] and not ch.link) and p or props_now
          local v = props_now and own(props_now)[key] or false
          if weight[v] == nil then
            order[#order + 1], weight[v] = v, 0
          end
          weight[v] = weight[v] + 1
        end
      end
      for _, v in ipairs(order) do
        if fonts[key] == nil or weight[v] > weight[fonts[key]] then
          fonts[key] = v
        end
      end
      fonts[key] = fonts[key] or false
    end
  end

  local out, e, parts = {}, 1, {} -- (parts: how many pieces of each deletion so far)
  local props = runs[1] and doc.run[runs[1][1]] -- text before the first run takes its properties
  props = props or inherited
  local plain_props = not (props and doc.link_props[props]) and props or inherited -- (the last not a link's)
  local has_inserts = false
  for _, m in ipairs(changes) do
    has_inserts = has_inserts or m[5].type == "insert"
  end
  local k = 0
  while true do
    k = k + 1
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
        -- (the first part, the change itself, only where the deletion starts: in a table, a cell after
        -- it on its row is another part, as the next row's cells are)
        local first = doc.change_first and doc.change_first[deletion[5]] -- (and only its first mark's: one deletion
        -- shown in several places, a cell's text each, has its text written once; see M.save)
        local starts_before = deletion[2] < para.first - 1 or deletion[2] == para.first - 1 and para.cell ~= nil and deletion[3] < (para.col or 0)
          or first ~= nil and (first[1] < deletion[2] or first[1] == deletion[2] and first[2] < deletion[3])
        parts[deletion[5]] = (parts[deletion[5]] or (starts_before and 1 or 0)) + 1
        last = { deletion = deletion[5], deleted = "", part = parts[deletion[5]] - 1, segs = {} }
        out[#out + 1] = last
      end
      last.deleted = last.deleted .. char
      -- (its formatting: the deleted text is written formatted, to come back as it was)
      local run = run_at(ch)
      local seg, fmt = last.segs[#last.segs], { props = run and doc.run[run[1]] or nil }
      for _, flag in ipairs(markup.FLAGS) do
        fmt[flag] = ch[flag] or false -- (false: as the formats say it)
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
    if ch.image and not (object and object.image and object.image.path == ch.image.path and not used_image[mark[1]]) then
      -- (its line typed again, or one above it: this paragraph's image of that file not yet taken,
      -- else one whose text was deleted anywhere, moved with dd and p)
      local found
      for _, m in ipairs(doc.ctx.entries[ch.image.path] and objects or {}) do
        local o = doc.object[m[1]]
        if not found and not used_image[m[1]] and o and o.visible and o.image and o.image.path == ch.image.path then
          found = m
        end
      end
      if found then
        mark, object = found, doc.object[found[1]]
        api.nvim_buf_set_extmark(buf, M.ns.object, ch.row, ch.col, {
          id = found[1], end_row = ch.end_row, end_col = ch.end_col, right_gravity = false, end_right_gravity = false,
        })
      else
        object = lost_field(ch) or object
      end
    end
    if ch.image and mark then
      used_image[mark[1]] = true
    end
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
      if not (object and object.field) then
        object = lost_field(ch) or object
      end
      assert(object and object.field, ("⟦%s⟧ isn't one of the document's fields: a field (a citation, a date...) can't be typed or copied here "
        .. "(u brings back one deleted; write \\⟦ for the character itself)"):format(ch.field.text))
      assert(object.field.text == ch.field.text, ("a field's text can't be edited (Zotero or Word would write it again): ⟦%s⟧ is ⟦%s⟧ "
        .. "in the document (u undoes the edit; delete the whole ⟦...⟧ to delete it)"):format(ch.field.text, object.field.text))
      out[#out + 1] = { object = object.node }
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
      -- (text after a link, typed at its end: as the text before the link, not as the link's run, its
      -- style's colour and underline)
      if props and not ch.link and doc.link_props[props] then
        props = plain_props
      elseif not ch.link then
        plain_props = props
      end
      local insert
      for _, m in ipairs(has_inserts and changes or {}) do
        if m[5].type == "insert" and covers(m, ch.row, ch.col) then
          insert = m[5]
        end
      end
      -- the same formatting as the text before it: one piece (its text is collected, joined below)
      local last = out[#out]
      assert(not ch.span_bad, ch.span_bad) -- (a span of text's attributes: right)
      -- (a link's text: its colour as it is, the link's, unless a span says one)
      local font, size, color = ch.font or fonts.font, ch.size or fonts.size, ch.color
      if color == nil and not ch.link then -- (false: none of its own, its style's: written so)
        color = fonts.color
      end
      local same = last and type(last.text) == "table" and last.props == (props or nil) and last.link == ch.link
        and last.insert == insert and last.font == font and last.size == size and last.color == color
      for _, flag in ipairs(markup.FLAGS) do
        same = same and last[flag] == (ch[flag] or false)
      end
      if same then
        last.text[#last.text + 1] = ch.ch
      else
        local piece = { text = { ch.ch }, props = props or nil, link = ch.link, insert = insert, font = font, size = size, color = color }
        for _, flag in ipairs(markup.FLAGS) do
          piece[flag] = ch[flag] or false -- (false: as the formats say it)
        end
        out[#out + 1] = piece
      end
      -- (the text after it, as long as nothing could make it different, taken at once: no marker
      -- there, no tracked change, the same formatting, and the same run of the document)
      if not has_inserts and #deletions == 0 then
        local text, nxt = out[#out].text, chars[k + 1]
        while nxt and nxt.ch ~= "" and nxt.ch ~= "\n" and nxt.ch ~= "￼" and not nxt.span_bad
          and not (events[e] and events[e][1] == k + 1)
          and nxt.link == ch.link and nxt.b == ch.b and nxt.i == ch.i and nxt.u == ch.u and nxt.s == ch.s
          and nxt.h == ch.h and nxt.sup == ch.sup and nxt.sub == ch.sub and nxt.font == ch.font and nxt.size == ch.size and nxt.color == ch.color do
          local r2 = run_at(nxt)
          if r2 and doc.run[r2[1]] ~= props then
            break
          end
          text[#text + 1] = nxt.ch
          k = k + 1
          nxt = chars[k + 1]
        end
      end
    end
    ::next_char::
  end
  for _, piece in ipairs(out) do
    if type(piece.text) == "table" then
      piece.text = table.concat(piece.text)
    end
  end
  -- an insertion (deletion) running past the end of the paragraph inserted (deleted) its break too:
  -- if it has all of the paragraph, from its start, or runs over the blank line after it (a line
  -- inserted, or deleted, at the end of a paragraph of several lines, the blank line after it as it
  -- was, is that line only: the paragraph's break stays)
  local stop = { para.last - 1, para.stop or #lines[#lines] }
  local start = { para.first - 1, para.col or 0 }
  local after = api.nvim_buf_get_lines(buf, para.last, para.last + 1, false)[1]
  local blank_after = not para.cell and after ~= nil and after:match("^%s*$") ~= nil
  for _, m in ipairs(changes) do
    local d = m[4]
    local before_end = m[2] < stop[1] or m[2] == stop[1] and m[3] <= stop[2]
    local past_end = d.end_row > stop[1] or d.end_row == stop[1] and d.end_col > stop[2]
    local whole = m[2] < start[1] or m[2] == start[1] and m[3] <= start[2]
    local over_blank = not blank_after or d.end_row > para.last or d.end_row == para.last and d.end_col > 0
    if m[5].type ~= "format" and before_end and past_end and (para.cell or whole or over_blank) then
      out[#out + 1] = { paragraph_mark = m[5] }
    end
  end
  return block, out
end

-- A fingerprint of a paragraph's pieces: equal fingerprints mean the paragraph's XML is still right.
-- Fonts and sizes, as shown: a paragraph's are what most of its text has (a plain paragraph's are
-- its format's: `shown`, on its format's line; another's are kept as they are, unsaid), and text in
-- another has them in its span ("[words]{font = Arial}"). A span's run_font, run_size: its text's
-- own (see the formats); block.font_base: what the paragraph's style gives.
local function fonts_of(block, spans, shown)
  local base = block.font_base or {}
  for _, key in ipairs(markup.SPAN_KEYS) do
    local inherit = block.font_inherit or base -- (what text that says nothing shows)
    local function value(s) -- (a link's colour is the link's: as its underline, not its text's)
      return s.text and not s.deletion and not (s.image or s.math or s.footnote or s.field) and not (key == "color" and s.link)
        and (s["run_" .. key] or inherit[key] or base[key]) or nil
    end
    local function same(a, b) -- (as measured: a font whatever its case)
      return pformat.measure(key, a) == pformat.measure(key, b)
    end
    local weight, order, main = {}, {}, nil -- (by measure: its most common spelling)
    for _, s in ipairs(spans) do
      local v = value(s)
      if v then
        local m = pformat.measure(key, v) or v
        if not weight[m] then
          order[#order + 1], weight[m] = m, { n = 0, spelling = {} }
        end
        weight[m].n = weight[m].n + #s.text
        weight[m].spelling[v] = (weight[m].spelling[v] or 0) + #s.text
      end
    end
    for _, m in ipairs(order) do
      if not main or weight[m].n > weight[main].n then
        main = m
      end
    end
    if main then
      local best
      for v, n in pairs(weight[main].spelling) do
        if not best or n > weight[main].spelling[best] or n == weight[main].spelling[best] and v < best then
          best = v
        end
      end
      main = best
    end
    if shown then
      block.para = block.para or {}
      -- (what it shows, wherever it comes from; unsaid if it is its style's: see the header)
      block.para[key] = main and not same(main, base[key]) and main or nil
      block.no_text = not main or nil -- (no text, so no font: whatever the paragraphs around it have)
    end
    for _, s in ipairs(spans) do
      local v = value(s)
      s[key] = v and not same(v, main) and v or nil
    end
  end
end

-- What a paragraph's fingerprint depends on besides its text and its marks: its format (from the
-- line above it), its box, its column's alignment.
-- What paragraphs are matched by (see match_lines): their text, and a table cell's its column too
-- (in a table, many cells read the same: an empty one is matched with one of its column).
local function match_key(text, col)
  return col and (text .. "\1" .. col) or text
end
local function block_keys(blocks)
  local out = {} -- (the paragraphs only: blocks also say what follows them, an unclosed box...)
  for k, b in ipairs(blocks) do
    out[k] = match_key(b.text, b.cell and b.cell.col)
  end
  return out
end
local function context_key(block)
  return table.concat({ block.format and pformat.key(block.format) or "", block.box and block.box.attrs or "",
    block.cell and tostring(block.cell.align) or "" }, "|")
end
-- (whether rows changed since the fingerprints were taken cover a paragraph)
local function touched(doc, block)
  -- (an empty paragraph, a blank line: whether it is one, or the line between two paragraphs, depends
  -- on the lines around it too: an edit next to it touches it)
  local first, last = block.first - 1, block.owns - 1
  if block.text == "" then
    first, last = (block.around and block.around[1] or block.first - 1) - 1, (block.around and block.around[2] or block.last + 1) - 1
  end
  for _, r in ipairs(doc.touched or {}) do
    if r[1] <= last and r[2] >= first then
      return true
    end
  end
  return false
end

local PIECE_KEYS = { "text", "props", "link", "insert", "font", "size", "color", "b", "i", "u", "s", "h", "sup", "sub", "object",
  "note", "image", "math", "footnote", "deletion", "deleted", "part", "segs", "line_break", "page_break", "insert_start",
  "insert_end", "comment_start", "comment_end", "paragraph_mark", "mark", "original", "visible" }
local PIECE_KNOWN = {}
for _, key in ipairs(PIECE_KEYS) do
  PIECE_KNOWN[key] = true
end
local function signature(block, list)
  local l = block.list
  local out = { table.concat({ tostring(block.heading), tostring(block.title), tostring(block.quote), tostring(block.align), tostring(block.caption), tostring(block.vspace), tostring(block.toc),
    block.para and pformat.key({ attrs = block.para }) or "",
    tostring(block.cell), l and table.concat({ l.kind, l.level, tostring(l.pattern), tostring(l.id), tostring(l.continues) }, ",") or "" }, "|") }
  for _, piece in ipairs(list) do
    -- (its keys in an order of their own: the usual ones as listed, any other after them, sorted)
    local keys = {}
    for _, key in ipairs(PIECE_KEYS) do
      if piece[key] ~= nil then
        keys[#keys + 1] = key
      end
    end
    local extra
    for key in pairs(piece) do
      if not PIECE_KNOWN[key] then
        extra = extra or {}
        extra[#extra + 1] = key
      end
    end
    if extra then
      table.sort(extra)
      vim.list_extend(keys, extra)
    end
    for _, key in ipairs(keys) do
      local v = piece[key]
      if key == "segs" then -- (the deleted text's formatting)
        v = table.concat(vim.tbl_map(function(seg)
          return tostring(seg.props) .. table.concat(vim.tbl_map(function(f) return tostring(seg[f]) end, markup.FLAGS)) .. seg.text
        end, v), "/")
      elseif type(v) == "table" then -- (identity, plus what can change inside comments, changes, images)
        local replies = table.concat(vim.tbl_map(function(r) return tostring(r) .. "|" .. (r.text or "") end, v.replies or {}), ",")
        v = table.concat({ tostring(v), replies, tostring(v.resolved), v.text or "", v.path or "", v.alt or "", v.path and floats.key(v) or "", v.tex or "", tostring(v.display) }, "/")
      end
      out[#out + 1] = key .. "=" .. tostring(v)
    end
    out[#out + 1] = ";"
  end
  return table.concat(out, " ")
end

-- ── Open ────────────────────────────────────────────────────────────────────

-- Tree-sitter's languages for documents: markdown's parsers under names of their own ("neoffice",
-- "neoffice_inline"), with markdown's queries but for bold, italics and strikethrough: markdown reads
-- a single "~" or "_" as marks ("~30 m ... ~100 m" struck through), neoffice doesn't, and shows them
-- as it reads them (M.decorate). Markdown's own highlighting, in other files, is as it was.
local languages
function M.languages()
  if languages ~= nil then
    return languages
  end
  languages = false
  local ts = vim.treesitter
  local md = api.nvim_get_runtime_file("parser/markdown.so", false)[1]
  local inline = api.nvim_get_runtime_file("parser/markdown_inline.so", false)[1]
  local function read(lang, name)
    local out = {}
    for _, file in ipairs(ts.query.get_files(lang, name)) do
      local f = io.open(file)
      if f then
        out[#out + 1] = f:read("*a")
        f:close()
      end
    end
    return table.concat(out, "\n")
  end
  local ok = md and inline and pcall(function()
    ts.language.add("neoffice", { path = md, symbol_name = "markdown" })
    ts.language.add("neoffice_inline", { path = inline, symbol_name = "markdown_inline" })
    ts.query.set("neoffice", "highlights", read("markdown", "highlights"))
    ts.query.set("neoffice", "injections", (read("markdown", "injections"):gsub('"markdown_inline"', '"neoffice_inline"')))
    local marks = read("markdown_inline", "highlights")
    for _, capture in ipairs({ "%(emphasis%)%s*@markup%.italic", "%(strong_emphasis%)%s*@markup%.strong", "%(strikethrough%)%s*@markup%.strikethrough" }) do
      marks = marks:gsub(capture, "")
    end
    ts.query.set("neoffice_inline", "highlights", marks)
    ts.query.set("neoffice_inline", "injections", read("markdown_inline", "injections"))
    ts.language.register("neoffice", "neoffice")
  end)
  languages = ok and true or false
  return languages
end

local function setup_buffer(buf)
  M.languages()
  vim.bo[buf].filetype = "neoffice"
  pcall(vim.treesitter.start, buf) -- markdown (see plugin/neoffice.lua): highlights the marks
  -- (underlines, superscripts... drawn as the text is decorated, see M.decorate_marks; by the syntax,
  -- syntax/neoffice.vim, when the marks are concealed: it hides them)
  vim.bo[buf].syntax = M.config.conceal and "neoffice" or ""
  -- Tab types a tab (a document's tab, shown as →), not spaces, whatever 'expandtab' and 'smarttab'
  -- say (a completion plugin that uses Tab hands it on to this, as blink.cmp does); >> and << move a
  -- list item by a level: two spaces, never a tab
  vim.keymap.set("i", "<Tab>", "<C-v><Tab>", { buffer = buf, desc = "neoffice: a tab" })
  local help = type(M.config.keymaps) == "table" and M.config.keymaps.help
  if help then -- (its help: the only key neoffice sets in a document)
    vim.keymap.set("n", help, function() M.help("neoffice") end, { buffer = buf, nowait = true, silent = true, desc = "neoffice: help" })
  end
  vim.bo[buf].expandtab, vim.bo[buf].shiftwidth = true, 2
  require("neoffice.lsp").start(buf)
  local function window()
    -- (all as :setlocal: for this buffer in this window, not the next buffer opened in it; wrapped
    -- lines from the first column, not under a tab that starts a paragraph)
    vim.opt_local.wrap, vim.opt_local.linebreak, vim.opt_local.breakindent = true, true, false
    if M.config.conceal then
      vim.opt_local.conceallevel = 2
    end
    if M.config.show_tabs then -- (in this window, for this buffer only: as :setlocal; drawn in the marks' gray by M.decorate)
      vim.opt_local.list = true
      vim.opt_local.listchars = "tab:→ "
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
      M.decorate_changed(buf)
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
  vim.schedule(function() -- (once the text shows: opening doesn't wait for the colours)
    if api.nvim_buf_is_valid(buf) then
      M.decorate_first(buf)
    end
  end)
end

-- Convert a file with LibreOffice (to = "docx", "doc:MS Word 97", ...) into `dir`; returns the new path.
-- (on_done: in the background, with a profile of its own, so as not to wait for it nor be in the
-- way of a conversion waited for; on_done(path or nil) once done)
local background = {} -- (the conversions waiting: one at a time, as they share a profile)
local function convert(path, to, dir, on_done)
  local profile = require("neoffice.office").profile(vim.fn.stdpath("cache") .. "/neoffice-libreoffice" .. (on_done and "-background" or ""))
  local cmd = { "soffice", profile, "--headless", "--norestore", "--convert-to", to, "--outdir", dir, path }
  local out = dir .. "/" .. vim.fn.fnamemodify(path, ":t:r") .. "." .. to:match("^%w+")
  if on_done then
    local function run()
      vim.system(cmd, {}, vim.schedule_wrap(function(r)
        table.remove(background, 1)
        on_done(r.code == 0 and vim.uv.fs_stat(out) and out or nil)
        if background[1] then
          background[1]()
        end
      end))
    end
    background[#background + 1] = run
    if #background == 1 then
      run()
    end
    return
  end
  local r = vim.system(cmd):wait()
  assert(r.code == 0 and vim.uv.fs_stat(out), "LibreOffice (soffice) could not convert " .. path)
  return out
end

-- .doc files as LibreOffice converts them to .docx, kept by the .doc's content (its SHA-256): one
-- opens at once the next time, in any session, while it is as it was (and one neoffice wrote, as
-- read back from what it wrote: see M.save). The last ones used only.
local DOC_CACHE_KEEP = 100
local function doc_cache()
  local dir = M.config.doc_cache
  if dir == false then
    return nil
  end
  return dir or vim.fs.joinpath(vim.fn.stdpath("cache"), "neoffice", "doc")
end
local function content_key(path)
  local f = io.open(path, "rb")
  if not f then
    return nil
  end
  local data = f:read("*a")
  f:close()
  return vim.fn.sha256(data)
end
local function cached_doc(key) -- (its .docx, if kept: marked as used)
  local dir = doc_cache()
  local file = dir and key and vim.fs.joinpath(dir, key .. ".docx")
  if file and vim.uv.fs_stat(file) then
    local now = os.time()
    pcall(vim.uv.fs_utime, file, now, now)
    return file
  end
end
local function keep_doc(key, docx) -- (and the least recently used go, beyond the last ones)
  local dir = doc_cache()
  if not (dir and key) then
    return
  end
  vim.fn.mkdir(dir, "p")
  local file = vim.fs.joinpath(dir, key .. ".docx")
  if not vim.uv.fs_copyfile(docx, file .. ".part") or not os.rename(file .. ".part", file) then
    os.remove(file .. ".part")
    return
  end
  local kept = {}
  for name, kind in vim.fs.dir(dir) do
    local stat = kind == "file" and name:match("%.docx$") and vim.uv.fs_stat(vim.fs.joinpath(dir, name))
    if stat then
      kept[#kept + 1] = { name = name, mtime = stat.mtime.sec }
    end
  end
  table.sort(kept, function(x, y) return x.mtime > y.mtime end)
  for k = DOC_CACHE_KEEP + 1, #kept do
    os.remove(vim.fs.joinpath(dir, kept[k].name))
  end
end

-- A copy of the plugin's empty document (templates/), to start a new one from.
local function template(ext)
  local here = vim.fs.dirname(debug.getinfo(1, "S").source:sub(2)) -- (lua/neoffice)
  local copy = vim.fn.tempname() .. "." .. ext
  assert(vim.uv.fs_copyfile(vim.fs.joinpath(here, "..", "..", "templates", "empty." .. ext), copy))
  return copy
end

-- How plain paragraphs' formats are shown (blocks: { align, para, no_text }, in order): for each,
-- { run = the format the "+" line above it starts (nil: none), own = what a line for it only says
-- (nil: none) }. Of the ways to say them, the one with the fewest lines, then the fewest characters:
-- at each paragraph the run goes on as it is, or becomes its format, or takes from it what the next
-- paragraph has too (an attribute at a time: a run's alignment isn't said again because the style
-- changes). (A paragraph of only images has the run's font, size and colour: none of its own.)
function M.format_lines(blocks)
  local function own_of(b, run)
    local attrs = b.para or {}
    if b.no_text then
      attrs = vim.tbl_extend("force", attrs, { font = run.attrs.font, size = run.attrs.size, color = run.attrs.color })
    end
    return { align = b.align, attrs = attrs }
  end
  local function value(f, key)
    if key == "align" then
      return pformat.measure("align", f.align or "left")
    end
    return pformat.measure(key, (f.attrs or {})[key])
  end
  local keys = vim.list_extend({ "align" }, vim.tbl_filter(function(x) return x ~= "align" end, pformat.KEYS))
  local function shared_with(run, own, nxt) -- (the run, with what of own the next one has too)
    local out = { align = run.align, attrs = vim.deepcopy(run.attrs or {}) }
    for _, key in ipairs(keys) do
      if value(own, key) ~= value(run, key) and nxt and value(nxt, key) == value(own, key) then
        if key == "align" then
          out.align = own.align
        else
          out.attrs[key] = own.attrs[key]
        end
      end
    end
    return out
  end
  local function cost(prev, run, own) -- (lines first, then characters)
    local c = 0
    if not pformat.same(prev, run) then
      c = c + 1 + (#markup.para_line_for(prev, run) + 1) / 10000 -- (and its "+")
    end
    if not pformat.same(run, own) then
      local align, attrs = pformat.diff(run, own)
      c = c + 1 + #markup.render_para_line({ align = align, attrs = attrs or {} }) / 10000
    end
    return c
  end
  local start = { run = { attrs = {} }, cost = 0 }
  local steps, states = {}, { start }
  for i, b in ipairs(blocks) do
    local nb = blocks[i + 1]
    local nxt = nb and { align = nb.align, attrs = nb.para or {} }
    local best = {}
    for _, st in ipairs(states) do
      local own = own_of(b, st.run)
      for _, run in ipairs({ st.run, own, shared_with(st.run, own, nxt) }) do
        local c = st.cost + cost(st.run, run, own)
        local key = pformat.key(run)
        if not best[key] or c < best[key].cost - 1e-9 then
          best[key] = { run = run, cost = c, from = st, own = own }
        end
      end
    end
    states = vim.tbl_values(best)
    table.sort(states, function(x, y) return x.cost < y.cost - 1e-9 or math.abs(x.cost - y.cost) <= 1e-9 and pformat.key(x.run) < pformat.key(y.run) end)
    while #states > 12 do -- (the best few ways so far: enough, as formats change little)
      table.remove(states)
    end
    steps[i] = states
  end
  local out, st = {}, states[1]
  for i = #blocks, 1, -1 do
    local prev = st.from.run
    local line = { run = not pformat.same(prev, st.run) and st.run or nil }
    if not pformat.same(st.run, st.own) then
      local align, attrs = pformat.diff(st.run, st.own)
      line.own = { align = align, attrs = attrs or {} }
    end
    out[i] = line
    st = st.from
  end
  return out
end

-- Which of n paragraphs (or blocks) are first ones: the first in the document, or right after a
-- heading, title or quote, as pandoc's First Paragraph; place(k): "break" (a heading, title, quote),
-- "other", or nil for one that doesn't count (an empty one).
function M.first_paragraphs(n, place)
  local out, before = {}, "break"
  for k = 1, n do
    local p = place(k)
    if p then
      out[k] = before == "break"
      before = p
    end
  end
  return out
end

-- The styles at the top: those of `used` (ids, in the order the document has them) and those they
-- are based on, each under the one it is based on (depth), saying what it says itself (attrs: what
-- isn't there it has from the one above it); the first of each tree, all it gives (view).
-- { { id, name, depth, base, view, attrs } } in order.
function M.style_entries(doc, used)
  local ctx, fmt = doc.ctx, doc.fmt
  local function real(id) -- (a paragraph style the document has)
    return id and fmt.style_id(ctx, id) == id and pcall(fmt.style_view, ctx, id)
  end
  local base_of, seen, order = {}, {}, {}
  local function add(id, depth)
    if seen[id] or depth > 10 then
      return
    end
    seen[id] = true
    local base = fmt.base_style and fmt.base_style(ctx, id)
    if base and base ~= id and real(base) then
      base_of[id] = base
      add(base, depth + 1)
    end
    order[#order + 1] = id
  end
  for _, id in ipairs(used) do
    add(id, 0)
  end
  local children, rank = {}, {}
  for k, id in ipairs(order) do
    rank[id] = k
    local b = base_of[id]
    if b then
      children[b] = children[b] or {}
      table.insert(children[b], id)
    end
  end
  local out, names, views = {}, {}, {}
  local function view(id)
    views[id] = views[id] or fmt.style_view(ctx, id)
    return views[id]
  end
  local function walk(id, depth)
    if depth > 10 then
      return
    end
    local entry = { id = id, name = fmt.style_name(ctx, id), depth = depth, base = base_of[id], view = view(id) }
    entry.attrs = base_of[id] and fmt.own_view(ctx, id) or entry.view
    if names[entry.name:lower()] then -- (two styles of one name, as Word allows: this one with its id)
      entry.name = ("%s (%s)"):format(entry.name, id)
    end
    names[entry.name:lower()] = true
    out[#out + 1] = entry
    local list = children[id] or {}
    table.sort(list, function(a, b) return rank[a] < rank[b] end)
    for _, c in ipairs(list) do
      walk(c, depth + 1)
    end
  end
  for _, id in ipairs(order) do
    if not base_of[id] then
      walk(id, 0)
    end
  end
  return out
end

-- The styles' lines at the top, each as its style now is (after a save that changed them, or
-- :Neoffice styles reset): what they show, and what the next save compares them with.
function M.refresh_header(buf, doc)
  local lines = api.nvim_buf_get_lines(buf, 0, -1, false)
  local header = doc.header and markup.header(lines)
  if not header then
    return
  end
  local found = {}
  for _, st in ipairs(header.styles) do
    local known = st.name and doc.header[st.name:lower()]
    local id = known and known.id or st.name and (doc.fmt.style_id(doc.ctx, st.name) or doc.fmt.style_id(doc.ctx, st.name:match("%(([^()]+)%)%s*$") or ""))
    if id then
      found[#found + 1] = { st = st, id = id }
    end
  end
  local by_id = {}
  for _, e in ipairs(M.style_entries(doc, vim.tbl_map(function(x) return x.id end, found))) do
    by_id[e.id] = e
  end
  local shown = {}
  for _, x in ipairs(found) do
    local e = vim.deepcopy(by_id[x.id])
    e.name = x.st.name
    shown[#shown + 1] = { e = e, row = x.st.row }
  end
  local width = markup.style_width(vim.tbl_map(function(s) return s.e end, shown))
  for _, s in ipairs(shown) do
    doc.header[s.e.name:lower()] = s.e
    local line = markup.render_style_line(s.e, width)
    if lines[s.row] ~= line then
      api.nvim_buf_set_lines(buf, s.row - 1, s.row, false, { line })
    end
  end
end

-- A paragraph style as {style = ...} names it: its name, or "Name (Id)" for one of two of a name.
function M.style_label(doc, id)
  local name = doc.fmt.style_name(doc.ctx, id)
  if doc.fmt.style_id(doc.ctx, name) ~= id then
    name = ("%s (%s)"):format(name, id)
  end
  return name
end

-- :Neoffice styles reset: the document's styles as a fresh document's (the plugin's empty one, see
-- fmt.reset_styles), and the styles at the top as they now are. Written with the next :w.
function M.reset_styles(buf)
  local doc = M.docs[buf]
  if not doc then
    return vim.notify("neoffice: Not a document buffer", vim.log.levels.WARN)
  elseif not doc.fmt.reset_styles then
    return vim.notify("neoffice: This document's styles can't be reset", vim.log.levels.WARN)
  end
  local ext = doc.fmt.MAIN == "content.xml" and "odt" or "docx"
  local here = vim.fs.dirname(debug.getinfo(1, "S").source:sub(2)) -- (lua/neoffice)
  local ok, err = pcall(doc.fmt.reset_styles, doc.ctx, vim.fs.joinpath(here, "..", "..", "templates", "empty." .. ext))
  if not ok then
    return vim.notify("neoffice: The styles couldn't be reset: " .. tostring(err):gsub("^[^\n]-%.lua:%d+: ", ""), vim.log.levels.ERROR)
  end
  -- (the lines at the top: each style's, as it now is; what the next save compares them with too)
  M.refresh_header(buf, doc)
  for _, entry in pairs(doc.header or {}) do -- (those not at the top any more: as they now are too)
    entry.view = doc.fmt.style_view(doc.ctx, entry.id)
    entry.attrs = entry.base and doc.fmt.own_view(doc.ctx, entry.id) or entry.view
  end
  vim.bo[buf].modified = true -- (to be written, whether the lines changed or not)
  M.decorate(buf)
  vim.notify("neoffice: Styles reset, as a new document's: :w writes them")
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
  -- (an insertion's or a comment's start or end among them: stepped over, and kept where it is, as
  -- when the caption was typed with changes tracked)
  local function marker(s) -- (where a deletion starts: not, what follows is deleted, no caption's)
    return s.insert_end or s.insert_start and s.insert_start.type ~= "delete" or s.comment_start or s.comment_end
  end
  local function plain(s) -- (text, not deleted)
    return s.text and not s.deletion and not s.object
  end
  for _, para in ipairs(paras) do
    local block, spans = para.block, para.spans
    if not (block.heading or block.title or block.quote or block.list) then
      -- (deleted text before it, a paragraph deleted and joined to it, changes tracked: stepped over,
      -- kept where it is)
      local k = 1
      while spans[k] and (spans[k].deletion or spans[k].insert_start and spans[k].insert_start.type == "delete") do
        if spans[k].deletion then
          k = k + 1
        else
          local ch = spans[k].insert_start
          k = k + 1
          while spans[k] and spans[k].insert_end ~= ch do
            k = k + 1
          end
          k = k + 1
        end
      end
      local skipped = k - 1
      -- the text before the first thing that isn't text, and that thing
      local lead = ""
      while spans[k] and (marker(spans[k]) or plain(spans[k])) do
        lead, k = lead .. (spans[k].text or ""), k + 1
      end
      local word, rest = lead:match("^%s*(%S+)%s+(.*)$")
      local kind = word and markup.caption_kind(word)
      local field = spans[k] and spans[k].field
      local cut -- (how many bytes of text to take away from the start, the field included or not)
      if kind and rest == "" and field and field.seq and markup.caption_kind(field.seq) == kind then
        -- "Table " + its number + the separator, in the text after
        local after, j = "", k + 1
        while spans[j] and (marker(spans[j]) or plain(spans[j])) do
          after, j = after .. (spans[j].text or ""), j + 1
        end
        local sep = separator(after)
        block.caption, block.caption_sep = kind, sep ~= "" and sep or nil
        block.caption_label, block.caption_seq = word, field.seq
        for i = k, skipped + 1, -1 do -- (the word and the number; the markers stay)
          if not marker(spans[i]) then
            table.remove(spans, i)
          end
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
      local i = skipped + 1
      while cut and cut > 0 and spans[i] and (marker(spans[i]) or spans[i].text and not spans[i].object) do -- (from the text spans)
        if marker(spans[i]) then
          i = i + 1
        else
          local n = math.min(cut, #spans[i].text)
          spans[i].text, cut = spans[i].text:sub(n + 1), cut - n
          if spans[i].text == "" then
            table.remove(spans, i)
          end
        end
      end
    end
  end
end

-- A picture's height in the document (cm), from its element.
local function image_height(s)
  local o = s.object
  local extent = o and xml.find(o, "wp:extent")
  if extent then
    return (tonumber(xml.attr(extent, "cy") or "") or 0) / 360000
  end
  return o and markup.cm(xml.attr(o, "svg:height") or "") or nil
end

-- A floating image (anchored to its paragraph: see floats.lua) shows at the start of its
-- paragraph, wherever its anchor is in the text (its place is measured from the paragraph, so it
-- doesn't move); written, it goes there. Its paragraph's floating images all placed below the
-- middle of its pictures (a figure with a picture on top, others under it): at its end, as they
-- show. (Not one inside a tracked change or a comment's range.)
local function hoist_floats(paras, page)
  for _, para in ipairs(paras) do
    local depth, moved, all = 0, {}, {}
    local inline, below = 0, true
    for k, s in ipairs(para.spans) do
      depth = depth + ((s.insert_start or s.comment_start) and 1 or 0) - ((s.insert_end or s.comment_end) and 1 or 0)
      if s.image and floats.floating(s.image) then
        all[#all + 1] = k
        if depth == 0 then
          moved[#moved + 1] = k
        end
      elseif s.image then
        inline = math.max(inline, image_height(s) or 0)
      end
    end
    for _, k in ipairs(all) do
      local top = page and floats.top(para.spans[k].image, page)
      below = below and (top == true or type(top) == "number" and top >= inline / 2)
    end
    if inline > 0 and below and #moved == #all and #moved > 0 then
      local list = {}
      for n = #moved, 1, -1 do
        table.insert(list, 1, table.remove(para.spans, moved[n]))
      end
      -- (at the end, from left to right, as they show side by side)
      local SIDE = { left = 0, center = 1, right = 2 }
      local function across(s)
        local x = s.image.x and floats.signed_cm(s.image.x)
        return x or (SIDE[s.image.float or "left"] or 0) * 100
      end
      for n, s in ipairs(list) do
        s._order = n
      end
      table.sort(list, function(a, b)
        local xa, xb = across(a), across(b)
        if xa ~= xb then
          return xa < xb
        end
        return a._order < b._order
      end)
      for _, s in ipairs(list) do
        s._order = nil
        para.spans[#para.spans + 1] = s
      end
    else
      moved = vim.tbl_filter(function(k) return k > 1 end, moved) -- (one first already: where it is)
      for n = #moved, 1, -1 do
        table.insert(para.spans, 1, table.remove(para.spans, moved[n] + (#moved - n)))
      end
    end
  end
end

-- neoffice works on its own copy of the document (the base: a .docx for .doc files, the template
-- for a new one), and copies it to the file when writing. So saving never depends on the file
-- staying where it was, and a file deleted meanwhile is simply written again.
-- Whether the file seems open in Word or LibreOffice: their lock files beside it ("~$name.docx",
-- with the name's first two letters dropped when long, "~$" before it when short; LibreOffice's
-- ".~lock.name.docx#"). Their save would undo what is written here, and the other way round.
local function open_elsewhere(path)
  local dir, name = vim.fs.dirname(path), vim.fs.basename(path)
  for _, lock in ipairs({ "~$" .. name:sub(3), "~$" .. name, ".~lock." .. name .. "#" }) do
    if vim.uv.fs_stat(vim.fs.joinpath(dir, lock)) then
      return lock
    end
  end
end
local function warn_elsewhere(path, doing)
  local lock = open_elsewhere(path)
  if lock then
    vim.notify(("neoffice: %s seems open in %s too (%s is beside it): a save there would undo what you %s here, and the other way round"):format(
      vim.fs.basename(path), lock:find("^%.~lock") and "LibreOffice" or "Word", lock, doing), vim.log.levels.WARN)
  end
end

-- What opening should tell, for a document neoffice edits but can't keep as it is in every way: a
-- signed one (a change saved makes its signature invalid), one protected in Word (neoffice doesn't
-- keep you to the parts it lets you edit: Word may flag changes elsewhere).
local function warn_kept(base, path)
  local entries = zip.list(base) or {}
  local signed = false
  for name in pairs(entries) do
    signed = signed or name:find("^_xmlsignatures/") ~= nil or name == "META-INF/documentsignatures.xml"
  end
  local said = {}
  if signed then
    said[#said + 1] = "it is digitally signed: a change you save makes its signature invalid"
  end
  local settings = entries["word/settings.xml"] and zip.read(base, "word/settings.xml") or ""
  local protection = settings:match("<w:documentProtection[^>]*>")
  if protection and protection:find('w:enforcement="[1tTo]') then -- ("1", "true", "on")
    local kind = protection:match('w:edit="(%w+)"') or "?"
    said[#said + 1] = ("it is protected in Word (editing restricted: %s), which neoffice doesn't keep you to"):format(kind)
  end
  if #said > 0 then
    vim.notify(("neoffice: %s: %s"):format(vim.fs.basename(path), table.concat(said, "; ")), vim.log.levels.WARN)
  end
end

local function open(buf, path)
  local ext = path:match("%.(%w+)$"):lower()
  local new = not vim.uv.fs_stat(path) -- a new document: created on the first :w
  local base
  if ext == "doc" then -- (converted by LibreOffice: once, while the file stays as it was, see doc_cache)
    if new then
      base = template("docx")
    else
      local key = content_key(path)
      local docx = cached_doc(key)
      if not docx then
        docx = convert(path, "docx", vim.fn.tempname())
        keep_doc(key, docx)
      end
      base = vim.fn.tempname() .. ".docx"
      assert(vim.uv.fs_copyfile(docx, base))
    end
  elseif new then
    base = template(ext)
  else
    base = vim.fn.tempname() .. "." .. ext
    assert(vim.uv.fs_copyfile(path, base))
  end
  local fmt = require("neoffice." .. (ext == "doc" and "docx" or ext))
  local ctx, paras = fmt.load(base)
  ctx.integrity = fmt.integrity and fmt.integrity(ctx) or {} -- (as opened: see M.save)
  ctx.bookmarks = M.paired_bookmarks(ctx.root) -- (those with both ends: kept so, see pair_bookmarks)
  if fmt.unique_drawings then
    fmt.unique_drawings(ctx)
  end
  if not new then
    warn_elsewhere(path, "write")
    warn_kept(base, path)
  end
  if (ctx.nested_tables or 0) > 0 then
    vim.notify("neoffice: a table inside a table's cell isn't shown; it is kept as it is", vim.log.levels.WARN)
  end
  find_captions(paras)
  hoist_floats(paras, fmt.page and fmt.page(ctx))
  ctx.dirty, ctx.files = {}, {} -- parts to write on save (name -> true, or content for new files)
  local doc = { path = path, base = base, is_doc = ext == "doc", fmt = fmt, ctx = ctx, paras = {}, tocs = {}, run = {}, object = {}, comment = {}, change = {}, caption_seps = {}, caption_as = {}, boxes = {}, link_props = setmetatable({}, { __mode = "k" }) }
  doc.stamp, doc.synced = stamp(path), not new -- (synced: the file has the base's content)
  doc.tracking = M.config.track_changes
  if doc.tracking == nil then
    doc.tracking = ctx.tracking or false
  end
  M.docs[buf] = doc
  for _, ns in pairs(M.ns) do
    api.nvim_buf_clear_namespace(buf, ns, 0, -1)
  end

  -- (numbered list items show their numbering's pattern, "1." "1.1." "a)": the document counts)

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
    -- (a cell's last paragraph's break inserted or deleted, a row inserted or deleted: not shown as a
    -- change of the "|" after it, or of its line's end, which accepting or rejecting would cut: a
    -- cell's end stays where it is)
    for k, para in ipairs(paras) do
      local cell, nxt = cells[k], cells[k + 1]
      local last = para.spans[#para.spans]
      if cell and not (nxt and nxt.tc == cell.tc) and last and last.insert_start and not last.text then
        table.remove(para.spans)
        local first = paras[k + 1] and paras[k + 1].spans[1]
        if first and first.insert_end == last.insert_start then
          table.remove(paras[k + 1].spans, 1)
        end
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
  -- Paragraphs' formats (see paras.lua): plain paragraphs one after another with the same one go
  -- under one line of it (before the first of them); one that differs says how, on its own.
  local format_before, own_before = {}, {} -- (a format's line before a paragraph: for those from it on, "+"; for it only)
  do
    local function content(spans)
      return vim.iter(spans):any(function(s)
        return s.text and s.text ~= "" or s.image or s.math or s.footnote or s.field or s.deletion or (s.object and s.visible)
      end)
    end
    local candidates = {}
    for k, para in ipairs(paras) do -- (fonts first: they are part of a paragraph's format)
      fonts_of(para.block, para.spans, plain(para.block) and not cells[k] and not para.box)
    end
    -- (its style, where it isn't the one plain paragraphs in its place have: "style = ...". Two places:
    -- first ones (first in the document, or after a heading, title or quote) and others; the style most of
    -- each have is the one there, doc.first_style and doc.plain_style; see M.first_paragraphs)
    if doc.fmt.style_of then
      local is_plain = {}
      for k, para in ipairs(paras) do
        is_plain[k] = not cells[k] and not para.box and plain(para.block) and para.node and content(para.spans)
      end
      local first = M.first_paragraphs(#paras, function(k)
        local para = paras[k]
        if cells[k] or para.box then
          return "other"
        elseif not content(para.spans) and plain(para.block) then
          return nil -- (an empty one: as if it weren't there)
        end
        local b = para.block
        return (b.heading or b.title or b.quote) and "break" or "other"
      end)
      local counts, order = { first = {}, other = {} }, { first = {}, other = {} }
      for k, para in ipairs(paras) do
        local id = is_plain[k] and doc.fmt.style_of(ctx, para.node)
        if type(id) == "string" then
          local where = first[k] and "first" or "other"
          if not counts[where][id] then
            table.insert(order[where], id)
          end
          counts[where][id] = (counts[where][id] or 0) + 1
        end
      end
      local default = doc.fmt.style_of(ctx, xml.elem(doc.fmt.MAIN == "content.xml" and "text:p" or "w:p"))
      local function most(where, fallback) -- (the one most have; of as many, the fallback, else the first seen)
        local best, n = fallback, counts[where][fallback] or 0
        for _, id in ipairs(order[where]) do
          if counts[where][id] > n then
            best, n = id, counts[where][id]
          end
        end
        return best
      end
      doc.plain_style = most("other", #order.other == 0 and most("first", default) or default)
      doc.first_style = most("first", doc.plain_style)
      doc.style_labels = {} -- (label shown -> its id: what a save reads back, even for a style not defined)
      doc.style_shown = {} -- (the paragraphs it was shown for: see M.save)
      if type(doc.plain_style) == "string" and type(doc.first_style) == "string" then
        for k, para in ipairs(paras) do
          local id = is_plain[k] and doc.fmt.style_of(ctx, para.node)
          if type(id) == "string" and id ~= (first[k] and doc.first_style or doc.plain_style) then
            local label = M.style_label(doc, id)
            doc.style_labels[label:lower()], doc.style_shown[k] = id, true
            para.block.para = para.block.para or {}
            para.block.para.style = label
          end
        end
      else
        doc.plain_style, doc.first_style = nil, nil
      end
    end
    for k, para in ipairs(paras) do
      local b = para.block
      if not cells[k] and not para.box and plain(b) then
        if content(para.spans) then
          candidates[#candidates + 1] = k
        else -- (an empty one: its run's)
          b.align, b.para = nil, nil
        end
      else
        b.para = nil -- (not a plain paragraph's: not neoffice's to change)
      end
    end
    -- which lines show the paragraphs' formats: those with the fewest lines, then the fewest
    -- characters, of the ways to say them (see M.format_lines)
    local plan = M.format_lines(vim.tbl_map(function(k) return paras[k].block end, candidates))
    for i, k in ipairs(candidates) do
      format_before[k], own_before[k] = plan[i].run, plan[i].own
      paras[k].block.align, paras[k].block.para = nil, nil
    end
  end
  -- a cell's font and size, when they aren't what most of its table's text has: "{size = 7pt} ..."
  do
    local function weigh(list, out)
      for _, para in ipairs(list) do
        for _, s in ipairs(para.spans) do
          if s.text and s.text ~= "" and not s.deletion and not (s.image or s.math or s.footnote or s.field) then
            for _, key in ipairs(s.link and { "font", "size" } or markup.SPAN_KEYS) do
              local v = s["run_" .. key] or false
              local m = v and (pformat.measure(key, v) or v) or false
              if key == "color" and m == "000000" then -- (black: as none of its own, automatic)
                v, m = false, false
              end
              out[key] = out[key] or { order = {}, n = {}, v = {} }
              if not out[key].n[m] then
                table.insert(out[key].order, m)
                out[key].n[m], out[key].v[m] = 0, v
              end
              out[key].n[m] = out[key].n[m] + #s.text
            end
          end
        end
      end
      local main = {}
      for key, w in pairs(out) do
        local best
        for _, m in ipairs(w.order) do
          if best == nil or w.n[m] > w.n[best] then
            best = m
          end
        end
        main[key] = { m = best, v = w.v[best] }
      end
      return main
    end
    local of_table = {}
    for k, para in ipairs(paras) do
      if cells[k] then
        of_table[cells[k].tbl] = of_table[cells[k].tbl] or {}
        table.insert(of_table[cells[k].tbl], para)
      end
    end
    local mains = {}
    for tbl, list in pairs(of_table) do
      mains[tbl] = weigh(list, {})
    end
    for k, para in ipairs(paras) do
      if cells[k] then
        local tmain, own = mains[cells[k].tbl], weigh({ para }, {})
        local b = para.block
        for _, key in ipairs(markup.SPAN_KEYS) do
          local t, c = tmain[key], own[key]
          if t and c and c.m ~= t.m then
            local base = b.font_inherit or b.font_base or {}
            local shown = c.v or base[key]
            if shown then
              b.para = b.para or {}
              b.para[key] = shown
            end
          end
        end
      end
    end
  end
  for k, para in ipairs(paras) do
    para.block.cell = cells[k] and true or nil
    if cells[k] then -- (spaces at the edges of a cell's text: not shown, as the table can't show them)
      -- (those before anything shown, an image..., or after it: a space after an image is between)
      local function trim(sp, pattern) -- (true: something shown, stop)
        if sp.text and not sp.deletion and not sp.object then
          sp.text = sp.text:gsub(pattern, "")
          return sp.text ~= ""
        end
        return sp.deletion or sp.image or sp.math or sp.footnote or sp.field or sp.line_break or sp.page_break
          or (sp.object and sp.visible) or false
      end
      for _, sp in ipairs(para.spans) do
        if trim(sp, "^%s+") then
          break
        end
      end
      for k2 = #para.spans, 1, -1 do
        if trim(para.spans[k2], "%s+$") then
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
      if s.deletion then -- (shown in the text, struck through; a paragraph break as ¶, see markup.DELETED_BREAK)
        s.text = s.deletion.text:gsub("\n", markup.DELETED_BREAK)
        s.deletion.no_text = s.text == "" or nil -- (a deleted image, tab...: a point, kept as it is)
      elseif s.text then
        s.text = s.text:gsub("\n", " ")
      elseif s.object and s.visible and not (s.image or s.math or s.footnote or s.field) then
        s.text, s.raw = "￼", true
      end
    end
    -- (an empty paragraph of exactly a height, nothing before or after it: vertical space, \vspace)
    local vspace = not cells[k] and not para.box and para.node and doc.fmt.vspace_of and plain(para.block)
      and markup.render(para.spans, {}) == "" and doc.fmt.vspace_of(doc.ctx, para.node)
    if vspace then
      para.block.vspace, para.block.para, para.block.align = markup.vspace_shown(vspace), nil, nil
    end
    texts[k], cols[k] = markup.render(para.spans, para.block)
    if cells[k] and texts[k]:find("^%s") or cells[k] and texts[k]:find("%s$") then
      -- (a cell's text as its row reads it: no spaces around it, even before a break or a tab not
      -- shown; the paragraph keeps them, written as it was unless it is edited)
      local lead = #texts[k]:match("^%s*")
      local trimmed = texts[k]:sub(lead + 1):gsub("%s+$", "")
      for _, c in pairs(cols[k]) do
        c[1] = math.max(0, math.min(c[1] - lead, #trimmed))
        c[2] = math.max(0, math.min(c[2] - lead, #trimmed))
      end
      texts[k] = trimmed
    end
    last_text = (texts[k] ~= "" or cells[k]) and k or last_text
    doc.paras[k] = para.node
    doc.tocs[k] = para.toc and para.toc.nodes or nil -- (a table of contents...: all it is, kept as it is)
    doc.caption_seps[k] = para.block.caption_sep
    doc.caption_as[k] = para.block.caption and { label = para.block.caption_label, seq = para.block.caption_seq, number = para.block.caption_number } or nil
    if para.box then -- (the document's boxes, and how they are placed: see M.save)
      doc.boxes[para.box.node] = floats.key(para.box.attrs)
    end
  end
  doc.texts = texts
  local heading_key = {} -- (the headings, as a table of contents lists them: see M.save)
  for k, t in ipairs(texts) do
    local kind = markup.block(t)
    if kind.heading and not cells[k] and not paras[k].box then
      local _, chars = markup.parse(t)
      heading_key[#heading_key + 1] = kind.heading .. " " .. vim.trim(table.concat(vim.tbl_map(function(c) return c.ch end, chars)))
    end
  end
  doc.heading_key = table.concat(heading_key, "\n")
  doc.keys = {}
  for k, text in ipairs(texts) do
    doc.keys[k] = match_key(text, cells[k] and column[k])
  end

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
  local shown_format = { attrs = {} } -- (the format before each format's line: see M.para_line_for)
  -- the styles its paragraphs have, at its top (only in a document with text): its default one
  -- first, then as they come
  -- its sections' page settings (see pages.lua): the first's at its top, each other's where it starts
  local pages = require("neoffice.pages")
  local section_at = {} -- (paragraph index -> section)
  doc.sections, doc.section_mark = nil, {}
  if M.config.page ~= false and doc.fmt.sections and last_text > 0 then
    doc.sections = doc.fmt.sections(ctx, paras)
    for j, sec in ipairs(doc.sections) do
      sec.view = pages.view(sec.props)
      if j > 1 then
        section_at[sec.start] = j
      end
    end
  end
  local page_line = doc.sections and pages.render(doc.sections[1].view, true) or nil
  doc.header = nil
  if page_line and not (M.config.styles ~= false and doc.fmt.style_of) then
    vim.list_extend(lines, markup.render_header(nil, page_line))
    lines[#lines + 1] = ""
  end
  if M.config.styles ~= false and doc.fmt.style_of and last_text > 0 then
    local seen, order = {}, {}
    local function add(id)
      if id and not seen[id] then
        seen[id], order[#order + 1] = true, id
      end
    end
    for _, para in ipairs(paras) do
      add(para.node and not para.block.toc and doc.fmt.style_of(ctx, para.node))
    end
    -- (and those of the tables of contents' paragraphs: their title's, their entries', to change how
    -- they look, though they aren't shown)
    local PARA_TAGS = doc.fmt.MAIN == "content.xml" and { ["text:p"] = true, ["text:h"] = true } or { ["w:p"] = true }
    for _, para in ipairs(paras) do
      for _, node in ipairs(para.toc and para.toc.nodes or {}) do
        local list = PARA_TAGS[node.tag] and { node } or xml.collect(node, PARA_TAGS)
        for _, p in ipairs(list) do
          add(doc.fmt.style_of(ctx, p))
        end
      end
    end
    local default = doc.fmt.style_of(ctx, xml.elem(doc.fmt.MAIN == "content.xml" and "text:p" or "w:p"))
    table.sort(order, function(a, b) return (a == default and 0 or 1) < (b == default and 0 or 1) end)
    local shown = M.style_entries(doc, order)
    doc.header = {}
    for _, entry in ipairs(shown) do
      doc.header[entry.name:lower()] = entry
    end
    vim.list_extend(lines, markup.render_header(shown, page_line))
    lines[#lines + 1] = ""
  end
  local section_rows = {}
  for k, para in ipairs(paras) do
    local r, at
    local cell, before, after = cells[k], cells[k - 1], cells[k + 1]
    local j = section_at[k]
    if j then -- (a section starts: its line, saying what changes)
      if lines[#lines] and lines[#lines] ~= "" then
        lines[#lines + 1] = ""
      end
      lines[#lines + 1] = "\\newpage {" .. pages.render(pages.delta(doc.sections[j - 1].view, doc.sections[j].view)) .. "}"
      section_rows[#section_rows + 1] = { #lines - 1, j }
      lines[#lines + 1] = ""
    end
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
      if format_before[k] then -- (the format of the paragraphs from here on: what changes, see paras.lua)
        local f = format_before[k]
        lines[#lines + 1] = markup.para_line_for(shown_format, f) .. "+"
        shown_format = f
      end
      if own_before[k] then -- (and its own, for it only)
        lines[#lines + 1] = markup.render_para_line(own_before[k])
      end
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
      local prefix = markup.prefix(para.block) -- (its marks: on the text's line, "<=> ...")
      local lead
      if rest and not rest:match("^%s*$") and not rest:match("^%s*\n") and not next((markup.block(rest))) and not rest:match("^|")
        and not markup.box_line(rest) and not markup.table_options(rest) then
        local again, start = markup.block(prefix .. rest) -- (the marks read back as they are)
        if prefix == "" then
          parts[1], joins[1], text = text:sub(1, split), 0, rest
        elseif start == #prefix + 1 and markup.prefix(again) == prefix then
          parts[1], joins[1], text = text:sub(#prefix + 1, split), 0, prefix .. rest
          lead = { prefix = #prefix, floats = split - #prefix }
        end
      end
      -- (or last, on a line of their own after the text: placed under it, see hoist_floats)
      local trailing = 0
      while floating == 0 and para.spans[#para.spans - trailing] and para.spans[#para.spans - trailing].image
        and floats.floating(para.spans[#para.spans - trailing].image) do
        trailing = trailing + 1
      end
      -- (not a list item's: the next item could follow it, and its line would read as that one's)
      local tail = trailing > 0 and trailing < #para.spans and not para.block.list and cols[k][#para.spans - trailing + 1]
      local cut = tail and tail[1]
      local head = cut and text:sub(1, cut)
      if head and head ~= "" and not head:match("[\n\\]$") and not head:match("^%s*$") and markup.floats_only(text:sub(cut + 1)) then
        text = head
      else
        cut = nil
      end
      for _, line in ipairs(vim.split(text, "\n", { plain = true })) do
        parts[#parts + 1] = line
      end
      if cut then
        joins[#parts] = 0
        parts[#parts + 1] = texts[k]:sub(cut + 1)
      end
      local starts, offset = {}, 0
      for i, line in ipairs(parts) do
        lines[#lines + 1], starts[#starts + 1], offset = line, offset, offset + #line + (joins[i] or 1)
      end
      at = function(col, ends) -- the position of a byte column of the paragraph's text
        col = M.buffer_offset({ lead = lead }, col, ends)
        local row = #starts
        -- (what ends where a line joined with nothing ends: on that line)
        while starts[row] > col or ends and row > 1 and starts[row] == col and joins[row - 1] == 0 do
          row = row - 1
        end
        -- (never past its line's end: after a line break that ends the paragraph, shown as the "\" that
        -- ends its line, with no line after it: a tracked one, or a range that ends after it)
        return r + row - 1, math.min(col - starts[row], #(lines[r + row] or ""))
      end
      local next_list = paras[k + 1] and paras[k + 1].block.list
      local tight = para.block.list and next_list and (next_list.group or next_list.id) == (para.block.list.group or para.block.list.id)
        and texts[k + 1] ~= ""
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
        if s.link and s.props then -- (a link's run: text typed after it isn't the link's, see pieces)
          doc.link_props[s.props] = true
        end
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
  for _, sr in ipairs(section_rows) do
    doc.section_mark[api.nvim_buf_set_extmark(buf, M.ns.section, sr[1], 0, { right_gravity = false })] = sr[2]
  end
  local alone = {}
  for _, r in ipairs(empty_rows) do
    api.nvim_buf_set_extmark(buf, M.ns.empty, r, 0, { invalidate = true, right_gravity = false })
    alone[r + 1] = true
  end
  local opened = M.paragraphs(lines, alone)
  for k, block in ipairs(opened) do
    assert(block.text == texts[k], "neoffice: paragraph layout mismatch" .. (vim.env.NEOFFICE_DEBUG and (" " .. k .. ": " .. vim.inspect(block.text) .. " ~= " .. vim.inspect(texts[k])
      .. "\n" .. table.concat(vim.list_slice(lines, math.max(1, block.first - 3), block.last + 3), "\n")) or "")) -- (see markup.render)
  end
  -- what each paragraph is as opened (doc.sigs, to tell at a save which changed): worked out once
  -- the text shows (see M.finish_sigs)
  doc.sigs, doc.ctx_keys, doc.touched, doc.all_dirty = {}, {}, {}, false
  doc.pending = { blocks = opened, next = 1, tick = api.nvim_buf_get_changedtick(buf) }
  doc.attached = api.nvim_buf_attach(buf, false, {
    on_bytes = function(_, _, _, start_row, _, _, old_end_row, _, _, new_end_row)
      local d = M.docs[buf]
      if d ~= doc then -- (this document's only: opened again, read again from disk, wiped: no more)
        return true
      end
      -- (the rows changed since the fingerprints were taken, as they move with later changes: their
      -- paragraphs' are taken again at the next save)
      local delta = new_end_row - old_end_row
      local function moved(r) -- (a range of rows as the edit moves it)
        if r[1] > start_row + old_end_row then
          r[1], r[2] = r[1] + delta, r[2] + delta
        elseif r[2] >= start_row then
          r[2] = math.max(r[2] + delta, start_row + new_end_row)
        end
      end
      -- (one range for the rows edited here and those it overlaps or touches: typing in a paragraph
      -- keeps one, however long it goes on)
      local new, kept = { start_row, start_row + new_end_row }, {}
      for _, r in ipairs(d.touched) do
        moved(r)
        if r[2] + 1 >= new[1] and r[1] <= new[2] + 1 then
          new[1], new[2] = math.min(new[1], r[1]), math.max(new[2], r[2])
        else
          kept[#kept + 1] = r
        end
      end
      kept[#kept + 1] = new
      d.touched = kept
      -- (and the rows to decorate again: see M.decorate_changed)
      local dirty = d.decor_dirty
      if dirty then
        moved(dirty)
        dirty[1], dirty[2] = math.min(dirty[1], start_row), math.max(dirty[2], start_row + new_end_row)
      else
        d.decor_dirty = { start_row, start_row + new_end_row }
      end
    end,
    on_reload = function() end, -- (read again from disk: opened again, as a new document; this one stays attached)
    on_detach = function() -- (no longer told what changes: every paragraph's fingerprint taken again at each save)
      if M.docs[buf] == doc then
        doc.untracked = true
      end
    end,
  })
  local function step(budget, most) -- (a few at a time, so typing isn't held up)
    local d = M.docs[buf]
    local p = d and d.pending
    if not p then
      return
    end
    if api.nvim_buf_get_changedtick(buf) ~= p.tick then -- (edited meanwhile: the rest once typing pauses)
      return M.finish_sigs_later(buf)
    end
    local deadline = vim.uv.hrtime() + (budget or 4e6) -- (4 ms at a time, so a key typed meanwhile isn't kept waiting)
    local k = p.next
    while k <= #p.blocks do
      d.sigs[k], d.ctx_keys[k] = signature(pieces(buf, p.blocks[k])), context_key(p.blocks[k])
      k = k + 1
      if vim.uv.hrtime() > deadline or most and k > most then
        break
      end
    end
    p.next = k
    if p.next > #p.blocks then
      d.pending = nil
    else
      vim.defer_fn(function() step() end, 0) -- (on a timer: keys typed meanwhile go first)
    end
  end
  vim.bo[buf].modified = false
  doc.unsaved_other = nil -- (a change neither to its text nor to its comments, not saved yet: see comments.settle)
  require("neoffice.comments").reset(buf) -- (as opened: what was done to its comments before is gone)
  setup_buffer(buf)
  require("neoffice.changes").start(buf)
  -- (the first ones now, for a few milliseconds, a hundred at most: all of a short document's, the
  -- first screens of a long one's, where an edit right after opening is; the rest after)
  step(20e6, 100)
  if new then
    vim.notify("neoffice: New document: it is created when you write it")
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
    vim.notify("neoffice: Could not read " .. path .. ": " .. err, vim.log.levels.ERROR)
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

-- A line at the top highlighted as tree-sitter shows YAML (an R Markdown file's header): each line
-- read by itself, without its indentation, as YAML reads the indented styles as wrong (a key with
-- braces has nothing under it), though each line alone is right. False where there's no YAML parser,
-- or the line alone isn't YAML (then neoffice's own: names, and the rest in gray).
local yaml_q
local function yaml_query()
  if yaml_q == nil then
    local ok, q = pcall(vim.treesitter.query.get, "yaml", "highlights")
    yaml_q = ok and q or false
  end
  return yaml_q
end
local function yaml_line(buf, r, line)
  local q = yaml_query()
  if not q then
    return false
  end
  local lead = #line:match("^%s*")
  local text = line:sub(lead + 1)
  local ok, parser = pcall(vim.treesitter.get_string_parser, text, "yaml")
  local tree = ok and parser:parse()[1]
  if not tree or tree:root():has_error() then
    return false
  end
  local k = 0 -- (a later capture over an earlier one of the same text, as tree-sitter shows them)
  for id, node in q:iter_captures(tree:root(), text, 0, 1) do
    local name = q.captures[id]
    local _, c1, _, c2 = node:range()
    if not name:find("^_") and name ~= "spell" and name ~= "nospell" and c2 > c1 then
      k = k + 1
      api.nvim_buf_set_extmark(buf, NS_DECOR, r - 1, lead + c1, { end_col = lead + c2, hl_group = "@" .. name .. ".yaml", priority = 150 + math.min(k, 40) })
    end
  end
  return true
end

-- Decorations that follow the text: titles highlighted (NeofficeTitle, above tree-sitter) and,
-- when the document numbers its headings, their numbers (1, 1.1, ...) shown before them.
-- (each paragraph's characters, by its text, for each buffer: of the texts decorated last time
-- only, so it doesn't grow)
local parsed = {}
-- Decorate rows s..e (1-based) of a buffer, whose decorations there are cleared: `lines` has them
-- (by row, and the row after them), `paras` their paragraphs (rows as the buffer's), `header` the
-- styles at the top if decorated (rows in s..e), and `number(key, row, col)` places a heading's
-- number (key: its level) or a caption's ("caption.table"). `fresh`: each paragraph's characters
-- read again, not kept (see `parsed`).
local function decorate_rows(buf, doc, lines, s, e, paras, header, number, fresh)
  for _, m in ipairs(api.nvim_buf_get_extmarks(buf, M.ns.change, { s - 1, 0 }, { e - 1, -1 }, { details = true, overlap = true })) do
    local ch = doc.change[m[1]]
    for row = math.max(m[2], s - 1), math.min(m[4].end_row - 1, e - 1) do -- (a deleted line break: shown as a struck ¶ at the end of its line)
      if ch and ch.type == "delete" then
        local hl = require("neoffice.changes").style(ch).hl_group
        api.nvim_buf_set_extmark(buf, NS_DECOR, row, #lines[row + 1], { virt_text = { { "¶", hl } }, virt_text_pos = "inline" })
      end
    end
  end
  if M.config.show_tabs then -- tabs: a → in the marks' gray, over the one 'list' shows (whatever colours the
    -- window is given: some colour schemes set its 'winhighlight' as they like)
    for r = s, e do
      local line = lines[r]
      local c = line:find("\t", 1, true)
      while c do
        api.nvim_buf_set_extmark(buf, NS_DECOR, r - 1, c - 1, { virt_text = { { "→", "NeofficeMark" } }, virt_text_pos = "overlay", priority = 150 })
        c = line:find("\t", c + 1, true)
      end
      -- (and a non-breaking space, which looks like a space and isn't: a ·, as LibreOffice shades it)
      c = line:find("\194\160", 1, true)
      while c do
        api.nvim_buf_set_extmark(buf, NS_DECOR, r - 1, c - 1, { virt_text = { { "·", "NeofficeMark" } }, virt_text_pos = "overlay", priority = 150 })
        c = line:find("\194\160", c + 2, true)
      end
    end
  end
  -- the styles at its top: each line as tree-sitter shows YAML (see yaml_line); without a YAML
  -- parser, the names, comments, and the rest in gray
  local function attributes(r, line, from)
    api.nvim_buf_set_extmark(buf, NS_DECOR, r - 1, from - 1, { end_col = #line, hl_group = "NeofficeMark", priority = 150 })
  end
  for r = s, header and math.min(header.last, e) or 0 do
    local line = lines[r]
    local s_, e_ = line:find("^%s+[^:{}]-%s*:")
    if line == "---" and yaml_query() then -- (as markdown shows them: not here)
    elseif line ~= "---" and yaml_line(buf, r, line) then -- (as an R Markdown file's YAML: see yaml_line)
    elseif line:match("^%s*#") then -- (a comment, as in YAML)
      api.nvim_buf_set_extmark(buf, NS_DECOR, r - 1, 0, { end_col = #line, hl_group = "NeofficeHeaderComment", priority = 150 })
    elseif line:match("^page:") then -- (the page's settings: its name as a style's)
      api.nvim_buf_set_extmark(buf, NS_DECOR, r - 1, 0, { end_col = 4, hl_group = "NeofficeStyleName", priority = 150 })
      attributes(r, line, 5)
    elseif line:match("^styles:") then
      api.nvim_buf_set_extmark(buf, NS_DECOR, r - 1, 0, { end_col = 6, hl_group = "NeofficeStyleName", priority = 150 })
      api.nvim_buf_set_extmark(buf, NS_DECOR, r - 1, 6, { end_col = #line, hl_group = "NeofficeMark", priority = 150 })
    elseif s_ and r > 2 and r < header.last then
      local first = line:find("%S")
      local name_end = line:sub(1, e_ - 1):find("%s*$") -- (its name, not the spaces before its ":")
      api.nvim_buf_set_extmark(buf, NS_DECOR, r - 1, first - 1, { end_col = name_end - 1, hl_group = "NeofficeStyleName", priority = 150 })
      attributes(r, line, name_end)
    else
      api.nvim_buf_set_extmark(buf, NS_DECOR, r - 1, 0, { end_col = #line, hl_group = "NeofficeMark", priority = 150 })
    end
  end
  for r = s, e do -- a table's options (above it); a format's line (its mark, its attributes)
    local line = lines[r]
    if line:find("^\\newpage%s*{") and markup.section_line(line) then -- (a section's: as a box's line)
      local brace = line:find("{", 1, true)
      api.nvim_buf_set_extmark(buf, NS_DECOR, r - 1, 0, { end_col = brace - 1, hl_group = "NeofficeSection", priority = 150 })
      api.nvim_buf_set_extmark(buf, NS_DECOR, r - 1, brace - 1, { end_col = #line, hl_group = "NeofficeMark", priority = 150 })
    end
    if line:find("^\\") and ((markup.block(line)).vspace or (markup.block(line)).toc) then -- (vertical space, a table of contents: as a page break)
      api.nvim_buf_set_extmark(buf, NS_DECOR, r - 1, 0, { end_col = #line, hl_group = "neofficeBreak", priority = 150 })
    end
    local fence, attrs = markup.box_line(line)
    if fence then -- (a box's line: its ::: as a table's |, its attributes as formatting)
      local brace = attrs and line:find("{", 1, true)
      api.nvim_buf_set_extmark(buf, NS_DECOR, r - 1, 0, { end_col = brace and brace - 1 or #line, hl_group = "NeofficeBox", priority = 120 })
      if brace then
        api.nvim_buf_set_extmark(buf, NS_DECOR, r - 1, brace - 1, { end_col = #line, hl_group = "NeofficeMark", priority = 120 })
      end
    end
    if line:sub(1, 1) == "{" and markup.table_options(line) and (lines[r + 1] or ""):match("^|") then
      api.nvim_buf_set_extmark(buf, NS_DECOR, r - 1, 0, { end_col = #line, hl_group = "NeofficeTableOptions", priority = 150 })
    elseif line:find("^[<%-{]") and not (lines[r + 1] or ""):match("^|") and markup.para_line(line) then
      local brace = line:find("{", 1, true)
      api.nvim_buf_set_extmark(buf, NS_DECOR, r - 1, 0, { end_col = brace and brace - 1 or #line, hl_group = "NeofficeAlign", priority = 250 })
      if brace then
        api.nvim_buf_set_extmark(buf, NS_DECOR, r - 1, brace - 1, { end_col = #line, hl_group = "NeofficeTableOptions", priority = 250 })
      end
    end
  end
  for _, para in ipairs(paras) do
    -- (its marks, in gray: only a paragraph with a character that could be one has any)
    if para.text:find("[\\*_~=+^%[{$!<⟦¶]") or para.text:find("^[#%%>%-]") or para.text:find("^%d+[.)]") then
      M.decorate_marks(buf, para, lines, fresh)
    end
    local block, text_start = markup.block(lines[para.first])
    -- the alignment mark (above Markdown's highlighting, which takes ">..." for a quote)
    local own, own_start, _, align_at = markup.block(para.text)
    if own.align and align_at then -- (its mark; on the line after its floating images, if they are on one of their own)
      local row = para.first - 1 + (para.lead and 1 or 0)
      local col = (para.col or 0) + align_at - 1
      api.nvim_buf_set_extmark(buf, NS_DECOR, row, col, {
        end_col = (para.col or 0) + own_start - 2, hl_group = "NeofficeAlign", priority = 250,
      })
    end
    if block.caption and not para.cell then -- its number (as the document shows it): "Table 1: ..."
      number("caption." .. block.caption, para.first - 1, #markup.captions()[block.caption])
      api.nvim_buf_set_extmark(buf, NS_DECOR, para.first - 1, 0, {
        end_col = text_start - 1, hl_group = "NeofficeCaption", priority = 200,
      })
    elseif block.title then
      api.nvim_buf_set_extmark(buf, NS_DECOR, para.first - 1, 0, {
        end_row = para.last - 1, end_col = #lines[para.last], hl_group = "NeofficeTitle", priority = 200,
      })
    elseif block.heading and doc.ctx.numbered then
      number(block.heading, para.first - 1, text_start - 1)
    end
  end
end

-- (a heading's or caption's number, shown: doc.numbers[mark] = its key, see decorate_rows)
local function put_number(buf, doc, key, row, col, text)
  local id = api.nvim_buf_set_extmark(buf, NS_DECOR, row, col, { virt_text = { { text, "NeofficeNumber" } }, virt_text_pos = "inline" })
  doc.numbers[id] = key
end
local function counter() -- (the numbers in order: the headings' "1.2 ", the captions' " 3")
  local count = {}
  return function(key)
    count[key] = (count[key] or 0) + 1
    if type(key) ~= "number" then
      return " " .. count[key]
    end
    for deeper = key + 1, 9 do
      count[deeper] = nil
    end
    local number = {}
    for level = 1, key do
      number[level] = count[level] or 0
    end
    return table.concat(number, ".") .. " "
  end
end
-- (the styles at the top, decorated here, also where tree-sitter would show them as YAML: the styles
-- under others are indented under them, which YAML doesn't read, a key with braces having none)
local function decorated_header(lines)
  return markup.header(lines)
end

function M.decorate(buf)
  buf = buf == 0 and api.nvim_get_current_buf() or buf
  local doc = M.docs[buf]
  if not doc then
    return
  end
  doc.decor_dirty, doc.decorating = nil, nil
  api.nvim_buf_clear_namespace(buf, NS_DECOR, 0, -1)
  for other in pairs(parsed) do -- (a buffer no longer a document's: forgotten)
    if not M.docs[other] then
      parsed[other] = nil
    end
  end
  parsed[buf] = { last = parsed[buf] and parsed[buf].now or {}, now = {} }
  local lines = api.nvim_buf_get_lines(buf, 0, -1, false)
  local header = markup.header(lines)
  doc.header_rows = header and header.last or 0
  doc.numbers = {}
  doc.has_boxes = false -- (whether a box's line is in it: see M.decorate_changed)
  for _, l in ipairs(lines) do
    if l:byte(1) == 58 and markup.box_line(l) then -- (":")
      doc.has_boxes = true
      break
    end
  end
  local next_number = counter()
  decorate_rows(buf, doc, lines, 1, #lines, M.blocks(buf), decorated_header(lines), function(key, row, col)
    put_number(buf, doc, key, row, col, next_number(key))
  end)
end

-- The first time (once the text shows): as M.decorate, a few paragraphs at a time (all of a short
-- document at once), so a long one doesn't hold up the first keys typed. If the text changes
-- meanwhile, all of it once typing pauses (until then, edits aren't decorated on their own).
function M.decorate_first(buf)
  local doc = M.docs[buf]
  if not doc then
    return
  end
  local tick = api.nvim_buf_get_changedtick(buf)
  doc.decor_dirty = nil
  api.nvim_buf_clear_namespace(buf, NS_DECOR, 0, -1)
  parsed[buf] = { last = parsed[buf] and parsed[buf].now or {}, now = {} }
  local lines = api.nvim_buf_get_lines(buf, 0, -1, false)
  local header = markup.header(lines)
  doc.header_rows = header and header.last or 0
  doc.numbers, doc.has_boxes = {}, false
  for _, l in ipairs(lines) do
    if l:byte(1) == 58 and markup.box_line(l) then -- (":")
      doc.has_boxes = true
      break
    end
  end
  local paras, shown, next_number = M.blocks(buf), decorated_header(lines), counter()
  local function number(key, row, col)
    put_number(buf, doc, key, row, col, next_number(key))
  end
  local state = {}
  doc.decorating = state
  local k, row = 1, 1 -- (the next paragraph, and the first row not decorated yet)
  local function slice(budget)
    if M.docs[buf] ~= doc or doc.decorating ~= state then -- (gone, or decorated since)
      return
    end
    if api.nvim_buf_get_changedtick(buf) ~= tick then -- (edited meanwhile: all of it once typing pauses)
      local function later(t)
        vim.defer_fn(function()
          if M.docs[buf] ~= doc or doc.decorating ~= state then
            return
          end
          local now = api.nvim_buf_get_changedtick(buf)
          if now ~= t then
            return later(now)
          end
          M.decorate(buf)
        end, 300)
      end
      return later(api.nvim_buf_get_changedtick(buf))
    end
    local deadline = vim.uv.hrtime() + budget
    while k <= #paras do
      local from = k -- (a few paragraphs, and the rows from the first not done to the next paragraph's)
      repeat
        k = k + 1
      until k > #paras or k - from >= 8
      local last = k <= #paras and paras[k].first - 1 or #lines
      decorate_rows(buf, doc, lines, row, last, vim.list_slice(paras, from, k - 1), row == 1 and shown or nil, number)
      row = last + 1
      if vim.uv.hrtime() > deadline then -- (the next few on a timer: keys typed meanwhile go first)
        return vim.defer_fn(function() slice(4e6) end, 0)
      end
    end
    if row <= #lines then -- (rows after the last paragraph: blank lines...)
      decorate_rows(buf, doc, lines, row, #lines, {}, row == 1 and shown or nil, number)
    end
    doc.decorating = nil
  end
  slice(20e6)
end

-- After an edit: only the paragraphs it touched decorated again (the rows Neovim said changed, out
-- to the blank lines around them: whole paragraphs, tables, format's lines), as typing goes. Their
-- headings' and captions' numbers stay as they were unless which there are changed: then all are
-- numbered again.
local SPAN = 64
function M.decorate_changed(buf)
  local doc = M.docs[buf]
  local dirty = doc and doc.decor_dirty
  if not dirty then
    return
  end
  if not doc.numbers or doc.decorating then -- (not decorated yet, or being decorated: that does all of it)
    return
  end
  doc.decor_dirty = nil
  local n = api.nvim_buf_line_count(buf)
  local s, e = math.min(dirty[1] + 1, n), math.min(dirty[2] + 1, n)
  while s > 1 do -- (up to a blank line)
    local from = math.max(1, s - SPAN)
    local chunk, found = api.nvim_buf_get_lines(buf, from - 1, s - 1, false), nil
    for i = #chunk, 1, -1 do
      if chunk[i] == "" then
        found = from + i - 1
        break
      end
    end
    if found then
      s = found + 1
      break
    end
    s = from
  end
  while e < n do -- (down to one)
    local to = math.min(n, e + SPAN)
    local chunk, found = api.nvim_buf_get_lines(buf, e, to, false), nil
    for i = 1, #chunk do
      if chunk[i] == "" then
        found = e + i
        break
      end
    end
    if found then
      e = found - 1
      break
    end
    e = to
  end
  if s <= (doc.header_rows or 0) + 1 then -- (the styles at the top: read from the first row)
    s = 1
  end
  if e < s or e - s > 5000 then -- (nothing left, or most of it: all of it)
    return M.decorate(buf)
  end
  local lines, slice = {}, api.nvim_buf_get_lines(buf, s - 1, math.min(e + 1, n), false)
  for i, l in ipairs(slice) do
    lines[s + i - 1] = l
  end
  slice[e - s + 2] = nil -- (the row after them: only looked at)
  local header = s == 1 and markup.header(slice) or nil
  local before = {} -- (read after the boxes' lines above them, so as inside a box or not as they are)
  if s == 1 then
    doc.header_rows = header and header.last or 0
  else -- (and a blank line: the rows can't read as the styles at the top, nor go on a paragraph)
    for _, l in ipairs(doc.has_boxes ~= false and api.nvim_buf_get_lines(buf, doc.header_rows or 0, s - 1, false) or {}) do
      if l:byte(1) == 58 and markup.box_line(l) then -- (":")
        before[#before + 1] = l
      end
    end
    before[#before + 1] = ""
  end
  local shift = s - 1 - #before
  slice = vim.list_extend(before, slice)
  local alone = {}
  for _, m in ipairs(api.nvim_buf_get_extmarks(buf, M.ns.empty, { s - 1, 0 }, { e - 1, -1 }, { details = true })) do
    if m[3] == 0 and not m[4].invalid then
      alone[m[2] + 1 - shift] = true
    end
  end
  local paras = {}
  for _, para in ipairs(M.paragraphs(slice, alone)) do
    if para.first + shift >= s then -- (not the blank line put first)
      para.first, para.last, para.owns = para.first + shift, para.last + shift, para.owns and para.owns + shift
      paras[#paras + 1] = para
    end
  end
  -- (a box's line there, before or now: whether what follows is in a box changes, so all of it)
  for r = s, e do
    if lines[r]:byte(1) == 58 and markup.box_line(lines[r]) then
      return M.decorate(buf)
    end
  end
  -- the numbers there: as they were, if the same headings and captions are there
  local was, old = {}, api.nvim_buf_get_extmarks(buf, NS_DECOR, { s - 1, 0 }, { e - 1, -1 }, { details = true })
  for _, m in ipairs(old) do
    if m[4].hl_group == "NeofficeBox" then
      return M.decorate(buf)
    end
  end
  for _, m in ipairs(old) do
    if doc.numbers[m[1]] then
      was[#was + 1] = { key = doc.numbers[m[1]], text = m[4].virt_text[1][1] }
      doc.numbers[m[1]] = nil
    end
  end
  api.nvim_buf_clear_namespace(buf, NS_DECOR, s - 1, e)
  local wanted = {}
  decorate_rows(buf, doc, lines, s, e, paras, header and decorated_header(slice), function(key, row, col)
    wanted[#wanted + 1] = { key = key, row = row, col = col }
  end, true)
  local same = #wanted == #was
  for k = 1, same and #wanted or 0 do
    same = same and wanted[k].key == was[k].key
  end
  if not same then -- (a heading or caption more, or less, or of another level: all numbered again)
    return M.decorate(buf)
  end
  for k, w in ipairs(wanted) do
    put_number(buf, doc, w.key, w.row, w.col, was[k].text)
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
local function chars_of(buf, text) -- (see `parsed`)
  parsed[buf] = parsed[buf] or { last = {}, now = {} }
  local cache = parsed[buf]
  local chars = cache.last[text] or cache.now[text] or select(2, markup.parse(text))
  cache.now[text] = chars
  return chars
end

function M.decorate_marks(buf, para, lines, fresh)
  local chars = fresh and select(2, markup.parse_shared(para.text)) or chars_of(buf, para.text) -- (only read)
  local function position(offset, ends) -- (the paragraph's lines were joined by one character)
    if para.cell then
      return para.first - 1, para.col + offset
    end
    offset = M.buffer_offset(para, offset, ends)
    local row = para.first
    while row < para.last and offset > #lines[row] do
      offset, row = offset - #lines[row] - M.join(para, row - para.first + 1), row + 1
    end
    return row - 1, offset
  end
  -- what isn't text is formatting (marks, attributes, a backslash...): in the marks' gray. Not what
  -- says what kind of paragraph it is ("# ", "% ", "- ", "> ", "Table: "): in that kind's colour
  local text = para.text
  local _, text_start, _, align_at, brace_at = markup.block(text)
  local at = (align_at or brace_at or text_start) - 1 -- (from its mark, or its braces: "{size = 9pt} ...")
  local function gray(from, to) -- (bytes from..to-1 of the text)
    if to > from and text:sub(from + 1, to):find("%S") then
      local r1, c1 = position(from)
      local r2, c2 = position(to, true)
      api.nvim_buf_set_extmark(buf, NS_DECOR, r1, c1, { end_row = r2, end_col = c2, hl_group = "NeofficeMark", priority = 120 })
    end
  end
  for _, c in ipairs(chars) do
    gray(at, c.col)
    at = math.max(at, c.col + (c.len or #c.ch))
    local brace = c.image and text:sub(c.col + 1, at):find("){", 1, true) -- (an image's attributes)
    if brace then
      gray(c.col + brace, at)
    end
  end
  gray(at, #text)
  -- a link or an image, all of it, in links' colour ("[text](url)", "![alt](path)": not an image's
  -- attributes, which are its format)
  local function paint(from, to, group) -- (bytes from..to-1 of the text)
    if to > from then
      local r1, c1 = position(from)
      local r2, c2 = position(to, true)
      api.nvim_buf_set_extmark(buf, NS_DECOR, r1, c1, { end_row = r2, end_col = c2, hl_group = group, priority = 125 })
    end
  end
  for k, c in ipairs(chars) do
    local len = c.len or #c.ch
    if c.image then
      local brace = text:sub(c.col + 1, c.col + len):find("){", 1, true)
      paint(c.col, brace and c.col + brace or c.col + len, "NeofficeLink")
    elseif c.link and not (chars[k - 1] and chars[k - 1].link == c.link) then
      local j = k
      while chars[j + 1] and chars[j + 1].link == c.link do
        j = j + 1
      end
      local before = chars[k - 1] and chars[k - 1].col + (chars[k - 1].len or #chars[k - 1].ch) or 0
      local open_at = text:sub(before + 1, c.col):match(".*()%[") -- (its "[": after marks before it)
      local after = chars[j].col + (chars[j].len or #chars[j].ch)
      local close = text:sub(after + 1, after + 2) == "](" and text:find(")", after + 3 + #c.link, true)
      paint(open_at and before + open_at - 1 or c.col, close or after, "NeofficeLink")
    end
  end
  local function mark(from, to, group) -- (chars from..to)
    local r1, c1 = position(chars[from].col)
    local r2, c2 = position(chars[to].col + (chars[to].len or #chars[to].ch), true)
    api.nvim_buf_set_extmark(buf, NS_DECOR, r1, c1, { end_row = r2, end_col = c2, hl_group = group, priority = 150 })
  end
  -- what syntax/neoffice.vim would show, unless it does (the marks concealed: see setup_buffer), under
  -- everything else as it would be: underlines, superscripts, subscripts, footnotes, fields, page breaks
  if not M.config.conceal then
    local function under(from, to, group) -- (bytes from..to-1 of the text)
      if to > from then
        local r1, c1 = position(from)
        local r2, c2 = position(to, true)
        api.nvim_buf_set_extmark(buf, NS_DECOR, r1, c1, { end_row = r2, end_col = c2, hl_group = group, priority = 90 })
      end
    end
    for _, key in ipairs({ "u", "sup", "sub" }) do -- (each run of such characters, as one: what is between them too)
      local group = key == "u" and "neofficeUnderline" or key == "sup" and "neofficeSuperscript" or "neofficeSubscript"
      local from
      for k = 1, #chars + 1 do
        local c = chars[k]
        if from and not (c and c[key]) then
          local last = chars[k - 1]
          under(chars[from].col, last.col + (last.len or #last.ch), group)
          from = nil
        end
        if c and c[key] and not from then
          from = k
        end
      end
    end
    for _, c in ipairs(chars) do
      if c.footnote then -- ("^[" and "]" as marks, its text as a footnote's)
        under(c.col, c.col + 2, "neofficeFootnoteMark")
        under(c.col + 2, c.col + c.len - 1, "neofficeFootnote")
        under(c.col + c.len - 1, c.col + c.len, "neofficeFootnoteMark")
      elseif c.field then
        under(c.col, c.col + c.len, "neofficeField")
      elseif c.page_break then
        under(c.col, c.col + c.len, "neofficeBreak")
      end
    end
  end
  local GROUP = { b = "NeofficeBold", i = "NeofficeItalic", s = "NeofficeStrike" }
  for _, key in ipairs({ "h", "b", "i", "s" }) do
    local from
    for k = 1, #chars + 1 do
      local c, prev = chars[k], chars[k - 1]
      if from and not (c and c[key] == prev[key]) then
        local group = key == "h" and highlight_group(prev[key]) or (languages and GROUP[key]) -- (else markdown's)
        if group then
          mark(from, k - 1, group)
        end
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
    return vim.notify("neoffice: Not a document buffer", vim.log.levels.WARN)
  end
  local on = arg == "on" or (arg ~= "off" and not doc.ctx.numbered)
  if on ~= doc.ctx.numbered then
    doc.fmt.number_headings(doc.ctx, on)
    vim.bo[buf].modified = true
    doc.unsaved_other = true
  end
  M.decorate(buf)
  vim.notify("neoffice: Heading numbers " .. (on and "on" or "off"))
end

-- Give columns [c1, c2) of a line this bold/italic/underline (rewriting its marks) and these
-- run properties. Used to reject a tracked formatting change.
-- The spans that render characters of markup.parse (all they read as), with `flags` changed for
-- those `inside(c)`.
local function spans_of(chars, inside, flags)
  return vim.tbl_map(function(c)
    local span = { text = c.ch ~= "\n" and c.ch or nil, link = c.link, line_break = c.ch == "\n" or nil,
      image = c.image, math = c.math, footnote = c.footnote, field = c.field, page_break = c.page_break,
      font = c.font, size = c.size, color = c.color }
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

-- The paragraphs' fingerprints as opened, all of them now (before a save, and before anything
-- that changes the document but not its text: a comment, a change accepted...). If the text changed
-- meanwhile, those that did (or whose rows were touched, or whose format did) aren't taken as they
-- are now: they are written at the next save.
function M.finish_sigs(buf)
  local doc = M.docs[buf]
  local p = doc and doc.pending
  if not p then
    return
  end
  doc.pending = nil
  if api.nvim_buf_get_changedtick(buf) == p.tick then
    for k = p.next, #p.blocks do
      doc.sigs[k], doc.ctx_keys[k] = signature(pieces(buf, p.blocks[k])), context_key(p.blocks[k])
    end
    return
  end
  local blocks = M.blocks(buf)
  local old_of = match_lines(doc.keys or doc.texts, block_keys(blocks))
  for k, block in ipairs(blocks) do
    local o = old_of[k]
    -- (and its format still as opened: a format's line above it changes it, not its rows)
    -- (and its text as opened: the lines matched may pair it with another paragraph)
    if o and o >= p.next and doc.sigs[o] == nil and block.text == p.blocks[o].text and not touched(doc, block)
      and context_key(block) == context_key(p.blocks[o]) then
      doc.sigs[o], doc.ctx_keys[o] = signature(pieces(buf, block)), context_key(p.blocks[o])
    end
  end
  for o = p.next, #p.blocks do
    if doc.sigs[o] == nil then
      doc.sigs[o] = false -- (never the same: written)
    end
  end
end

-- The same, after edits, a few at a time and only while typing pauses (each time the text changes,
-- its paragraphs are matched with those opened again, once it pauses): so typing right after
-- opening isn't held up. (Anything that needs them all, a save..., takes the rest at once.)
local IDLE = 300
function M.finish_sigs_later(buf, delay) -- (delay: until the next few, if not a pause in typing)
  local doc = M.docs[buf]
  local p = doc and doc.pending
  if not p or p.waiting then
    return
  end
  p.waiting = true
  local tick = api.nvim_buf_get_changedtick(buf)
  vim.defer_fn(function()
    p.waiting = nil
    if M.docs[buf] ~= doc or doc.pending ~= p then
      return
    end
    if api.nvim_buf_get_changedtick(buf) ~= tick then -- (still typing)
      return M.finish_sigs_later(buf)
    end
    if p.map_tick ~= tick then -- (the paragraphs as they are now, and those they were)
      p.map = { blocks = M.blocks(buf), next = 1 }
      p.map.old_of = match_lines(doc.keys or doc.texts, block_keys(p.map.blocks))
      p.map_tick = tick
    end
    local map, deadline = p.map, vim.uv.hrtime() + 4e6
    while map.next <= #map.blocks do
      local block, o = map.blocks[map.next], map.old_of[map.next]
      map.next = map.next + 1
      -- (as in M.finish_sigs)
      if o and o >= p.next and doc.sigs[o] == nil and block.text == p.blocks[o].text and not touched(doc, block)
        and context_key(block) == context_key(p.blocks[o]) then
        doc.sigs[o], doc.ctx_keys[o] = signature(pieces(buf, block)), context_key(p.blocks[o])
      end
      if vim.uv.hrtime() > deadline then -- (the next few right after, unless typing goes on)
        return M.finish_sigs_later(buf, 0)
      end
    end
    doc.pending = nil
    for o = p.next, #p.blocks do
      if doc.sigs[o] == nil then
        doc.sigs[o] = false -- (never the same: written)
      end
    end
  end, delay or IDLE)
end

function M.finish_all() -- (every document's: before a change from a command or a panel)
  for buf, doc in pairs(M.docs) do
    M.finish_sigs(buf)
    doc.all_dirty = true -- (a change not to the text: every paragraph's fingerprint taken again at the save)
  end
end

-- The buffer's paragraphs, and the style of each that the document has as it is (unchanged since
-- read or written; nil for one edited or new): for the language server.
function M.block_styles(buf)
  local doc = M.docs[buf]
  local blocks = M.blocks(buf)
  local out = {}
  if doc and doc.fmt.style_of then
    local old_of = match_lines(doc.keys or doc.texts, block_keys(blocks))
    for k in ipairs(blocks) do
      local node = old_of[k] and doc.paras[old_of[k]]
      out[k] = node and doc.fmt.style_of(doc.ctx, node) or nil
    end
  end
  return blocks, out
end


local PARAGRAPHS = { ["w:p"] = true, ["text:p"] = true, ["text:h"] = true }
local CONTENT = vim.tbl_extend("force", PARAGRAPHS, { ["w:tbl"] = true, ["text:list"] = true, ["text:list-item"] = true, ["table:table"] = true })

local function empty(el)
  return not vim.iter(el.children):any(function(c) return CONTENT[c.tag] end)
end

-- A bookmark's ends (Word's and LibreOffice's): the other end's tag, and what names the bookmark.
local BOOKMARK_ENDS = {
  ["w:bookmarkStart"] = { "w:bookmarkEnd", "w:id" }, ["w:bookmarkEnd"] = { "w:bookmarkStart", "w:id" },
  ["text:bookmark-start"] = { "text:bookmark-end", "text:name" }, ["text:bookmark-end"] = { "text:bookmark-start", "text:name" },
}

-- The bookmarks with both their ends in the document: { [start tag .. " " .. id] = a copy of its start }.
local function paired_bookmarks(root)
  local seen, out = {}, {}
  for _, el in ipairs(xml.collect(root, BOOKMARK_ENDS)) do
    local other, key = unpack(BOOKMARK_ENDS[el.tag])
    local id = xml.attr(el, key) or ""
    seen[el.tag .. " " .. id] = el
    local start = el.tag:match("[Ss]tart$") and el or seen[other .. " " .. id]
    if seen[other .. " " .. id] then
      out[start.tag .. " " .. id] = xml.copy(start)
    end
  end
  return out, seen
end
M.paired_bookmarks = paired_bookmarks -- (at open: see M.open)

-- A bookmark that had both its ends and lost one (its paragraph written again without it): the other
-- end put back, in the paragraph of the one left (at its end, or at its start), so it has both.
local function pair_bookmarks(root, paired)
  local _, seen = paired_bookmarks(root)
  for key in pairs(paired or {}) do
    local start = seen[key]
    local start_tag, id = key:match("^(%S+) (.*)$")
    local end_tag, name = unpack(BOOKMARK_ENDS[start_tag])
    local stop = seen[end_tag .. " " .. id]
    local left = (start == nil) ~= (stop == nil) and (start or stop)
    local p = left
    while p and not PARAGRAPHS[p.tag] do
      p = p.parent
    end
    if p then
      if start then -- (its end, at the end of the start's paragraph)
        xml.append(p, xml.elem(end_tag, { { name, id } }))
      else -- (its start as it was, at the start of the end's paragraph: after its properties)
        local first = 1
        while p.children[first] and (p.children[first].tag == "w:pPr") do
          first = first + 1
        end
        xml.append(p, xml.copy(paired[key]), first)
      end
    end
  end
end

-- Remove deleted paragraphs. An empty list item or list goes too; a table cell keeps an empty
-- paragraph (it needs one), unless every paragraph of its table was deleted: then the table goes.
-- (A bookmark's end in one, whose other end stays, goes to the nearest paragraph that stays: a
-- bookmark over several paragraphs keeps both its ends; one over deleted paragraphs only goes.)
local function remove_paragraphs(fmt, nodes)
  local going, rescued = {}, {}
  for _, node in ipairs(nodes) do
    going[node] = true
  end
  for _, node in ipairs(nodes) do
    local siblings = node.parent and node.parent.children or {}
    local at = node.parent and xml.index(node)
    for _, el in ipairs(at and xml.collect(node, BOOKMARK_ENDS) or {}) do
      local near, after
      for k = at + 1, #siblings do -- (the next paragraph that stays, else the one before)
        if PARAGRAPHS[siblings[k].tag] and not going[siblings[k]] then
          near, after = siblings[k], true
          break
        end
      end
      for k = at - 1, 1, -1 do
        if near then
          break
        end
        if PARAGRAPHS[siblings[k].tag] and not going[siblings[k]] then
          near = siblings[k]
        end
      end
      if near then
        rescued[#rescued + 1] = { el = el, near = near, after = after }
      end
    end
  end
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
  if #rescued > 0 then
    local root = rescued[1].near
    while root.parent do
      root = root.parent
    end
    local present = {} -- (the bookmarks' ends still in the document)
    for _, el in ipairs(xml.collect(root, BOOKMARK_ENDS)) do
      present[el.tag .. " " .. (xml.attr(el, BOOKMARK_ENDS[el.tag][2]) or "")] = true
    end
    for _, r in ipairs(rescued) do
      local other, key = unpack(BOOKMARK_ENDS[r.el.tag])
      local id = xml.attr(r.el, key) or ""
      -- (its other end there, and itself not: written again already, its mark moved to a paragraph kept)
      if present[other .. " " .. id] and not present[r.el.tag .. " " .. id] then
        xml.remove(r.el)
        if r.after then -- (at the start of the paragraph after it, after its properties)
          local first = 1
          while r.near.children[first] and (r.near.children[first].tag == "w:pPr") do
            first = first + 1
          end
          xml.append(r.near, r.el, first)
        else -- (at the end of the one before it)
          xml.append(r.near, r.el)
        end
        present[r.el.tag .. " " .. id] = true
      end
    end
  end
  for cell, template in pairs(cells) do
    local tbl = cell.parent
    while tbl and tbl.tag ~= "w:tbl" and tbl.tag ~= "table:table" do
      tbl = tbl.parent
    end
    -- (Word's cell ends with a paragraph: not with a table in it, kept, its paragraphs around it deleted)
    local last = vim.iter(cell.children):filter(function(c) return CONTENT[c.tag] end):last()
    if tbl and tbl.parent and #xml.collect(tbl, PARAGRAPHS) == 0 then
      xml.remove(tbl)
    elseif (empty(cell) or cell.tag == "w:tc" and last and last.tag ~= "w:p" and last.tag ~= "w:sdt" and last.tag ~= "w:customXml") and (not tbl or tbl.parent) then
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
  -- (each element, its children and its attributes as they are: those flat, name, value...; nil
  -- when it had none, false when an empty list)
  local nodes, kids_of, attrs_of, count, seen = {}, {}, {}, 0, {}
  local function walk(root)
    if seen[root] then
      return
    end
    seen[root] = true
    local stack, top = { root }, 1
    while top > 0 do
      local n = stack[top]
      stack[top], top = nil, top - 1
      local kids, copy = n.children, {}
      for k = 1, #kids do
        local c = kids[k]
        copy[k] = c
        if c.children then
          top = top + 1
          stack[top] = c
        end
      end
      local attrs, flat = n.attrs, nil
      if attrs then
        flat = false
        if #attrs > 0 then
          flat = {}
          for k = 1, #attrs do
            local kv = attrs[k]
            flat[2 * k - 1], flat[2 * k] = kv[1], kv[2]
          end
        end
      end
      count = count + 1
      nodes[count], kids_of[count], attrs_of[count] = n, copy, flat
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
    for i = 1, count do
      local n, flat = nodes[i], attrs_of[i]
      n.children = kids_of[i]
      if flat == nil then
        n.attrs = nil
      else
        local attrs = {}
        for k = 1, flat and #flat / 2 or 0 do
          attrs[k] = { flat[2 * k - 1], flat[2 * k] }
        end
        n.attrs = attrs
      end
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
    old_of = match_lines(doc.keys or doc.texts, block_keys(blocks))
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
    return vim.notify("neoffice: This document wasn't opened correctly: not saving it", vim.log.levels.ERROR)
  end
  target = vim.fn.fnamemodify(target or doc.path, ":p")
  local ok, err = pcall(function()
    M.finish_sigs(buf) -- (in here: whatever goes wrong, the file isn't written, and you are told)
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
    assert(not blocks.unclosed_box, blocks.unclosed_box and ("line %d: this box is never closed (a line \":::\" after its last paragraph closes it)"):format(blocks.unclosed_box))
    assert(not blocks.stray_format, blocks.stray_format and ("line %d: a format's line without \"+\" is for the paragraph right below it, and none is (a \"+\" at its end makes it for all the paragraphs after it)"):format(blocks.stray_format))
    for _, block in ipairs(blocks) do -- (a box holds paragraphs: a table in one would be taken apart)
      assert(not (block.cell and block.box), block.cell and block.box and ("line %d: a table can't go in a box (\":::\"), only paragraphs can"):format(block.first))
    end
    doc.comment_at = settle_comments(buf, blocks)
    doc.change_first = {} -- (where each deletion's first mark is: see pieces)
    for _, m in ipairs(api.nvim_buf_get_extmarks(buf, M.ns.change, 0, -1, {})) do
      local ch = doc.change[m[1]]
      if ch and ch.type == "delete" and not doc.change_first[ch] then
        doc.change_first[ch] = { m[2], m[3] }
      end
    end
    local texts = {}
    for k, p in ipairs(blocks) do
      texts[k] = p.text
    end
    local old_of, deleted = match_lines(doc.keys or doc.texts, block_keys(blocks))
    -- (a table of contents... is kept as it is, or goes: a line typed over it, or one typed as one over
    -- a paragraph, is a new paragraph, or a new one, where the other went)
    for k, block in ipairs(blocks) do
      local o = old_of[k]
      if o then
        local now = (markup.block(block.text)).toc
        local was = doc.tocs[o] and (markup.block(doc.texts[o] or "")).toc
        if (now or was) and now ~= was then
          deleted[#deleted + 1], old_of[k] = o, nil
        end
      end
    end
    align_numbers(buf, blocks, old_of)

    -- 1. Work out every paragraph first: if anything is wrong (a missing image, an equation that
    -- can't be converted), nothing changes.
    -- the header's styles: what changed in each (an error if a line isn't right, or names a style
    -- the document hasn't); the header deleted, or a style's line, changes nothing
    -- (a style under another says what it says itself: one taken off its line, it has from the one
    -- above it again; the first of a tree says all it gives: one taken off, its default. See
    -- M.style_entries. The styles under one that changes, as in Word, follow it where they say nothing)
    local restyle = {} -- (style id -> { set = { key = value or false }, unset = { keys } })
    local header = doc.header and markup.header(api.nvim_buf_get_lines(buf, 0, -1, false))
    for _, st in ipairs(header and header.styles or {}) do
      assert(not st.bad, ("the styles at the top, line %d: %s"):format(st.row, st.bad or ""))
      local known = doc.header[st.name:lower()]
      local id = known and known.id or doc.fmt.style_id(doc.ctx, st.name)
        or doc.fmt.style_id(doc.ctx, st.name:match("%(([^()]+)%)%s*$") or "") -- ("Name (Id)": one of two of a name)
      assert(id, ("not a paragraph style of this document: %s (the styles at the top, line %d)"):format(st.name, st.row))
      if not known then -- (a line for a style that wasn't shown: as it would have been)
        known = M.style_entries(doc, { id })
        known = vim.iter(known):find(function(e) return e.id == id end)
      end
      local old, under = known.attrs, known.base ~= nil
      local set, unset = {}, {}
      for _, key in ipairs(pformat.STYLE_KEYS) do
        local was, now = old[key], st.attrs[key]
        if now ~= nil and (was == nil or not pformat.same_style_value(key, was, now)) then
          set[key] = now
        elseif now == nil and was ~= nil then
          if under then
            unset[#unset + 1] = key
          else
            set[key] = false
          end
        end
      end
      if next(set) or #unset > 0 then
        restyle[id] = { set = set, unset = unset }
      end
      st.id = id
    end

    -- the sections' page settings (see pages.lua): the document's at its top, each section's from its
    -- line (by its mark: the section it was; without one, a new one); only if they changed
    local section_plan
    if not doc.sections then -- (a section's line where they aren't shown: not dropped without a word)
      for _, block in ipairs(blocks) do
        assert(not block.section, block.section and ("line %d: sections aren't shown (setup({ page = false })), so a section's line can't be written"):format(block.section.row))
      end
    end
    if doc.sections then
      local pages = require("neoffice.pages")
      local top = markup.header(api.nvim_buf_get_lines(buf, 0, -1, false))
      local first = doc.sections[1].view
      if top and top.page then
        assert(not top.page.bad, ("the page's line, at the top (line %d): %s"):format(top.page.row, top.page.bad or ""))
        first = pages.apply(first, top.page.attrs)
      end
      local after = blocks.section_after or blocks[#blocks] and blocks[#blocks].section_after
      assert(not after, after and ("line %d: %s"):format(after.row, after.bad))
      local list = { { orig = 1, el = doc.sections[1].el, want = first, first_block = 1 } }
      local mark_at = {}
      for _, m in ipairs(api.nvim_buf_get_extmarks(buf, M.ns.section, 0, -1, {})) do
        mark_at[m[2]] = doc.section_mark[m[1]]
      end
      for k, block in ipairs(blocks) do
        local sec = block.section
        if sec then
          assert(not sec.bad, ("line %d: %s"):format(sec.row, sec.bad or ""))
          assert(k > 1, ("line %d: a section's line before the first paragraph (the first section's page is at the top: page: {...})"):format(sec.row))
          local j = mark_at[sec.row - 1]
          local orig = j and not vim.iter(list):any(function(s) return s.orig == j end) and j or nil
          list[#list + 1] = { orig = orig, el = orig and doc.sections[orig].el, want = pages.apply(list[#list].want, sec.attrs), first_block = k }
        end
      end
      local same = #list == #doc.sections
      for j, sec in ipairs(list) do
        local was = doc.sections[j]
        same = same and was and sec.orig == j and pages.render(sec.want) == pages.render(was.view)
          and old_of[sec.first_block] == was.start and (j == 1 or old_of[sec.first_block - 1] == was.start - 1)
      end
      section_plan = not same and list or nil
    end

    local plans, equations = {}, {}
    local reusable = not doc.all_dirty and not doc.untracked and not doc.tracking and next(restyle) == nil
    local first_block = M.first_paragraphs(#blocks, function(k) -- (first paragraphs: see M.first_paragraphs)
      local b = blocks[k]
      if b.cell or b.box then
        return "other"
      end
      local kind, start = markup.block(b.text)
      if kind.vspace then
        return nil -- (an empty paragraph, as opened)
      elseif kind.fence or kind.row then
        return "other"
      elseif b.text:sub(start or 1):match("^%s*$") and plain(kind) then
        return nil
      end
      return (kind.heading or kind.title or kind.quote) and "break" or "other"
    end)
    local settled_on = {} -- (paragraphs a comment was moved to, see settle_comments: written again, with it)
    if next(doc.comment_at or {}) then
      for k, block in ipairs(blocks) do
        for _, at in pairs(doc.comment_at) do
          for _, pos in ipairs({ { at[1], at[2] }, { at[3], at[4] } }) do
            local here = block.cell and pos[1] == block.first - 1 and pos[2] >= block.col and pos[2] <= block.stop
              or not block.cell and pos[1] >= block.first - 1 and pos[1] <= block.owns - 1
            settled_on[k] = settled_on[k] or here
          end
        end
      end
    end
    for k, block in ipairs(blocks) do
      assert(not (block.format and block.format.bad), block.format and block.format.bad) -- (its format's line: right)
      local o = old_of[k]
      if o and doc.texts[o] ~= block.text then -- (a mark it lost the other one of: see check_marks)
        check_marks(buf, block, doc.texts[o])
      end
      -- (an empty one too, unless it holds something, a bookmark...: empty lines all look alike, and
      -- its marks may have gone with another's; its fingerprint, from what its rows hold, tells)
      local bare = block.text ~= "" or not vim.iter(o and doc.paras[o] and doc.paras[o].children or {}):any(function(c)
        return c.tag and c.tag ~= "w:pPr"
      end)
      if reusable and o and bare and doc.texts[o] == block.text and doc.sigs[o] and doc.ctx_keys[o] == context_key(block) and not touched(doc, block)
        and not settled_on[k] then
        -- (as it was, its text, rows and format: its fingerprint as it was; its pieces only if needed)
        plans[k] = { old = o, lazy = block, sig = doc.sigs[o], cell = block.cell, box = block.box }
        goto planned
      end
      do
      local kind, list = pieces(buf, block, not old_of[k])
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
      assert(not kind.vspace_bad, kind.vspace_bad and ("line %d: %s"):format(block.first, kind.vspace_bad))
      if kind.toc then -- (a table of contents...: between paragraphs of the text; a new one, a table of contents only)
        assert(not block.cell and not block.box, ("line %d: %s goes between paragraphs of the text, not in a table cell or a box"):format(block.first, markup.INDEX_LINES[kind.toc]))
        assert(old_of[k] or kind.toc == "contents", ("line %d: %s is kept where the document has one, but neoffice can't make a new one (\\tableofcontents it can)"):format(block.first, markup.INDEX_LINES[kind.toc]))
      end
      -- (one that was \vspace and isn't now, typed over: nothing of the space to stay on it)
      kind.was_vspace = old_of[k] and not kind.vspace and (markup.block(doc.texts[old_of[k]] or "")).vspace ~= nil or nil
      if kind.vspace then -- (vertical space: an empty paragraph of the usual style, of that height)
        assert(not block.cell and not block.box, ("line %d: \\vspace goes between paragraphs of the text, not in a table cell or a box"):format(block.first))
        kind.style_id = doc.plain_style
      end
      if doc.plain_style and plain(kind) and not block.cell and not block.box then -- (a plain paragraph's style: see paras.lua)
        -- (the one it says; else the one it has, if that is one of the two usual ones and it wasn't
        -- shown: the usual one where it was; else the usual one where it is: a new paragraph's, one
        -- whose {style = ...} was taken away, a heading made text)
        local want = kind.para and kind.para.style
        local id = want and ((doc.style_labels or {})[want:lower()] or doc.fmt.style_id(doc.ctx, want)
          or doc.fmt.style_id(doc.ctx, want:match("%(([^()]+)%)%s*$") or ""))
        assert(not want or id, ("line %d: not a paragraph style of this document: %s"):format(block.first, want or ""))
        local current = o and doc.paras[o] and doc.fmt.style_of(doc.ctx, doc.paras[o])
        if not id and not (doc.style_shown or {})[o] and (current == doc.plain_style or current == doc.first_style) then
          id = current
        end
        kind.style_id = id or (first_block[k] and doc.first_style or doc.plain_style)
      end
      plans[k] = { old = old_of[k], block = kind, pieces = list, sig = signature(kind, list), cell = block.cell, box = block.box }
      assert(not (block.format and block.format.bad), block.format and block.format.bad) -- (its format's line: right)
      assert(not kind.para_bad, kind.para_bad) -- (its own attributes, after its mark: right)
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
      ::planned::
      if block.box and not block.box.parsed then -- (a box's attributes: for a box that is new)
        local attrs, bad = floats.parse_box(block.box.attrs)
        assert(attrs, bad)
        block.box.parsed = attrs
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
    local ctx, paras, prev, plain, placements, seps, box_keys, written_as = doc.ctx, {}, nil, nil, {}, {}, {}, {}
    local notes = vim.list_extend(vim.tbl_values(doc.change), vim.tbl_values(doc.comment))
    for _, c in pairs(doc.comment) do
      vim.list_extend(notes, c.replies)
    end
    local restore = checkpoint(ctx, notes)
    local integrity_after
    local tocs_now, new_tocs, heading_key = {}, {}, doc.heading_key -- (tables of contents...: kept, new ones; the headings: see below)
    local written, failure = pcall(function()
    -- the styles first (what the paragraphs are written against); those whose look changed (also
    -- those under one that changed): their paragraphs written again, so what they say of their own
    -- is against it
    local restyled_ids = {}
    if next(restyle) then
      local looks = {}
      local ids = vim.list_extend(vim.list_extend({}, doc.fmt.styles_in_use and doc.fmt.styles_in_use(ctx) or {}), vim.tbl_keys(restyle))
      for _, id in ipairs(ids) do
        local ok, view = pcall(doc.fmt.style_view, ctx, id)
        looks[id] = ok and view or nil
      end
      for id, change in pairs(restyle) do
        if #change.unset > 0 then
          doc.fmt.unset_style(ctx, id, change.unset)
        end
        if next(change.set) then
          doc.fmt.set_style(ctx, id, change.set)
        end
      end
      for id, was in pairs(looks) do
        local now = doc.fmt.style_view(ctx, id)
        for _, key in ipairs(pformat.STYLE_KEYS) do
          if not pformat.same_style_value(key, was[key], now[key]) then
            restyled_ids[id] = true
            break
          end
        end
      end
      for id in pairs(restyle) do
        restyled_ids[id] = true
      end
    end
    local layout = require("neoffice.tables").start(doc.fmt, ctx, vim.tbl_map(function(plan)
      return { cell = plan.cell, node = plan.old and doc.paras[plan.old] }
    end, plans), M.cell_of)
    local function outside(node) -- (the outermost table it is in, or itself; or a block content
      -- control (a table of contents...) it ends: what comes after it goes after it; or the paragraph
      -- its box is in, for one in a box: what goes next to it goes next to that paragraph, not in the box)
      local cell = M.cell_of(node)
      if cell then
        return outside(cell.tbl)
      end
      local box = doc.fmt.box_of and doc.fmt.box_of(node)
      local up = box and box.parent
      while up and not PARAGRAPHS[up.tag] do
        up = up.parent
      end
      if up then
        return outside(up)
      end
      local content, last = node.parent, nil
      for k = content and #content.children or 0, 1, -1 do -- (its last element)
        if content.children[k].tag then
          last = content.children[k]
          break
        end
      end
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
    -- (which were in a table's cell, asked before any is placed: the layout takes cells out of their
    -- tables as it goes, those deleted from the buffer, and a paragraph that was in one may come later)
    local was_in_cell = {}
    for k, plan in ipairs(plans) do
      was_in_cell[k] = plan.old and doc.paras[plan.old] and M.cell_of(doc.paras[plan.old]) ~= nil
    end
    local outer -- (the last paragraph that isn't in a box)
    tocs_now, new_tocs = {}, {} -- (tables of contents...: kept, and new ones, made once the headings are written)
    for k, plan in ipairs(plans) do
      local node = plan.old and doc.paras[plan.old]
      local toc_line = (markup.block(blocks[k].text)).toc
      local function fresh(template) -- (a new paragraph; a new table of contents' place, for now)
        return toc_line and xml.elem(doc.fmt.MAIN == "content.xml" and "text:p" or "w:p") or doc.fmt.new_para(template, ctx)
      end
      local prev_box = prev and box_of(prev)
      -- (where a new paragraph goes, after the one before it: asked before the layout takes a cell out)
      local anchor = prev and outside(prev) or doc.paras[1] and outside(doc.paras[1])
      local placed = layout:node({ cell = plan.cell, node = node, edited = not plan.old or plan.sig ~= doc.sigs[plan.old] or blocks[k].text ~= doc.texts[plan.old], empty = blocks[k].text == "" }, prev_box and outer or prev)
      if node and (placed and placed ~= node or not placed and was_in_cell[k]) then
        deleted[#deleted + 1], plan.old, node = plan.old, nil, nil -- (moved into or out of a table cell)
      end
      node = placed or node
      -- (moved into or out of a box; or into one that has no box in the document yet: a new one)
      if node and (box_of(node) ~= (plan.box and boxes[plan.box]) or plan.box and not boxes[plan.box]) then
        deleted[#deleted + 1], plan.old, node = plan.old, nil, nil
      end
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
        node = fresh(plain or doc.paras[1])
        if prev_box.parent and xml.index(prev_box) then
          xml.append(prev_box.parent.parent, node, xml.index(prev_box.parent))
        elseif outer then
          local after = outside(outer)
          xml.append(after.parent, node, xml.index(after) + 1)
        else
          local first = doc.paras[1] and outside(doc.paras[1])
          if first then
            xml.append(first.parent, node, xml.index(first))
          else -- (the document had no paragraph at all)
            node = doc.fmt.first_para(ctx)
          end
        end
      elseif not node and not anchor then -- the document had no paragraph at all
        node = doc.fmt.first_para(ctx)
      elseif not node then -- a new paragraph, formatted like the one before it
        node = fresh(plain or prev or doc.paras[1])
        xml.append(anchor.parent, node, xml.index(anchor) + (prev and 1 or 0))
      end
      -- (what a new paragraph after it is formatted like: not vertical space, which has a height of its own)
      plain = not plan.cell and not plan.box and not (markup.block(blocks[k].text)).vspace and not toc_line and node or plain
      if toc_line then
        tocs_now[k] = plan.old and doc.tocs[plan.old] or nil
        if not plan.old then
          new_tocs[#new_tocs + 1] = { k = k, place = node }
        end
      end
      outer = not plan.box and node or outer
      seps[k] = plan.old and doc.caption_seps[plan.old] or nil
      written_as[k] = plan.old and doc.caption_as[plan.old] or nil
      -- (a paragraph whose style changed: written again, so what it says of its own is against it)
      -- (not an empty one: its format isn't in the buffer, a blank line has none; it keeps its own)
      local restyled = next(restyle) and node and plan.old and not plan.cell and blocks[k].text ~= "" and restyled_ids[doc.fmt.style_of(ctx, node)]
      -- (and one whose text isn't the old paragraph's: whatever its fingerprint says)
      -- (a table of contents...: never written; a new one is made below)
      if not toc_line and (not plan.old or plan.sig ~= doc.sigs[plan.old] or restyled or blocks[k].text ~= doc.texts[plan.old]) then
        plan.written = true
        if not plan.pieces then -- (one taken as it was, written all the same: moved into a table, a box...)
          plan.block, plan.pieces = pieces(buf, plan.lazy, false)
        end
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
            or { object = doc.fmt.caption_number(ctx, as.seq or markup.captions()[kind]), visible = true, insert = plainly.insert }
          local labels = { vim.tbl_extend("force", plainly, { text = label .. " " }), number, vim.tbl_extend("force", plainly, { text = seps[k] }) }
          list = vim.list_slice(plan.pieces)
          local at = 1 -- (after deleted text before it, see find_captions; in its insertion, after where
          -- it starts: rejecting it takes them away too)
          while list[at] and (list[at].deletion or plainly.insert and (list[at].insert_start or list[at].comment_start)) do
            at = at + 1
          end
          for n, piece in ipairs(labels) do
            table.insert(list, at + n - 1, piece)
          end
          -- (all of it deleted, changes tracked, its break too: its word, number and separator deleted
          -- with it, as text, the number it showed; else accepting it would leave them)
          local mark = vim.iter(plan.pieces):find(function(p) return p.paragraph_mark and p.paragraph_mark.type == "delete" end)
          local live = vim.iter(plan.pieces):any(function(p) return p.text and p.text ~= "" end)
          local deleted = vim.iter(plan.pieces):any(function(p) return p.deletion end)
          if mark and deleted and not live then
            local shown = as.number or (node and require("neoffice.xml").text(node):match(vim.pesc(label) .. "%s*(%d+)")) or "1"
            local ch = mark.paragraph_mark
            local text = label .. " " .. shown .. seps[k]
            list = vim.list_extend({ { deletion = { type = "delete", author = ch.author, date = ch.date, parent = ch }, deleted = text, part = 0, segs = { vim.tbl_extend("force", plainly, { text = text }) } } }, plan.pieces)
          end
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
      if tocs_now[k] then -- (what comes after it goes after all of it: a TOC field's last paragraph)
        prev = tocs_now[k][#tocs_now[k]]
      end
    end
    layout:finish()
    -- the new tables of contents, of the headings as now written; and Word told to fill the fields
    -- (their page numbers) when it opens the document next, if one was made, or the headings one
    -- lists changed
    local headings = {}
    for k, block in ipairs(blocks) do
      local kind = markup.block(block.text)
      if kind.heading and not block.cell and not block.box then
        local _, chars = markup.parse(block.text)
        headings[#headings + 1] = { level = kind.heading, text = vim.trim(table.concat(vim.tbl_map(function(c) return c.ch end, chars))), node = paras[k] }
      end
    end
    for _, t in ipairs(new_tocs) do
      -- (a title right before it, a paragraph in the TOC heading's style: its, not a second one)
      local k, before = xml.index(t.place), nil
      local BLOCK = { ["w:p"] = true, ["w:tbl"] = true, ["w:sdt"] = true, ["text:p"] = true, ["text:h"] = true,
        ["table:table"] = true, ["text:list"] = true, ["text:section"] = true, ["text:table-of-content"] = true }
      for j = (k or 1) - 1, 1, -1 do -- (the block before it, past bookmarks' ends and the like between them)
        local c = t.place.parent.children[j]
        if c.tag and BLOCK[c.tag] then
          before = c
          break
        end
      end
      local title = doc.fmt.toc_title and doc.fmt.toc_title(ctx, before) and before or nil
      local el = doc.fmt.new_toc(ctx, headings, title)
      local parent, at = t.place.parent, xml.index(t.place)
      xml.remove(t.place)
      xml.append(parent, el, at)
      paras[t.k], tocs_now[t.k] = el, { el }
      ctx.dirty[doc.fmt.MAIN] = true
    end
    heading_key = table.concat(vim.tbl_map(function(h) return h.level .. " " .. h.text end, headings), "\n")
    if doc.fmt.update_fields_on_open and next(tocs_now) and (#new_tocs > 0 or heading_key ~= doc.heading_key) then
      doc.fmt.update_fields_on_open(ctx)
    end
    if #deleted > 0 then
      local nodes = {} -- (a table of contents: all it is)
      for _, old in ipairs(deleted) do
        vim.list_extend(nodes, doc.tocs[old] or { doc.paras[old] })
      end
      remove_paragraphs(doc.fmt, nodes)
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
        if plans[k + 1] and plans[k + 1].cell and paras[k + 1] then
          -- (a table right after it: in a paragraph of its own before the table, not in the one after
          -- the table, where it would show: the one it is in already, if that is it)
          local tbl = outside(paras[k + 1])
          local here = run.parent
          while here and not PARAGRAPHS[here.tag] do
            here = here.parent
          end
          local mine = here and not vim.tbl_contains(paras, here) and tbl.parent == here.parent
            and xml.index(here) == xml.index(tbl) - 1
          if mine then
            at = here
          else
            at = doc.fmt.new_para(plain or doc.paras[1], ctx)
            xml.append(tbl.parent, at, xml.index(tbl))
          end
        end
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
      if not box_keys[run] and run.parent and xml.index(run) then
        xml.remove(run)
        ctx.dirty[doc.fmt.MAIN] = true
      end
    end
    if section_plan then -- (each section where its line is, with what it says)
      for j, sec in ipairs(section_plan) do
        local first = sec.first_block -- (its first paragraph not in a box: the box's are shown before it)
        while blocks[first + 1] and blocks[first].box do
          first = first + 1
        end
        sec.first = paras[first]
        sec.holder = section_plan[j + 1] and paras[section_plan[j + 1].first_block - 1] or nil
      end
      doc.fmt.set_sections(ctx, section_plan)
    end
    if ctx.dirty[doc.fmt.MAIN] and doc.fmt.number_captions then -- (captions added, removed, moved)
      doc.fmt.number_captions(ctx)
    end

    -- 3. Write the parts that changed (nothing at all if nothing did).
    local live = { comments = vim.tbl_values(doc.comment), changes = {}, mark_decisions = doc.mark_decisions }
    for _, ch in pairs(doc.change) do
      live.changes[ch] = true
    end
    -- (never a document with more of what Word or LibreOffice would repair than it had: a bookmark
    -- twice, a cell without its paragraph...; then nothing is written)
    -- (every paragraph of the buffer with something in it in the document: none lost, as one left in
    -- a table's cell that was taken out would be; then nothing is written. An empty one may go: an
    -- empty cell beyond the table's columns, padding)
    for k, node in ipairs(paras) do
      local top = node
      while top.parent do
        top = top.parent
      end
      assert(top == ctx.root or blocks[k] and blocks[k].text == "", ("neoffice would have lost the text at line %d (this is a bug in neoffice, please report it)"):format(blocks[k] and blocks[k].first or 0))
    end
    pair_bookmarks(ctx.root, ctx.bookmarks)
    if doc.fmt.unique_drawings then
      doc.fmt.unique_drawings(ctx)
    end
    local entries, trees = doc.fmt.save(ctx, live)
    for _, c in ipairs(doc.fmt.comment_kept and live.comments or {}) do -- (every comment you see: in the document)
      assert(doc.fmt.comment_kept(ctx, c), ("neoffice would have lost the comment \"%s\" (this is a bug in neoffice, please report it)"):format(vim.fn.strcharpart(c.text or "", 0, 40)))
    end
    if doc.fmt.integrity then -- (after the comments, notes... are saved too: all that is written)
      local now = doc.fmt.integrity(ctx)
      for what, n in pairs(now) do
        assert(n <= (ctx.integrity[what] or 0), ("neoffice would have written %s (this is a bug in neoffice, please report it)"):format(what))
      end
      integrity_after = now
    end
    for name, bytes in pairs(entries) do -- (never a part Word or LibreOffice couldn't read: nothing is written then)
      if type(bytes) == "string" and (name:match("%.xml$") or name:match("%.rels$")) then
        local ok, why = xml.check((trees or {})[name] or bytes) -- (the tree it was written from: as good, and quick)
        assert(ok, ("neoffice would have written a broken %s (%s): this is a bug in neoffice, please report it"):format(name, why))
      end
    end
    if next(entries) then -- the base gets the changes...
      if not vim.uv.fs_stat(doc.base) then -- (gone, as when the temporary files are cleaned: the file is the same, if as last written or read)
        assert(doc.synced and not doc.is_doc and doc.stamp and stamp(doc.path) == doc.stamp,
          "neoffice's copy of the document is gone, and the file isn't as it was: write to another file (:w other." .. doc.path:match("%w+$") .. ")")
        vim.fn.mkdir(vim.fs.dirname(doc.base), "p")
        zip.copy(doc.path, doc.base)
      end
      zip.write(doc.base, entries)
      doc.synced = false
    end
    end)
    if not written then
      restore()
      error(failure, 0)
    end
    ctx.integrity = integrity_after or ctx.integrity -- (as now written: what the next save is held to)
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
    doc.tocs, doc.heading_key = tocs_now, heading_key
    doc.keys = block_keys(blocks)
    if doc.sections then -- (the sections as they now are, each line marked as its section)
      local pages = require("neoffice.pages")
      doc.sections = doc.fmt.sections(ctx, paras)
      for _, sec in ipairs(doc.sections) do
        sec.view = pages.view(sec.props)
      end
      api.nvim_buf_clear_namespace(buf, M.ns.section, 0, -1)
      doc.section_mark = {}
      local j = 1
      for _, block in ipairs(blocks) do
        if block.section then
          j = j + 1
          doc.section_mark[api.nvim_buf_set_extmark(buf, M.ns.section, block.section.row - 1, 0, { right_gravity = false })] = j
        end
      end
    end
    if next(restyle) then -- (the styles as they now are: what the header is compared with next, and shows)
      M.refresh_header(buf, doc)
    end
    for k, block in ipairs(blocks) do -- (those written may have new marks: new images, equations...)
      doc.sigs[k] = plans[k].written and signature(pieces(buf, block)) or plans[k].sig
      doc.ctx_keys[k] = context_key(block)
    end
    doc.touched, doc.all_dirty = {}, false -- (the fingerprints are the buffer's as it now is)
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
      warn_elsewhere(target, "wrote")
      zip.copy(file, target) -- (all of it, checked, then in place of the file: see zip.copy)
      if doc.is_doc and doc_cache() then -- (what was written, read back in the background: opened at once next time)
        local key = content_key(file)
        if key and not cached_doc(key) then
          convert(file, "docx", vim.fn.tempname(), function(docx)
            if docx then
              keep_doc(key, docx)
            end
          end)
        end
      end
      if target == doc.path then
        doc.synced, doc.stamp = true, stamp(target)
      end
    end
  end)
  doc.comment_at, doc.change_first = nil, nil
  if not ok then
    err = tostring(err):gsub("^[^\n]-%.lua:%d+: ", "") -- (where in neoffice it was found: not for you)
    return vim.notify("neoffice: Save failed, the file was not changed: " .. err, vim.log.levels.ERROR)
  end
  if target == doc.path then
    vim.bo[buf].modified = false
    doc.unsaved_other = nil
    require("neoffice.comments").saved(buf) -- (as saved: where undoing its comments makes it unmodified)
  end
  vim.notify("neoffice: Saved " .. vim.fn.fnamemodify(target, ":~:."))
  if target == doc.path then -- (its PDF again, if it is exported after each save: see pdf.lua)
    require("neoffice.pdf").saved(buf, doc.path)
  end
end

return M
