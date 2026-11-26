-- Reading documents, for other plugins (manuscript.nvim's merge, for one). A stable interface:
-- neoffice's internals may change, this doesn't (M.version says if it does).
--
-- A paragraph is { first, last (its buffer lines), text, sent, revised, comments }:
--   sent      its characters as the document was before its tracked changes (all rejected, with
--             the formatting from before formatting changes)
--   revised   its characters with them (all accepted)
--   comments  the comments on it: { author, text, resolved, replies = { { author, text }... } }
-- A character is { ch, b, i, u, s, h, sup, sub, link, ... } (see markup.parse): ch is "" for an
-- image, equation, footnote or field, which have { image }, { math }, { footnote }, { field }.
local api = vim.api
local M = { version = 1 }

local function neo()
  return require("neoffice")
end

function M.is_document(buf)
  return neo().docs[buf == 0 and api.nvim_get_current_buf() or buf] ~= nil
end

-- The document's file.
function M.path(buf)
  local doc = neo().docs[buf == 0 and api.nvim_get_current_buf() or buf]
  return doc and doc.path
end

-- Whether it has tracked changes.
function M.has_changes(buf)
  return #api.nvim_buf_get_extmarks(buf, neo().ns.change, 0, -1, { limit = 1 }) > 0
end

local function covers(m, row, col)
  local d = m[4]
  return (m[2] < row or m[2] == row and m[3] <= col) and (d.end_row > row or d.end_row == row and d.end_col > col)
end

-- A paragraph's characters as sent and as revised.
local function versions(buf, block)
  local doc = neo().docs[buf]
  local _, chars = require("neoffice.markup").parse(block.text)
  local lines = api.nvim_buf_get_lines(buf, block.first - 1, block.last, true)
  local row, start = 1, -(block.col or 0)
  for _, ch in ipairs(chars) do
    local join = require("neoffice").join
    while row < #lines and (ch.col > start + #lines[row] or ch.col == start + #lines[row] and join(block, row) == 0) do
      start, row = start + #lines[row] + join(block, row), row + 1
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

-- The paragraphs of a document's buffer.
function M.paragraphs(buf)
  buf = buf == 0 and api.nvim_get_current_buf() or buf
  local doc = neo().docs[buf]
  if not doc then
    return nil
  end
  local out = {}
  for _, block in ipairs(neo().blocks(buf)) do
    local sent, revised = versions(buf, block)
    local comments = {}
    for _, m in ipairs(api.nvim_buf_get_extmarks(buf, neo().ns.comment, { block.first - 1, 0 }, { block.last - 1, -1 }, {})) do
      local c = doc.comment[m[1]]
      if c then
        comments[#comments + 1] = {
          author = c.author, text = c.text, resolved = c.resolved or false,
          replies = vim.tbl_map(function(r) return { author = r.author, text = r.text } end, c.replies),
        }
      end
    end
    out[#out + 1] = { first = block.first, last = block.last, text = block.text, sent = sent, revised = revised, comments = comments }
  end
  return out
end

-- The paragraphs of a document file (.docx, .odt), read in a hidden buffer; nil and the error if it
-- can't be read.
function M.file(path)
  if not vim.uv.fs_stat(path) then -- (neoffice would open it as a new document)
    return nil, "no such file: " .. path
  end
  local b = api.nvim_create_buf(false, true)
  local notify = vim.notify
  local err
  vim.notify = function(msg) err = msg end -- (its errors: the caller's to report)
  pcall(neo().open, b, vim.fn.fnamemodify(path, ":p"))
  vim.notify = notify
  local out = M.paragraphs(b)
  api.nvim_buf_delete(b, { force = true })
  return out, not out and (err or "can't read " .. path) or nil
end

-- Hunks { start a, count a, start b, count b } between two lists of strings (as vim.diff gives).
function M.diff(a, b)
  return neo().diff(a, b)
end

return M
