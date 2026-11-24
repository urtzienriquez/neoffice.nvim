-- A language server for documents, running in Neovim itself (vim.lsp.start with a function as its
-- command): nothing to install. It reads the buffer as it is, and answers:
--   textDocument/documentSymbol  the title and the headings, nested (gO, pickers, outlines)
--   textDocument/foldingRange    each heading's section
--   textDocument/hover           the comment, or the tracked change, at the cursor
local api = vim.api
local M = {}

local function neo()
  return require("neoffice")
end

local function plain(text) -- (a paragraph's text, without its marks)
  local _, chars = require("neoffice.markup").parse(text)
  return vim.trim(table.concat(vim.tbl_map(function(c) return c.ch end, chars)):gsub("%s+", " "))
end

-- The title and the headings: { name, level (0: the title), first, last (0-based rows: its section) }.
local function headings(buf)
  local doc, list, count = neo().docs[buf], {}, {}
  local blocks = neo().blocks(buf)
  for _, para in ipairs(blocks) do
    local block = require("neoffice.markup").block(para.text)
    if (block.heading or block.title) and not para.cell then
      local name = plain(para.text)
      if block.heading and doc and doc.ctx.numbered then -- (with its number, as shown)
        count[block.heading] = (count[block.heading] or 0) + 1
        for deeper = block.heading + 1, 9 do
          count[deeper] = nil
        end
        local number = {}
        for level = 1, block.heading do
          number[level] = count[level] or 0
        end
        name = table.concat(number, ".") .. " " .. name
      end
      list[#list + 1] = { name = name, level = block.heading or 0, first = para.first - 1, line_end = para.last - 1 }
    end
  end
  local last_row = api.nvim_buf_line_count(buf) - 1
  for k, h in ipairs(list) do -- (a section goes on until the next heading of its level or above)
    h.last = last_row
    for j = k + 1, #list do
      if list[j].level <= h.level then
        h.last = list[j].first - 1
        break
      end
    end
  end
  return list
end

local function range(first, last, buf)
  local line = api.nvim_buf_get_lines(buf, last, last + 1, false)[1] or ""
  return { start = { line = first, character = 0 }, ["end"] = { line = last, character = #line } }
end

local function symbols(buf)
  local top, stack = {}, {}
  for _, h in ipairs(headings(buf)) do
    local symbol = {
      name = h.name ~= "" and h.name or "(untitled)",
      kind = h.level == 0 and vim.lsp.protocol.SymbolKind.File or vim.lsp.protocol.SymbolKind.String,
      range = range(h.first, math.max(h.last, h.line_end), buf),
      selectionRange = range(h.first, h.line_end, buf),
      children = {},
    }
    while #stack > 0 and stack[#stack].level >= h.level do
      table.remove(stack)
    end
    table.insert(#stack > 0 and stack[#stack].symbol.children or top, symbol)
    stack[#stack + 1] = { level = h.level, symbol = symbol }
  end
  return top
end

local function folds(buf)
  local out = {}
  for _, h in ipairs(headings(buf)) do
    local last = h.last
    while last > h.first and (api.nvim_buf_get_lines(buf, last, last + 1, false)[1] or "") == "" do
      last = last - 1 -- (not the blank lines before the next heading)
    end
    if last > h.first then
      out[#out + 1] = { startLine = h.first, endLine = last, kind = "region" }
    end
  end
  return out
end

local function covers(m, row, col)
  local d = m[4]
  return (m[2] < row or m[2] == row and m[3] <= col) and (d.end_row > row or d.end_row == row and d.end_col >= col)
end

local function hover(buf, pos)
  local doc, out = neo().docs[buf], {}
  if not doc then
    return nil
  end
  local row, col = pos.line, pos.character
  for _, m in ipairs(api.nvim_buf_get_extmarks(buf, neo().ns.comment, { row, 0 }, { row, -1 }, { details = true, overlap = true })) do
    local c = doc.comment[m[1]]
    if c and covers(m, row, col) then
      out[#out + 1] = ("**%s** · %s%s\n\n%s"):format(c.author, c.date:sub(1, 10), c.resolved and " · resolved" or "", c.text)
      for _, r in ipairs(c.replies) do
        out[#out + 1] = ("↳ **%s** · %s\n\n%s"):format(r.author, r.date:sub(1, 10), r.text)
      end
    end
  end
  for _, m in ipairs(api.nvim_buf_get_extmarks(buf, neo().ns.change, { row, 0 }, { row, -1 }, { details = true, overlap = true })) do
    local ch = doc.change[m[1]]
    if ch and covers(m, row, col) then
      local kind = ({ insert = "inserted", delete = "deleted", format = "formatting changed" })[ch.type]
      kind = ch.moved and (ch.type == "insert" and "moved here" or "moved away") or kind
      out[#out + 1] = ("*%s* by **%s** · %s"):format(kind, ch.author, (ch.date or ""):sub(1, 10))
    end
  end
  return #out > 0 and { contents = { kind = "markdown", value = table.concat(out, "\n\n---\n\n") } } or nil
end

-- The server: requests are answered at once, from the buffer.
function M.server(dispatchers)
  local closing, id = false, 0
  local handlers = {
    initialize = function()
      return {
        capabilities = {
          positionEncoding = "utf-8", -- (columns are bytes, as in Neovim)
          textDocumentSync = 0, -- (it reads the buffer: nothing to send)
          documentSymbolProvider = true,
          foldingRangeProvider = true,
          hoverProvider = true,
        },
        serverInfo = { name = "neoffice" },
      }
    end,
    ["textDocument/documentSymbol"] = function(params) return symbols(vim.uri_to_bufnr(params.textDocument.uri)) end,
    ["textDocument/foldingRange"] = function(params) return folds(vim.uri_to_bufnr(params.textDocument.uri)) end,
    ["textDocument/hover"] = function(params) return hover(vim.uri_to_bufnr(params.textDocument.uri), params.position) end,
    shutdown = function() return vim.NIL end,
  }
  return {
    request = function(method, params, callback)
      id = id + 1
      local handler = handlers[method]
      local ok, result = pcall(handler or error, params)
      vim.schedule(function()
        if handler and ok then
          callback(nil, result)
        else
          callback({ code = -32601, message = handler and tostring(result) or ("not supported: " .. method) }, nil)
        end
      end)
      return true, id
    end,
    notify = function(method)
      if method == "exit" then
        dispatchers.on_exit(0, 15)
      end
      return true
    end,
    is_closing = function() return closing end,
    terminate = function() closing = true end,
  }
end

function M.start(buf)
  if neo().config.lsp == false then
    return
  end
  vim.lsp.start({ name = "neoffice", cmd = M.server, root_dir = vim.fs.dirname(api.nvim_buf_get_name(buf)) }, { bufnr = buf })
end

return M
