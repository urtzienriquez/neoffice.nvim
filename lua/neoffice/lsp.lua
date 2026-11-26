-- neoffice_ls: a language server for documents, running in Neovim itself (vim.lsp.start with a
-- function as its command): nothing to install. It reads the buffer as it is, and answers:
--   textDocument/documentSymbol  the title and the headings, nested (gO, pickers, outlines)
--   textDocument/foldingRange    each heading's section
--   textDocument/hover           the comment, or the tracked change, at the cursor
--   diagnostics                  attributes in braces (a table's options, an image's width): an
--                                error if one isn't any, a hint if not written "name = value, ..."
--   textDocument/codeAction      for those hints: write them so
--   textDocument/formatting      tables aligned (with their line of dashes), attributes tidy
--   (and rangeFormatting)        (see format.lua)
--   textDocument/completion      a table's options and their values, an image's width, a
--                                highlight's colour
local api = vim.api
local M = { NAME = "neoffice_ls" }

local function neo()
  return require("neoffice")
end

local function plain(text) -- (a paragraph's text, without its marks)
  local _, chars = require("neoffice.markup").parse(text)
  return vim.trim(table.concat(vim.tbl_map(function(c) return c.ch end, chars)):gsub("%s+", " "))
end

-- The title and the headings: { name, level (0: the title), first, last (0-based rows: its section) };
-- with `captions`, the captions too ("Table 1: ...", level: math.huge, in the section they are in).
local function headings(buf, captions)
  local doc, list, count = neo().docs[buf], {}, {}
  local blocks = neo().blocks(buf)
  for _, para in ipairs(blocks) do
    local block = require("neoffice.markup").block(para.text)
    if captions and block.caption and not para.cell then
      local key = "caption." .. block.caption
      count[key] = (count[key] or 0) + 1
      local label = require("neoffice.markup").captions()[block.caption]
      list[#list + 1] = { name = label .. " " .. count[key] .. ": " .. plain(para.text), level = math.huge, first = para.first - 1, line_end = para.last - 1, caption = true }
    end
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
    h.last = h.caption and h.line_end or last_row
    for j = k + 1, h.caption and 0 or #list do -- (a caption: only its own lines)
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
  for _, h in ipairs(headings(buf, true)) do
    local symbol = {
      name = h.name ~= "" and h.name or "(untitled)",
      kind = h.caption and vim.lsp.protocol.SymbolKind.Object or h.level == 0 and vim.lsp.protocol.SymbolKind.File or vim.lsp.protocol.SymbolKind.String,
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

-- Attributes in braces: a table's options line ("{width = full, height = tall}") and an image's
-- width ("{width = 12cm}"). An error if a table's option isn't one; a hint (with the text to write
-- instead: data.new) if they aren't written "name = value", separated by commas.
local function attribute_diagnostics(buf)
  local markup, tables, out = require("neoffice.markup"), require("neoffice.tables"), {}
  local lines = api.nvim_buf_get_lines(buf, 0, -1, false)
  local function add(r, c1, c2, severity, message, new)
    out[#out + 1] = {
      range = { start = { line = r - 1, character = c1 }, ["end"] = { line = r - 1, character = c2 } },
      severity = severity, message = message, source = "neoffice_ls", data = new and { new = new } or nil,
    }
  end
  for r, line in ipairs(lines) do
    local inside = line:sub(1, 1) == "{" and markup.table_options(line)
    if inside and (lines[r + 1] or ""):match("^|") then
      local ok, err = pcall(tables.parse, inside)
      local _, tidy = markup.attributes(inside, tables.name_of)
      if not ok then
        add(r, 0, #line, vim.lsp.protocol.DiagnosticSeverity.Error, (tostring(err):gsub("^.-:%d+: ", "")))
      elseif "{" .. tidy .. "}" ~= line then
        add(r, 0, #line, vim.lsp.protocol.DiagnosticSeverity.Hint, "write it as {" .. tidy .. "}", "{" .. tidy .. "}")
      end
    end
    for s, inside, e in line:gmatch("==(){([^{}]*)}()") do -- a highlight's colour
      local name = inside:match("^%s*colou?r%s*=%s*([#%w]+)%s*$") or inside:match("^%s*([#%w]+)%s*$")
      local color = name and markup.color(name)
      if not color then
        add(r, s - 1, e - 1, vim.lsp.protocol.DiagnosticSeverity.Hint,
          "not a highlight colour: " .. inside .. " (they are " .. table.concat(vim.fn.sort(vim.tbl_keys(markup.COLORS)), ", ") .. ", or #rrggbb)")
      elseif color ~= inside then
        add(r, s - 1, e - 1, vim.lsp.protocol.DiagnosticSeverity.Hint, "write it as {" .. color .. "}", "{" .. color .. "}")
      end
    end
    local p = line:find("![", 1, true)
    while p do
      local _, _, after, width, attrs = markup.image_at(line, p)
      local tidy = width and "{width = " .. width .. "}"
      if attrs and line:sub(attrs.s, attrs.e) ~= tidy then
        add(r, attrs.s - 1, attrs.e, vim.lsp.protocol.DiagnosticSeverity.Hint, "write it as " .. tidy, tidy)
      end
      p = line:find("![", math.max(after or 0, p + 2), true)
    end
  end
  return out
end

-- Completion: inside a table's options line ("{" above a table), its options ("rules = grid") or,
-- after "name =", that option's values; an image's width after "){width ="; a highlight's colour
-- after "=={".
local KIND = vim.lsp.protocol.CompletionItemKind
local function completion(buf, pos)
  local markup, tables = require("neoffice.markup"), require("neoffice.tables")
  local lines = api.nvim_buf_get_lines(buf, pos.line, pos.line + 2, false)
  local line = lines[1] or ""
  local before = line:sub(1, pos.character)
  local items = {}
  local function add(label, start, kind, doc, insert)
    items[#items + 1] = {
      label = label, kind = kind, documentation = doc, filterText = label,
      textEdit = { newText = insert or label, range = { start = { line = pos.line, character = start }, ["end"] = pos } },
    }
  end
  local color = before:match("==%{([#%w]*)$")
  if color then
    local start = #before - #color
    for name, hex in pairs(markup.COLORS) do
      add(name, start, KIND.Color, "#" .. hex, name .. "}")
    end
    return items
  end
  local width = before:match("%)%{%s*width%s*=%s*([^%s}]*)$")
  if width then
    for _, w in ipairs({ "100%", "75%", "50%", "8cm", "12cm", "16cm" }) do
      add(w, #before - #width, KIND.Value, "the image's width: a length (cm, mm, in) or a percentage of the text's width")
    end
    return items
  end
  local inside = line:sub(1, 1) == "{" and not before:find("}", 1, true) and (lines[2] or ""):match("^|") and before:sub(2)
  if not inside then
    return items
  end
  local used = {}
  for _, a in ipairs((markup.attributes(inside))) do
    used[a.name or tables.name_of(a.value) or ""] = true
  end
  local item = inside:match("([^,;]*)$") -- (the one being written)
  local name, value = item:match("^%s*([%a_]+)%s*=%s*(%S*)$")
  local DESCRIBE = tables.DESCRIBE
  if name then -- its values
    local key = name:lower()
    key = key == "borders" and "rules" or key == "align" and "position" or key
    for _, v in ipairs(tables.WORDS[key] or {}) do
      add(v, #before - #value, KIND.EnumMember, DESCRIBE[key .. "." .. v])
    end
    if key == "width" then
      for _, v in ipairs({ "80%", "60%", "12cm" }) do
        add(v, #before - #value, KIND.Value, DESCRIBE["width.length"])
      end
    end
  else -- the options not given yet, each with its values
    local typed = item:match("^%s*(.-)$")
    for _, key in ipairs(tables.ORDER) do
      if not used[key] or typed ~= "" and tables.name_of(typed) == key then
        for _, v in ipairs(tables.WORDS[key]) do
          add(key .. " = " .. v, #before - #typed, KIND.Property, DESCRIBE[key .. "." .. v])
        end
      end
    end
  end
  return items
end

local function code_actions(buf, params)
  local uri, out = params.textDocument.uri, {}
  for _, d in ipairs(attribute_diagnostics(buf)) do
    if d.data and d.range.start.line >= params.range.start.line and d.range.start.line <= params.range["end"].line then
      out[#out + 1] = {
        title = d.message:gsub("^write", "Write"), kind = "quickfix", diagnostics = { d }, isPreferred = true,
        edit = { changes = { [uri] = { { range = d.range, newText = d.data.new } } } },
      }
    end
  end
  return out
end

-- The server: requests are answered at once, from the buffer.
function M.server(dispatchers)
  local closing, id, pending = false, 0, {}
  local handlers = {
    initialize = function()
      return {
        capabilities = {
          positionEncoding = "utf-8", -- (columns are bytes, as in Neovim)
          -- (it reads the buffer: changes only say when to look at it again, for diagnostics)
          textDocumentSync = { openClose = true, change = 2 },
          codeActionProvider = true,
          documentFormattingProvider = true,
          documentRangeFormattingProvider = true,
          completionProvider = { triggerCharacters = { "{", "=", ",", " " } },
          documentSymbolProvider = true,
          foldingRangeProvider = true,
          hoverProvider = true,
        },
        serverInfo = { name = "neoffice_ls" },
      }
    end,
    ["textDocument/documentSymbol"] = function(params) return symbols(vim.uri_to_bufnr(params.textDocument.uri)) end,
    ["textDocument/foldingRange"] = function(params) return folds(vim.uri_to_bufnr(params.textDocument.uri)) end,
    ["textDocument/hover"] = function(params) return hover(vim.uri_to_bufnr(params.textDocument.uri), params.position) end,
    ["textDocument/codeAction"] = function(params) return code_actions(vim.uri_to_bufnr(params.textDocument.uri), params) end,
    ["textDocument/formatting"] = function(params)
      return require("neoffice.format").edits(vim.uri_to_bufnr(params.textDocument.uri))
    end,
    ["textDocument/rangeFormatting"] = function(params)
      return require("neoffice.format").edits(vim.uri_to_bufnr(params.textDocument.uri), params.range.start.line, params.range["end"].line)
    end,
    ["textDocument/completion"] = function(params)
      return { isIncomplete = false, items = completion(vim.uri_to_bufnr(params.textDocument.uri), params.position) }
    end,
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
    notify = function(method, params)
      if method == "exit" then
        dispatchers.on_exit(0, 15)
      elseif method == "textDocument/didOpen" or method == "textDocument/didChange" then
        local uri = params.textDocument.uri
        if not pending[uri] then -- (once typing pauses)
          pending[uri] = true
          vim.defer_fn(function()
            pending[uri] = nil
            local buf = vim.uri_to_bufnr(uri)
            if not closing and api.nvim_buf_is_loaded(buf) then
              local ok, diagnostics = pcall(attribute_diagnostics, buf)
              dispatchers.notification("textDocument/publishDiagnostics", { uri = uri, diagnostics = ok and diagnostics or {} })
            end
          end, 150)
        end
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
  vim.lsp.start({ name = M.NAME, cmd = M.server, root_dir = vim.fs.dirname(api.nvim_buf_get_name(buf)) }, { bufnr = buf })
end

return M
