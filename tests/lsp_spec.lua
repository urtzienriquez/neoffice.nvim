-- The language server for documents: headings as symbols, sections as folds, hover on comments
-- and changes.
local H = dofile("tests/helpers.lua")
local neo = require("neoffice")
local api = vim.api

local function request(method, params)
  local buf = api.nvim_get_current_buf()
  assert(vim.wait(2000, function() return #vim.lsp.get_clients({ bufnr = buf, name = "neoffice" }) > 0 end))
  params = vim.tbl_extend("force", { textDocument = { uri = vim.uri_from_bufnr(buf) } }, params or {})
  local result
  for _, r in pairs(vim.lsp.buf_request_sync(buf, method, params, 2000) or {}) do
    result = r.result
  end
  return result
end

local function outline(list, out, indent)
  out, indent = out or {}, indent or ""
  for _, s in ipairs(list or {}) do
    out[#out + 1] = indent .. s.name
    outline(s.children, out, indent .. "  ")
  end
  return out
end

describe("the language server", function()
  it("gives the title and headings as symbols, nested, with their sections", function()
    H.open(H.copy("rich.docx"))
    api.nvim_buf_set_lines(0, 0, 0, false, { "% The **title**", "", "## A subsection", "", "Text.", "" })
    local symbols = request("textDocument/documentSymbol")
    assert.same({ "The title", "  A subsection", "  Introduction" }, outline(symbols))
    assert.equals(H.row(4) - 1, symbols[1].children[2].range.start.line) -- (Introduction: its section, to the end)
    assert.equals(api.nvim_buf_line_count(0) - 1, symbols[1].children[2].range["end"].line)
  end)

  it("numbers headings as they are shown, and folds each section", function()
    H.open(H.copy("rich.docx"))
    vim.cmd("Neoffice numbers on")
    api.nvim_buf_set_lines(0, -1, -1, false, { "", "# Second", "", "## Second, first", "", "Text." })
    assert.same({ "1 Introduction", "2 Second", "  2.1 Second, first" }, outline(request("textDocument/documentSymbol")))
    local folds = vim.tbl_map(function(f) return f.startLine end, request("textDocument/foldingRange"))
    assert.same({ 0, H.row(13) - 1, H.row(14) - 1 }, folds)
  end)

  it("shows the comment, or the change, under the cursor", function()
    H.open(H.copy("rich.docx"))
    local hover = request("textDocument/hover", { position = { line = H.row(3) - 1, character = 3 } })
    assert.truthy(hover.contents.value:find("Alice", 1, true) and hover.contents.value:find("A comment", 1, true))
    hover = request("textDocument/hover", { position = { line = H.row(4) - 1, character = 12 } })
    assert.truthy(hover.contents.value:find("inserted", 1, true) and hover.contents.value:find("Bob", 1, true))
    assert.is_nil(request("textDocument/hover", { position = { line = H.row(12) - 1, character = 2 } }))
  end)

  it("doesn't start with lsp = false", function()
    neo.config.lsp = false
    local buf = H.open(H.copy("rich.odt"))
    neo.config.lsp = true
    vim.wait(200)
    assert.equals(0, #vim.lsp.get_clients({ bufnr = buf, name = "neoffice" }))
  end)
end)
