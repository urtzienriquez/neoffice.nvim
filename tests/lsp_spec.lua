-- The language server for documents: headings as symbols, sections as folds, hover on comments
-- and changes.
local H = dofile("tests/helpers.lua")
local neo = require("neoffice")
local api = vim.api

local function request(method, params)
  local buf = api.nvim_get_current_buf()
  assert(vim.wait(2000, function() return #vim.lsp.get_clients({ bufnr = buf, name = "neoffice_ls" }) > 0 end))
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
    assert.equals(0, #vim.lsp.get_clients({ bufnr = buf, name = "neoffice_ls" }))
  end)
end)

describe("the language server, on attributes in braces", function()
  local function diagnostics(buf)
    assert(vim.wait(2000, function() return #vim.lsp.get_clients({ bufnr = buf, name = "neoffice_ls" }) > 0 end))
    local ns = vim.lsp.diagnostic.get_namespace(vim.lsp.get_clients({ bufnr = buf, name = "neoffice_ls" })[1].id)
    local out
    vim.wait(2000, function()
      out = vim.diagnostic.get(buf, { namespace = ns })
      return #out > 0
    end)
    return out
  end

  it("hints how to write them (name = value, separated by commas), and writes them so", function()
    local buf = H.open(H.copy("rich.docx"))
    local r = H.row(12)
    api.nvim_buf_set_lines(0, r - 1, r - 1, false, { "{grid,width=full  tall}", "| a | b |", "|---|---|", "| 1 | 2 |", "" })
    assert(vim.wait(2000, function() return #vim.lsp.get_clients({ bufnr = buf, name = "neoffice_ls" }) > 0 end))
    local found = diagnostics(buf)
    assert.equals(1, #found)
    assert.equals(vim.diagnostic.severity.HINT, found[1].severity)
    assert.equals("write it as {rules = grid, width = full, height = tall}", found[1].message)
    local actions = request("textDocument/codeAction", {
      range = { start = { line = r - 1, character = 0 }, ["end"] = { line = r - 1, character = 0 } },
      context = { diagnostics = {} },
    })
    assert.equals(1, #actions)
    vim.lsp.util.apply_workspace_edit(actions[1].edit, "utf-8")
    assert.equals("{rules = grid, width = full, height = tall}", H.lines()[r])
  end)

  it("says when a table's option isn't one, before writing", function()
    local buf = H.open(H.copy("rich.docx"))
    local r = H.row(12)
    api.nvim_buf_set_lines(0, r - 1, r - 1, false, { "{width = dotted}", "| a | b |", "|---|---|", "| 1 | 2 |", "" })
    local found = diagnostics(buf)
    assert.equals(vim.diagnostic.severity.ERROR, found[1].severity)
    assert.truthy(found[1].message:find("not a table option: width = dotted", 1, true), found[1].message)
  end)

  it("hints how to write an image's width", function()
    local buf = H.open(H.copy("rich.docx"))
    local r = H.row(5)
    -- ({width=: no spaces)
    api.nvim_buf_set_lines(0, r - 1, r, false, { (H.lines()[r]:gsub("{width = ", "{width" .. "=")) })
    local found = diagnostics(buf)
    assert.truthy(found[1].message:find("write it as {width = ", 1, true))
  end)
end)

describe("formatting", function()
  local function format(range)
    local buf = api.nvim_get_current_buf()
    assert(vim.wait(2000, function() return #vim.lsp.get_clients({ bufnr = buf, name = "neoffice_ls" }) > 0 end))
    if range then
      vim.lsp.buf.format({ name = "neoffice_ls", range = range })
    else
      vim.lsp.buf.format({ name = "neoffice_ls" })
    end
  end

  it("aligns tables, adds their line of dashes, and tidies attributes; a row with fewer cells keeps them (merged)", function()
    H.open(H.copy("rich.docx"))
    local r = H.row(12)
    api.nvim_buf_set_lines(0, r - 1, r - 1, false, { "{grid,width=full}", "| one | **two** | three |", "| 1 | 2 | 3 |", "|4|5", "" })
    format()
    assert.same({
      "{rules = grid, width = full}",
      "| one | **two** | three |",
      "|----:|--------:|------:|", -- (columns of numbers aligned right, as saving would)
      "|   1 |       2 |     3 |",
      "|   4 |       5 |",
    }, vim.list_slice(H.lines(), r, r + 4))
    local before = H.lines()
    format() -- (formatted: nothing more to do)
    assert.same(before, H.lines())
  end)

  it("keeps the formatting, comments and changes of the cells' text", function()
    H.open(H.copy("rich.docx"))
    H.cursor(11, "two")
    H.input = "On a cell"
    require("neoffice.comments").add()
    local r = H.row(10)
    H.replace(10, "cell one", "cell **one**")
    -- (unaligned, as typing leaves it: spaces taken out between the cells, not the text)
    local line = H.lines()[r]
    for _, sp in ipairs(vim.fn.reverse(require("neoffice.markup").cell_spans(line))) do
      api.nvim_buf_set_text(0, r - 1, sp.e, r - 1, sp.e + 1, {})
      api.nvim_buf_set_text(0, r - 1, sp.s - 2, r - 1, sp.s - 1, {})
    end
    assert.equals("|cell **one**|cell two|", H.lines()[r])
    format()
    assert.equals("| cell **one** | cell two |", H.lines()[r])
    local ns = neo.ns
    local m = vim.iter(api.nvim_buf_get_extmarks(0, ns.comment, 0, -1, { details = true })):find(function(m)
      return H.doc().comment[m[1]].text == "On a cell"
    end)
    assert.equals("two", H.lines()[m[2] + 1]:sub(m[3] + 1, m[3] + 3)) -- (the comment, where it was: on "two")
    H.reopen()
    assert.equals("| cell **one** | cell two |", H.compact(H.lines()[H.row(10)]))
  end)

  it("formats only the tables in a range", function()
    H.open(H.copy("rich.docx"))
    local r = H.row(12)
    api.nvim_buf_set_lines(0, r - 1, r - 1, false, { "|a|b|", "", "|c|d|", "" })
    format({ start = { r + 2, 0 }, ["end"] = { r + 2, 0 } })
    assert.same({ "|a|b|", "", "| c   | d   |", "|-----|-----|" }, vim.list_slice(H.lines(), r, r + 3))
  end)

  it("tidies an image's width, inside its braces", function()
    H.open(H.copy("rich.docx"))
    local r = H.row(5)
    api.nvim_buf_set_lines(0, r - 1, r, false, { (H.lines()[r]:gsub("{width = ", "{ width" .. "=")) })
    format()
    assert.truthy(H.lines()[r]:find("){width = 1.4cm} here", 1, true), H.lines()[r])
  end)
end)

describe("completion", function()
  local function labels(row, line)
    H.open(H.copy("rich.docx"))
    local r = H.row(12)
    api.nvim_buf_set_lines(0, r - 1, r - 1, false, { line, "| a | b |", "|---|---|", "| 1 | 2 |", "" })
    local result = request("textDocument/completion", { position = { line = r - 1, character = #line } })
    local out = vim.tbl_map(function(i) return i.label end, result.items)
    table.sort(out)
    return out, result.items
  end

  it("offers a table's options, those not given yet, with their values", function()
    local got = labels(nil, "{rules = grid, ")
    assert.is_false(vim.tbl_contains(got, "rules = grid"))
    assert.is_true(vim.tbl_contains(got, "width = full"))
    assert.is_true(vim.tbl_contains(got, "height = tall"))
    assert.is_true(vim.tbl_contains(got, "position = left"))
  end)

  it("offers an option's values after its name", function()
    local got, items = labels(nil, "{height = ")
    assert.same({ "compact", "normal", "tall" }, got)
    assert.truthy(items[1].documentation)
  end)

  it("offers highlight colours", function()
    H.open(H.copy("rich.docx"))
    local r = H.row(12)
    api.nvim_buf_set_lines(0, r - 1, r - 1, false, { "Some ==text=={gr", "" })
    local result = request("textDocument/completion", { position = { line = r - 1, character = #"Some ==text=={gr" } })
    local got = vim.tbl_map(function(i) return i.label end, result.items)
    assert.is_true(vim.tbl_contains(got, "green"))
    assert.equals("green}", vim.iter(result.items):find(function(i) return i.label == "green" end).textEdit.newText)
  end)
end)

describe("the language server, on highlights and captions", function()
  it("hints a highlight colour that isn't one, and tidies one written otherwise", function()
    local buf = H.open(H.copy("rich.docx"))
    local r = H.row(12)
    api.nvim_buf_set_lines(0, r - 1, r - 1, false, { "Some ==text=={gren} and ==more=={color = Blue}.", "" })
    assert(vim.wait(2000, function() return #vim.lsp.get_clients({ bufnr = buf, name = "neoffice_ls" }) > 0 end))
    local ns = vim.lsp.diagnostic.get_namespace(vim.lsp.get_clients({ bufnr = buf, name = "neoffice_ls" })[1].id)
    local found
    vim.wait(2000, function()
      found = vim.diagnostic.get(buf, { namespace = ns })
      return #found >= 2
    end)
    local messages = vim.tbl_map(function(d) return d.message end, found)
    table.sort(messages)
    assert.truthy(messages[1]:find("not a highlight colour: gren", 1, true))
    assert.equals("write it as {blue}", messages[2])
    vim.lsp.buf.format({ name = "neoffice_ls" })
    assert.equals("Some ==text=={gren} and ==more=={blue}.", H.lines()[r])
  end)

  it("lists captions with the headings, numbered, in their sections", function()
    H.open(H.copy("rich.docx"))
    local r = H.row(2)
    api.nvim_buf_set_lines(0, r, r, false, { "", "Table: Sprint speeds.", "", "Figure: A lizard." })
    assert.same({ "Introduction", "  Table 1: Sprint speeds.", "  Figure 1: A lizard." }, outline(request("textDocument/documentSymbol")))
  end)
end)

describe("snippets", function()
  local function complete(text)
    local buf = api.nvim_get_current_buf()
    assert(vim.wait(2000, function() return #vim.lsp.get_clients({ bufnr = buf, name = "neoffice_ls" }) > 0 end))
    local n = api.nvim_buf_line_count(0)
    api.nvim_buf_set_lines(0, n, n, false, { "", text })
    local params = { textDocument = { uri = vim.uri_from_bufnr(buf) }, position = { line = n + 1, character = #text } }
    local items = {}
    for _, r in pairs(vim.lsp.buf_request_sync(buf, "textDocument/completion", params, 2000) or {}) do
      items = r.result and (r.result.items or r.result) or items
    end
    return items, n + 2
  end
  local function expand(text, name)
    local items, row = complete(text)
    local item = vim.iter(items):find(function(i) return i.label == name end)
    assert.truthy(item, name)
    local e = item.textEdit
    api.nvim_buf_set_text(0, row - 1, e.range.start.character, row - 1, e.range["end"].character, { "" })
    vim.cmd("startinsert")
    api.nvim_win_set_cursor(0, { row, e.range.start.character })
    vim.snippet.expand(e.newText)
    vim.snippet.stop()
    vim.cmd("stopinsert")
    return row
  end

  it("are offered for blocks at a line's start, for text after a space, and nowhere else", function()
    H.open(H.copy("rich.docx"))
    local labels = function(items) return vim.tbl_map(function(i) return i.label end, items) end
    assert.same({ "table" }, vim.tbl_filter(function(l) return l == "table" end, labels((complete("tab")))))
    assert.truthy(vim.tbl_contains(labels((complete("Some words foot"))), "footnote"))
    assert.is_false(vim.tbl_contains(labels((complete("Some words tab"))), "table"))
    assert.same({}, (complete("{rules = gr")))
  end)

  it("write neoffice's markup, which saves as a table with its caption, a figure, a section, a footnote", function()
    local dir = vim.fn.tempname()
    vim.fn.mkdir(dir .. "/figures", "p")
    assert(vim.uv.fs_copyfile(H.FIXTURES .. "rich.docx", dir .. "/s.docx"))
    assert(vim.uv.fs_copyfile(H.FIXTURES .. "dot.png", dir .. "/figures/file.png"))
    require("neoffice").config.page = true
    H.open(dir .. "/s.docx")
    for _, name in ipairs({ "table", "figure", "landscape" }) do
      expand(name:sub(1, 3), name)
    end
    local row = expand("A word foot", "footnote")
    assert.equals("A word^[Note.]", H.lines()[row])
    H.reopen()
    local lines = H.lines()
    for _, want in ipairs({ "Table: Caption.", "Figure: Caption.", "\\newpage {orientation = landscape}", "\\newpage {orientation = portrait}", "A word^[Note.]" }) do
      assert.truthy(vim.tbl_contains(lines, want), want .. "\n" .. table.concat(lines, "\n"))
    end
    require("neoffice").config.page = false
  end)
end)
