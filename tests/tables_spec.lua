-- Tables as pipe tables: editing cells, adding and removing rows, columns and tables.
local H = dofile("tests/helpers.lua")
local zip = require("neoffice.zip")
local api = vim.api

local function table_rows() -- the table's lines (without the line of dashes), unaligned
  return vim.tbl_filter(function(l) return l:match("^|") and not l:match("^|:?%-") end, H.compact_lines())
end
local function dashes() -- the lines of dashes, unaligned
  return vim.tbl_filter(function(l) return l:match("^|:?%-") ~= nil end, H.compact_lines())
end
local function set_dashes(line, n) -- (of the n-th table)
  local k = 0
  for r, l in ipairs(H.lines()) do
    if l:match("^|:?%-") then
      k = k + 1
      if k == (n or 1) then
        return api.nvim_buf_set_lines(0, r - 1, r, false, { line })
      end
    end
  end
end
local function options() -- the options line above the first table, if it has one
  local lines = H.lines()
  for r, l in ipairs(lines) do
    if l:match("^|") then
      return (lines[r - 1] or ""):match("^{.*}$")
    end
  end
end
local function set_options(text) -- (above the first table; nil: none)
  local lines = H.lines()
  for r, l in ipairs(lines) do
    if l:match("^|") then
      local has = (lines[r - 1] or ""):match("^{.*}$") and 1 or 0
      return api.nvim_buf_set_lines(0, r - 1 - has, r - 1, false, text and { text } or {})
    end
  end
end

local function set_table(rows) -- replace the table's lines (a line of dashes goes after the first)
  local lines = H.lines()
  local first, last
  for k, l in ipairs(lines) do
    if l:match("^|") then
      first, last = first or k, k
    end
  end
  local new = { rows[1], "|" .. ("---|"):rep(select(2, rows[1]:gsub("|", "")) - 1) }
  vim.list_extend(new, vim.list_slice(rows, 2))
  api.nvim_buf_set_lines(0, first - 1, last, false, new)
end

for _, name in ipairs({ "rich.docx", "rich.odt" }) do
  local ext = name:match("%w+$")
  local ROW, CELL = ext == "docx" and "<w:tr[ >]" or "<table:table%-row[ >]", ext == "docx" and "<w:tc[ >]" or "<table:table%-cell[ >]"
  local function count(pattern)
    return select(2, zip.read(H.doc().path, H.MAIN[ext]):gsub(pattern, ""))
  end

  describe(name, function()
    local path
    before_each(function()
      path = H.copy(name)
      H.open(path)
    end)

    it("shows a table as a pipe table, and its look above it", function()
      local lines = H.compact_lines()
      local r = vim.fn.index(lines, "| A | B |") + 1
      assert.same({ "| A | B |", "|---|---|", "| cell one | cell two |", "" }, vim.list_slice(lines, r, r + 3))
      -- (pandoc's tables: a line under their first row, on the left; how wide, said in the .odt)
      assert.equals(ext == "docx" and "{rules = header, position = left}" or "{rules = header, width = 14cm, position = left}", lines[r - 1])
    end)

    it("aligns a table's columns in the buffer, unless it is too wide", function()
      local r = vim.fn.index(H.lines(), "| A        | B        |") + 1
      assert.truthy(r > 0, table.concat(H.lines(), "\n"))
      assert.same({ "|----------|----------|", "| cell one | cell two |" }, vim.list_slice(H.lines(), r + 1, r + 2))
      set_table({ "| A | B |", "| cell one | " .. ("long "):rep(30) .. "|" })
      H.reopen()
      assert.truthy(vim.tbl_contains(H.lines(), "| A | B |")) -- (wider than 120 columns: not aligned)
    end)

    it("edits a cell, and only its paragraph", function()
      local before = H.paragraph_xml(path)
      H.replace(10, "cell one", "first cell")
      H.reopen()
      assert.same({ "| A | B |", "| first cell | cell two |" }, table_rows())
      assert.equals(1, H.differing(before, H.paragraph_xml(path)))
    end)

    it("adds rows (at the end and in between)", function()
      set_table({ "| A | B |", "| x | y |", "| cell one | cell two |", "| p | q |" })
      H.reopen()
      assert.same({ "| A | B |", "| x | y |", "| cell one | cell two |", "| p | q |" }, table_rows())
      assert.equals(4, count(ROW))
    end)

    it("removes rows", function()
      set_table({ "| A | B |" })
      H.reopen()
      assert.same({ "| A | B |" }, table_rows())
      assert.equals(1, count(ROW))
      assert.equals("Last paragraph.", H.paras()[#H.paras()])
    end)

    it("adds and removes columns", function()
      set_table({ "| A | B | C |", "| cell one | cell two | three |" })
      H.reopen()
      assert.same({ "| A | B | C |", "| cell one | cell two | three |" }, table_rows())
      assert.equals(6, count(CELL))
      if ext == "docx" then
        assert.equals(3, count("<w:gridCol[ />]"))
      end
      set_table({ "| B | C |", "| cell two | three |" })
      H.reopen()
      assert.same({ "| B | C |", "| cell two | three |" }, table_rows())
      assert.equals(4, count(CELL))
    end)

    it("keeps several paragraphs in a cell, separated by <br>", function()
      H.replace(10, "cell one", "one <br> two")
      H.reopen()
      assert.same({ "| A | B |", "| one <br> two | cell two |" }, table_rows())
      assert.equals(4, count(CELL))
      H.replace(10, "one <br> two", "one")
      H.reopen()
      assert.same({ "| A | B |", "| one | cell two |" }, table_rows())
    end)

    it("escapes | and <br> in cells, and aligns cells", function()
      H.replace(10, "cell one", "a \\| b \\<br>")
      H.replace(11, "cell two", "<|> centered")
      H.reopen()
      assert.same({ "| A | B |", "| a \\| b \\<br> | <|> centered |" }, table_rows())
    end)

    it("typesets a new table: booktabs, widths from its content, compact, centred, numbers right", function()
      local r = H.row(12)
      api.nvim_buf_set_lines(0, r - 1, r - 1, false, { "Table 1. A caption.", "", "| Name | Count |", "|---|---|", "| first | 2 |", "| second | 10 |", "" })
      H.reopen()
      local lines = table.concat(H.compact_lines(), "\n")
      -- (no options line: all the defaults; the column of numbers aligned right)
      assert.truthy(lines:find("Table 1. A caption.\n\n| Name | Count |\n|---|---:|\n| first | 2 |\n| second | 10 |\n\nLast paragraph.", 1, true), lines)
      assert.equals(2, count(ext == "docx" and "<w:tbl>" or "<table:table "))
      local main = zip.read(H.doc().path, H.MAIN[ext])
      if ext == "docx" then
        local tbl = main:match(".*<w:tbl>(.-)</w:tbl>")
        assert.truthy(tbl:find('<w:jc w:val="center"/>', 1, true))
        assert.truthy(tbl:find('<w:top w:val="single" w:sz="8"', 1, true)) -- (thick, above)
        assert.truthy(tbl:find('<w:top w:w="28" w:type="dxa"/>', 1, true)) -- (compact)
        assert.truthy(tbl:find("<w:tblHeader/>", 1, true)) -- (the header on every page)
        assert.truthy(tbl:find("<w:cantSplit/>", 1, true))
        assert.truthy(tbl:find('<w:spacing w:before="0" w:after="0"', 1, true))
        assert.truthy(main:find('<w:keepNext/></w:pPr><w:r><w:t xml:space="preserve">Table 1. A caption.', 1, true)) -- (kept with it)
        local a, b = tbl:match('<w:gridCol w:w="(%d+)"/><w:gridCol w:w="(%d+)"/>')
        assert.truthy(tonumber(a) > tonumber(b)) -- ("second" is wider than "Count")
      else
        assert.truthy(main:find("<table:table%-header%-rows>"))
        assert.truthy(main:find('table:align="center"', 1, true))
        assert.truthy(main:find('fo:padding%-top="0.05cm"'))
        assert.truthy(main:find('fo:keep%-together="always"'))
      end
    end)

    it("changes a table's look with the line above it, and back to the defaults without it", function()
      set_options("{grid, width=full, height = tall}")
      H.reopen()
      assert.equals("{rules = grid, width = full, height = tall}", options())
      if ext == "docx" then
        assert.truthy(zip.read(H.doc().path, H.MAIN[ext]):find('<w:tblW w:w="5000" w:type="pct"/>', 1, true))
      end
      set_options("{position = right, rules=rows, width = 8cm}") -- (the order doesn't matter)
      H.reopen()
      assert.equals("{rules = rows, width = 8cm, position = right}", options())
      set_options("{booktabs, 60%, normal, center}") -- (defaults aren't shown)
      H.reopen()
      assert.equals("{width = 60%, height = normal}", options())
      set_options(nil)
      H.reopen()
      assert.is_nil(options())
      assert.same({ "| A | B |", "| cell one | cell two |" }, table_rows())
    end)

    it("typesets a typeset table again when it changes; a table it didn't typeset only when its line does", function()
      set_options("{rules = grid}")
      H.reopen()
      local function widths() -- (of the first table's columns)
        local main = zip.read(H.doc().path, H.MAIN[ext])
        local out = {}
        if ext == "docx" then
          for w in main:gmatch('<w:gridCol w:w="(%d+)"') do
            out[#out + 1] = tonumber(w)
          end
        else
          for name in main:match("<table:table .-</table:table>"):gmatch('<table:table%-column table:style%-name="([^"]+)"') do
            out[#out + 1] = tonumber(main:match('style:name="' .. name .. '".-style:column%-width="([%d.]+)cm"'))
          end
        end
        return out
      end
      local before = widths()
      set_table({ "| A | B |", "| a much longer first cell than before | cell two |", "| new | row |" })
      H.reopen()
      assert.equals("{rules = grid}", options()) -- (the new row has the table's lines too)
      assert.truthy(widths()[1] > before[1])
      -- pandoc's table (not typeset by neoffice): a new row keeps its look as it is
      H.open(H.copy(name))
      set_table({ "| A | B |", "| cell one | cell two |", "| new | row |" })
      H.reopen()
      assert.equals(ext == "docx" and "{rules = header, position = left}" or "{rules = header, width = 14cm, position = left}", options())
    end)

    it("doesn't save a table option it doesn't know, and says which ones it knows", function()
      local before = H.bytes(path)
      set_options("{rules = dotted}")
      local errors = {}
      local notify = vim.notify
      vim.notify = function(msg) errors[#errors + 1] = msg end
      H.save()
      vim.notify = notify
      assert.equals(before, H.bytes(path))
      assert.truthy(table.concat(errors, "\n"):find("booktabs", 1, true))
    end)

    it("keeps the table Markdown: its options line doesn't break the highlighting", function()
      set_options("{width = full, height = tall}")
      local tree = vim.treesitter.get_parser(0, "markdown"):parse()[1]
      assert.truthy(tree:root():sexpr():find("pipe_table_delimiter_row", 1, true))
    end)

    it("aligns columns from the line of dashes, and a cell its own way", function()
      set_table({ "| A | B |", "| cell one | cell two |", "| x | 10 |" })
      set_dashes("|:---:|---:|")
      H.reopen()
      assert.equals("|:---:|---:|", dashes()[1]:match("^%S+"))
      assert.same({ "| A | B |", "| cell one | cell two |", "| x | 10 |" }, table_rows())
      -- every paragraph has it (in the file)
      local main = zip.read(H.doc().path, H.MAIN[ext])
      local right = ext == "docx" and '<w:jc w:val="right"/>' or nil
      if right then
        assert.equals(3, select(2, main:gsub(right, "")))
        assert.equals(3, select(2, main:gsub('<w:jc w:val="center"/>', "")))
      end
      -- a cell aligned otherwise keeps its mark, and the column its alignment
      H.replace(10, "cell one", "<=> cell one")
      H.reopen()
      -- (the first column is no longer all centered: each cell says how it is)
      assert.same({ "| <|> A | B |", "| <=> cell one | cell two |", "| <|> x | 10 |" }, table_rows())
      assert.equals("|---|---:|", dashes()[1]:match("^%S+"))
    end)

    it("deletes a table", function()
      local lines = H.lines()
      for k = #lines, 1, -1 do -- (and the line of dashes)
        if lines[k]:match("^|") then
          api.nvim_buf_set_lines(0, k - 1, k, false, {})
        end
      end
      H.reopen()
      assert.same({}, table_rows())
      assert.equals(0, count(ROW))
    end)

    it("turns a table cell into a paragraph and back without mixing them up", function()
      local r = H.row(12)
      api.nvim_buf_set_lines(0, r - 2, r, false, { "| Last paragraph. |" }) -- (a row of the table)
      H.reopen()
      assert.same({ "| A | B |", "| cell one | cell two |", "| Last paragraph. |" }, table_rows())
    end)

    it("keeps comments and tracked changes in cells in their place", function()
      H.cursor(11, "two")
      require("neoffice.comments").add(nil, "On a cell")
      vim.cmd("Neoffice track on")
      H.replace(10, "one", "uno")
      H.reopen()
      local ns = require("neoffice").ns
      local m = vim.iter(api.nvim_buf_get_extmarks(0, ns.comment, 0, -1, {})):find(function(m)
        return H.doc().comment[m[1]].text == "On a cell"
      end)
      assert.equals("two |", H.lines()[m[2] + 1]:sub(m[3] + 1))
      vim.cmd("Neoffice accept all")
      H.reopen()
      assert.same({ "| A | B |", "| cell uno | cell two |" }, table_rows())
    end)

    if H.has_libreoffice then
      it("writes tables LibreOffice reads", function()
        set_table({ "| A | B | C |", "| cell one | cell two | three |", "| new | row | here |" })
        H.save()
        local ok, text = H.libreoffice_reads(path, "here")
        assert.is_true(ok, text)
        assert.truthy(text:find("three", 1, true), text)
      end)
    end
  end)
end

describe("table typesetting", function()
  local tables = require("neoffice.tables")

  it("reads the words of a table's options, in any order, and says which are wrong", function()
    assert.same({ width = "full", height = "tall" }, tables.parse("width = full, height = tall"))
    assert.same({ width = "full", height = "tall" }, tables.parse("width=full,height=tall"))
    assert.same({ width = "full", height = "tall" }, tables.parse("full tall")) -- (values alone)
    assert.same({ rules = "grid", width = "80%", position = "left" }, tables.parse("left, 80% all"))
    assert.same({ rules = "rows", width = "12cm" }, tables.parse("borders = rows width=12cm"))
    assert.has_error(function() tables.parse("height = full") end)
    assert.has_error(function() tables.parse("dotted") end)
    assert.has_error(function() tables.parse("colour=red") end)
    assert.has_error(function() tables.parse("150%") end)
    assert.equals("", tables.line(tables.defaults()))
    assert.equals("rules = grid, width = full", tables.line(vim.tbl_extend("force", tables.defaults(), { rules = "grid", width = "full" })))
  end)

  it("tells numbers from text", function()
    for _, n in ipairs({ "12", "-3.5", "1,234", "4.3 ± 0.2", "< 0.001", "12%", "98 (12)", "0.5 [0.2, 0.9]", "**7**" }) do
      assert.is_true(tables.is_number(n), n)
    end
    for _, t in ipairs({ "North", "12 lizards", "p < 0.01 here", "A1" }) do
      assert.is_false(tables.is_number(t), t)
    end
    assert.is_nil(tables.is_number("–")) -- (neither: an empty cell)
  end)

  it("makes columns as wide as their content, and the widest give way when they don't fit", function()
    local em = 0.4
    local short = { { { "12" } }, { { "3.5" } } }
    local w = tables.widths(short, "auto", em, 16)
    assert.truthy(w[1] < 2 and w[2] < 2)
    local full = tables.widths(short, "full", em, 16)
    assert.truthy(math.abs(full[1] + full[2] - 16) < 0.01)
    local long = { { { "Site" } }, { { "Mean speed (m/s)" } }, { { ("a long description of a habitat "):rep(4) } } }
    local fit = tables.widths(long, "auto", em, 16)
    assert.truthy(math.abs(fit[1] + fit[2] + fit[3] - 16) < 0.01) -- (as wide as the text)
    assert.same(tables.widths({ long[2] }, "auto", em, 16)[1], fit[2]) -- (the middle one doesn't wrap)
  end)

  it("tells the rules apart, thin and thick lines too", function()
    local function grid(preset, rows, cols)
      local out = {}
      for r = 1, rows do
        out[r] = {}
        for c = 1, cols do
          out[r][c] = tables.cell_borders(preset, r, rows, c, cols)
        end
      end
      return out
    end
    for preset in pairs(tables.BORDERS) do
      assert.equals(preset, tables.borders_of(grid(preset, 4, 3)))
      if preset ~= "none" and preset ~= "header" then
        assert.equals(preset, tables.borders_of(grid(preset, 2, 2)), preset)
      end
    end
  end)
end)

-- A table with merged cells (as forms have: a label, and one cell for its value across the rest).
local MERGED = {
  docx = function(name, text)
    if name ~= "word/document.xml" then
      return text
    end
    local function tc(t, span)
      return "<w:tc><w:tcPr><w:tcW w:w=\"" .. (span and 6000 or 2000) .. "\" w:type=\"dxa\"/>" .. (span and '<w:gridSpan w:val="' .. span .. '"/>' or "")
        .. "</w:tcPr><w:p><w:r><w:t xml:space=\"preserve\">" .. t .. "</w:t></w:r></w:p></w:tc>"
    end
    local tbl = '<w:tbl><w:tblPr><w:tblW w:w="8000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>'
      .. "<w:tr>" .. tc("Name") .. tc(" ", 3) .. "</w:tr>"
      .. "<w:tr>" .. tc("Gender") .. tc(" ") .. tc("Born") .. tc(" ") .. "</w:tr></w:tbl><w:p/>"
    return (text:gsub("<w:body>", "<w:body>" .. tbl, 1))
  end,
  odt = function(name, text)
    if name ~= "content.xml" then
      return text
    end
    local function tc(t, span)
      return "<table:table-cell" .. (span and ' table:number-columns-spanned="' .. span .. '"' or "") .. "><text:p>" .. t .. "</text:p></table:table-cell>"
        .. ("<table:covered-table-cell/>"):rep((span or 1) - 1)
    end
    local tbl = '<table:table table:name="Form"><table:table-column table:number-columns-repeated="4"/>'
      .. "<table:table-row>" .. tc("Name") .. tc(" ", 3) .. "</table:table-row>"
      .. "<table:table-row>" .. tc("Gender") .. tc(" ") .. tc("Born") .. tc(" ") .. "</table:table-row></table:table><text:p/>"
    return (text:gsub("<office:text>", "<office:text>" .. tbl, 1))
  end,
}

for _, ext in ipairs({ "docx", "odt" }) do
  describe(ext .. " a table with merged cells", function()
    local path, row
    before_each(function()
      path = H.copy("rich." .. ext, MERGED[ext])
      H.open(path)
      row = vim.iter(ipairs(H.lines())):find(function(_, l) return l:find("^| Name") end)
      assert.truthy(row, table.concat(H.lines(), "\n"))
    end)
    local function append(r, text) -- (typed at its end: as formatting would add them)
      local line = H.lines()[r]
      api.nvim_buf_set_text(0, r - 1, #line, r - 1, #line, { text })
    end

    -- the merged table's rows: each cell's span (covered cells, in an .odt: "c")
    local function shape()
      local main = zip.read(path, H.MAIN[ext])
      local out = {}
      if ext == "docx" then
        local tbl = main:match("<w:tbl>.-</w:tbl>")
        out.grid = select(2, tbl:gsub("<w:gridCol ", ""))
        for tr in tbl:gmatch("<w:tr>.-</w:tr>") do
          local spans = {}
          for tc in tr:gmatch("<w:tc>.-</w:tc>") do
            spans[#spans + 1] = tc:match('gridSpan w:val="(%d+)"') or "1"
          end
          out[#out + 1] = table.concat(spans, ",")
        end
      else
        local tbl = main:match('<table:table table:name="Form".-</table:table>')
        for tr in tbl:gmatch("<table:table%-row>(.-)</table:table%-row>") do
          local spans = {}
          for kind, attrs in tr:gmatch("<table:(%S+)([^>]*)") do
            if kind == "table-cell" then
              spans[#spans + 1] = attrs:match('number%-columns%-spanned="(%d+)"') or "1"
            elseif kind == "covered-table-cell" or kind == "covered-table-cell/" then
              spans[#spans + 1] = "c"
            end
          end
          out[#out + 1] = table.concat(spans, ",")
        end
      end
      return out
    end

    it("keeps them when a row is padded with empty cells (as another formatter would): nothing changes", function()
      local before = shape()
      append(row, "   |   |")
      H.save()
      assert.same(before, shape())
      H.reopen()
      assert.truthy(H.lines()[row]:match("^| Name%s+|%s+|$"), H.lines()[row])
    end)

    it("takes a column from a merged cell for a cell typed with text, and the table keeps its columns", function()
      append(row, " Middle | M |")
      H.reopen()
      assert.truthy(H.lines()[row]:match("^| Name%s+|%s+| Middle%s+| M%s+|$"), H.lines()[row])
      assert.is_true(require("neoffice.xml").check(zip.read(path, H.MAIN[ext])))
      local got = shape()
      if ext == "docx" then
        assert.same({ grid = 4, "1,1,1,1", "1,1,1,1" }, got)
      else
        assert.same({ "1,1,1,1", "1,1,1,1" }, got)
      end
      if H.has_libreoffice then
        local ok, out = H.libreoffice_reads(path, "Middle")
        assert.is_true(ok, out)
      end
    end)
  end)
end
