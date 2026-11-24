-- Tables as pipe tables: editing cells, adding and removing rows, columns and tables.
local H = dofile("tests/helpers.lua")
local zip = require("neoffice.zip")
local api = vim.api

local function table_rows() -- the table's lines (without the line of dashes)
  return vim.tbl_filter(function(l) return l:match("^|") and not l:match("^|%-") end, H.lines())
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

    it("shows a table as a pipe table", function()
      local lines = H.lines()
      local r = vim.fn.index(lines, "| A | B |") + 1
      assert.same({ "| A | B |", "|---|---|", "| cell one | cell two |", "" }, vim.list_slice(lines, r, r + 3))
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

    it("makes a new table", function()
      local r = H.row(12)
      api.nvim_buf_set_lines(0, r - 1, r - 1, false, { "| N | M |", "|---|---|", "| 1 | 2 |", "" })
      H.reopen()
      local lines = table.concat(H.lines(), "\n")
      assert.truthy(lines:find("| N | M |\n|---|---|\n| 1 | 2 |\n\nLast paragraph.", 1, true))
      assert.equals(2, count(ext == "docx" and "<w:tbl>" or "<table:table "))
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
      H.input = "On a cell"
      require("neoffice.comments").add()
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
