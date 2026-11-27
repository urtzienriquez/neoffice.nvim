-- Bugs found by opening, editing and saving many documents (and edits made at random), one test
-- each, on documents made here: they must stay fixed.
local H = dofile("tests/helpers.lua")
local neo = require("neoffice")
local markup = require("neoffice.markup")
local zip = require("neoffice.zip")
local api = vim.api

local DOT = H.FIXTURES .. "dot.png"

local function edited(name, edits) -- a copy of a fixture, its parts edited: { [part] = { { from, to } } }
  return H.copy(name, function(entry, text)
    for _, e in ipairs(edits[entry] or {}) do
      local new, n = text:gsub(e[1], function() return e[2] end, 1)
      assert(n == 1, "fixture edit didn't apply: " .. e[1])
      text = new
    end
    return text
  end)
end
local function row_of(pattern) -- the first buffer line (1-based) that matches
  for r, l in ipairs(H.lines()) do
    if l:find(pattern) then
      return r, l
    end
  end
end
local function tidy(x) -- (XML as written again: "<a />" as "<a/>")
  return (x:gsub("%s+/>", "/>"))
end
local function saves_unchanged(path) -- a save without changes leaves the file as it is
  local before = H.bytes(path)
  vim.bo.modified = true
  H.save()
  assert.equals(before, H.bytes(path))
end
local function notes(fn) -- the messages said while fn runs
  local said, notify = {}, vim.notify
  vim.notify = function(msg) said[#said + 1] = msg end
  local ok, err = pcall(fn)
  vim.notify = notify
  assert(ok, err)
  return table.concat(said, "\n")
end

describe("reading", function()
  it("joins text typed after a paragraph's floating images, on their line, with nothing between", function()
    local lines = { "Text ![a](x.png){width = 15cm}", "![b](y.png){width = 5cm, float = left, place = page, x = 1cm, y = 2cm} X", "", "Next." }
    local paras = neo.paragraphs(lines)
    assert.equals("Text ![a](x.png){width = 15cm}![b](y.png){width = 5cm, float = left, place = page, x = 1cm, y = 2cm} X", paras[1].text)
  end)

  it("reads an alignment's mark alone, ending the text, as an empty paragraph's", function()
    assert.equals("center", markup.block("<|>").align)
    local block, chars = markup.parse("<|>")
    assert.equals("center", block.align)
    assert.same({}, chars)
  end)

  it("writes text that looks like attributes, after a mark, so that it doesn't read as a format's line", function()
    for _, align in ipairs({ "justify", "center", "right" }) do
      local spans = { { text = "{first = 0.5cm}" } }
      local line = markup.render(spans, { align = align })
      assert.is_nil(markup.para_line(line), line)
      assert.is_true(markup.same(line, spans), line)
      assert.equals(align, markup.parse(line).align)
    end
  end)

  it("takes neoffice's row heights for its own, not LibreOffice's padding of 0.097cm", function()
    local tables = require("neoffice.tables")
    assert.equals("normal", tables.height_of(0.1))
    assert.equals("normal", tables.height_of(57 * 2.54 / 1440)) -- (0.1cm, as written in twips)
    assert.is_nil(tables.height_of(55 * 2.54 / 1440))
    assert.is_nil(tables.height_of(0.097))
  end)
end)

for _, ext in ipairs({ "docx", "odt" }) do
  describe(ext, function()
    it("opens a table cell with nothing but its alignment, and saves it unchanged", function()
      local path = H.copy("rich." .. ext)
      H.open(path)
      local r = row_of("^| cell one")
      api.nvim_buf_set_lines(0, r - 1, r, false, { "| cell one | <|> |" })
      H.reopen()
      assert.truthy(H.doc(), "opened")
      local _, line = row_of("^| cell one")
      assert.truthy(line:match("| <|> *|$"), line)
      saves_unchanged(path)
    end)

    it("opens floating images at a list item's start and end", function()
      local path = H.copy("empty." .. ext)
      H.open(path)
      local float = function(y) return "![](" .. DOT .. "){width = 1cm, float = left, place = page, x = 1cm, y = " .. y .. "}" end
      api.nvim_buf_set_lines(0, 0, -1, false, { "- first item", "- " .. float("2cm") .. "second item", "- third item " .. float("5cm"), "- fourth item", "", "After." })
      H.reopen()
      assert.truthy(H.doc(), "opened")
      local text = table.concat(H.paras(), "\n")
      for _, item in ipairs({ "first item", "second item", "third item", "fourth item", "After." }) do
        assert.truthy(text:find(item, 1, true), text)
      end
      saves_unchanged(path)
    end)

    it("keeps a table inside a table's cell as it is, and says it isn't shown", function()
      local inner = ext == "docx"
          and { "word/document.xml", "cell two</w:t></w:r></w:p></w:tc>", 'cell two</w:t></w:r></w:p><w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="1000"/></w:tblGrid>'
            .. "<w:tr><w:tc><w:p><w:r><w:t>inner one</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>inner two</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"
            .. "<w:p><w:r><w:t>after inner</w:t></w:r></w:p></w:tc>" }
        or { "content.xml", "cell two</text:p></table:table%-cell>", 'cell two</text:p><table:table table:name="Inner"><table:table-column table:number-columns-repeated="2"/>'
          .. '<table:table-row><table:table-cell office:value-type="string"><text:p>inner one</text:p></table:table-cell>'
          .. '<table:table-cell office:value-type="string"><text:p>inner two</text:p></table:table-cell></table:table-row></table:table>'
          .. '<text:p text:style-name="P2">after inner</text:p></table:table-cell>' }
      local path = edited("rich." .. ext, { [inner[1]] = { { inner[2], inner[3] } } })
      local said = notes(function() H.open(path) end)
      assert.truthy(said:find("a table inside a table's cell", 1, true), said)
      local text = table.concat(H.lines(), "\n")
      assert.is_nil(text:find("inner one", 1, true), text)
      local _, line = row_of("^| cell one")
      assert.truthy(line:match("cell two <br> after inner"), line) -- (the outer table as it is: two cells)
      saves_unchanged(path)
      local k = vim.fn.match(H.paras(), "after inner") + 1
      H.replace(k, "after inner", "after inner X")
      H.reopen()
      local main = zip.read(path, H.MAIN[ext])
      assert.truthy(main:find("inner one", 1, true) and main:find("inner two", 1, true) and main:find("after inner X", 1, true))
      _, line = row_of("^| cell one")
      assert.truthy(line:match("^| cell one *| cell two <br> after inner X *|$"), line)
    end)

    it("doesn't give text typed after a link, at its paragraph's end, the link's style", function()
      H.open(H.copy("rich." .. ext))
      local k = vim.fn.match(H.paras(), "(https://example.org)") + 1
      H.replace(k, ". Tom & Jerry <3, 2\\*3.", "") -- (the link last)
      H.reopen()
      assert.truthy(vim.endswith(H.paras()[k], "[link](https://example.org)"), H.paras()[k])
      H.replace(k, "(https://example.org)", "(https://example.org) X")
      H.reopen()
      assert.truthy(vim.endswith(H.paras()[k], "[link](https://example.org) X"), H.paras()[k])
    end)

    it("keeps a space after an image at a cell's start", function()
      H.open(H.copy("rich." .. ext))
      local r = row_of("^| cell one")
      api.nvim_buf_set_lines(0, r - 1, r, false, { "| ![](" .. DOT .. "){width = 1cm} kept | cell two |" })
      H.reopen()
      local _, line = row_of("kept")
      assert.truthy(line:find("{width = 1cm} kept", 1, true), line)
    end)

    it("gives a format's line to the paragraphs after it when it is written right after opening", function()
      local many = {} -- (long enough that the paragraphs' fingerprints aren't all taken at once)
      for n = 1, 120 do
        many[#many + 1] = ext == "docx" and ("<w:p><w:r><w:t>Paragraph %d.</w:t></w:r></w:p>"):format(n) or ("<text:p>Paragraph %d.</text:p>"):format(n)
      end
      local path = edited("rich." .. ext, ext == "docx" and { ["word/document.xml"] = { { "<w:body>", "<w:body>" .. table.concat(many) } } }
        or { ["content.xml"] = { { "<office:text>", "<office:text>" .. table.concat(many) } } })
      vim.cmd("silent! %bwipeout!")
      vim.cmd("edit " .. vim.fn.fnameescape(path)) -- (no wait: before the paragraphs' fingerprints are all taken)
      assert.truthy(H.doc().pending)
      local r = H.row(3)
      api.nvim_buf_set_lines(0, r - 1, r - 1, false, { "{indent-left = 1cm}+" })
      H.reopen()
      local formatted = {}
      for _, b in ipairs(neo.blocks(0)) do
        if b.format and b.format.attrs["indent-left"] == "1cm" then
          formatted[#formatted + 1] = b.text
        end
      end
      assert.truthy(#formatted >= 3, vim.inspect(formatted))
      assert.truthy(vim.tbl_contains(formatted, "Last paragraph."), vim.inspect(formatted))
    end)

    it("opens a paragraph whose text looks like attributes after its mark, right above a table", function()
      local path = H.copy("rich." .. ext)
      H.open(path)
      local r = row_of("^| A")
      api.nvim_buf_set_lines(0, r - 1, r - 1, false, { "<=> {first = 0.5cm}" })
      H.reopen()
      assert.truthy(H.doc(), "opened")
      H.reopen()
      assert.truthy(H.doc(), "opened again")
      saves_unchanged(path)
    end)
  end)
end

describe("docx", function()
  it("shows two paragraph styles of one name each with its own, and changes only the one said", function()
    local path = edited("rich.docx", {
      ["word/styles.xml"] = { { "</w:styles>", '<w:style w:type="paragraph" w:customStyle="1" w:styleId="BodyText0"><w:name w:val="Body Text"/>'
        .. '<w:basedOn w:val="Normal"/><w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr></w:style></w:styles>' } },
      ["word/document.xml"] = { { '<w:pStyle w:val="BodyText" /></w:pPr><w:r><w:t xml:space="preserve">Last paragraph',
        '<w:pStyle w:val="BodyText0" /></w:pPr><w:r><w:t xml:space="preserve">Last paragraph' } },
    })
    neo.config.styles = true
    local ok, err = pcall(function()
      H.open(path)
      assert.truthy(row_of("^%s+Body Text:"), table.concat(H.lines(), "\n"))
      local r, line = row_of("^%s+Body Text %(BodyText0%):")
      assert.truthy(r, table.concat(H.lines(), "\n"))
      saves_unchanged(path)
      local before = zip.read(path, "word/styles.xml")
      api.nvim_buf_set_lines(0, r - 1, r, false, { (line:gsub("}$", ", indent-left = 2cm}")) })
      H.reopen()
      local styles = zip.read(path, "word/styles.xml")
      local function style(s, id)
        return tidy(s:match('<w:style [^>]*w:styleId="' .. id .. '".-</w:style>'))
      end
      assert.equals(style(before, "BodyText"), style(styles, "BodyText"))
      assert.truthy(style(styles, "BodyText0"):find('w:left="1134"', 1, true), style(styles, "BodyText0"))
    end)
    neo.config.styles = false
    assert(ok, err)
  end)

  it("saves unchanged an equation of several rows (its LaTeX on one line)", function()
    if not H.has_pandoc then
      return
    end
    local path = H.copy("empty.docx")
    H.open(path)
    api.nvim_buf_set_lines(0, 0, -1, false, { "$$\\begin{cases} a & b \\\\ c & d \\end{cases}$$" })
    H.reopen()
    for _, object in pairs(H.doc().object) do
      if object.math then
        assert.is_nil(object.math.tex:find("\n", 1, true), object.math.tex)
      end
    end
    saves_unchanged(path)
  end)

  it("keeps a centred title's own left alignment when its text changes", function()
    local path = edited("rich.docx", { ["word/document.xml"] = { { "<w:body>",
      '<w:body><w:p><w:pPr><w:pStyle w:val="Title"/><w:jc w:val="left"/></w:pPr><w:r><w:t>A left title</w:t></w:r></w:p>' } } })
    H.open(path)
    assert.equals("% <<- A left title", H.paras()[1])
    H.replace(1, "left title", "left title X")
    H.reopen()
    assert.equals("% <<- A left title X", H.paras()[1])
    assert.truthy(zip.read(path, "word/document.xml"):find('<w:pStyle w:val="Title"/><w:jc w:val="left"/>', 1, true))
  end)

  it("doesn't typeset again a table with LibreOffice's cell padding when its text changes", function()
    local mar = '<w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
      .. '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
      .. '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders>'
      .. '<w:tblCellMar><w:top w:w="55" w:type="dxa"/><w:left w:w="55" w:type="dxa"/><w:bottom w:w="55" w:type="dxa"/><w:right w:w="55" w:type="dxa"/></w:tblCellMar>'
    local path = edited("rich.docx", { ["word/document.xml"] = { { "<w:tblLook ", mar .. "<w:tblLook " } } })
    H.open(path)
    local function table_props()
      local main = zip.read(path, "word/document.xml")
      return tidy(main:match("<w:tblPr>.-</w:tblPr>") .. main:match("<w:tblGrid>.-</w:tblGrid>"))
    end
    local before = table_props()
    for _, cell in ipairs({ "cell one", "cell two" }) do
      local k = vim.fn.index(H.paras(), cell) + 1
      H.replace(k, cell, cell .. " X")
    end
    H.reopen()
    assert.truthy(row_of("cell one X"))
    assert.equals(before, table_props())
  end)

  it("writes back a deleted line break as one, when its paragraph changes", function()
    local p = '<w:p><w:r><w:t xml:space="preserve">Kept line.</w:t></w:r><w:del w:id="91" w:author="Rev" w:date="2024-05-01T10:00:00Z"><w:r><w:br/></w:r>'
      .. '<w:r><w:delText xml:space="preserve">Gone line.</w:delText></w:r></w:del><w:r><w:br/></w:r><w:r><w:t xml:space="preserve">Last line</w:t></w:r></w:p>'
    local path = edited("rich.docx", { ["word/document.xml"] = { { "<w:body>", "<w:body>" .. p } } })
    H.open(path)
    H.replace(1, "Last line", "Last line X")
    H.reopen()
    local main = zip.read(path, "word/document.xml")
    local del = main:match("<w:del [^>]*>.-</w:del>")
    assert.truthy(del and del:find("<w:br/>", 1, true) and del:find("Gone line.", 1, true), del)
    assert.is_nil(del:find('<w:delText xml:space="preserve"> </w:delText>', 1, true))
  end)

  it("shows a deletion of a line break alone, and keeps it when its paragraph changes", function()
    local p = '<w:p><w:r><w:t xml:space="preserve">First line.</w:t></w:r><w:del w:id="92" w:author="Rev" w:date="2024-05-01T10:00:00Z"><w:r><w:br/></w:r></w:del>'
      .. '<w:r><w:t xml:space="preserve">Second line</w:t></w:r></w:p>'
    local path = edited("rich.docx", { ["word/document.xml"] = { { "<w:body>", "<w:body>" .. p } } })
    H.open(path)
    local shown = H.paras()[1]
    assert.equals("First line.\\\nSecond line", shown) -- (the break: deleted, shown so)
    H.replace(1, "Second line", "Second line X")
    H.reopen()
    assert.equals(shown .. " X", H.paras()[1])
    local del = zip.read(path, "word/document.xml"):match("<w:del [^>]*>.-</w:del>")
    assert.truthy(del and del:find("<w:br/>", 1, true), del)
  end)
end)

describe("marks that close nothing", function()
  it("are found where they would be written as characters, not where characters are meant", function()
    local function marks(text, anywhere)
      return vim.tbl_map(function(o) return o.mark .. "@" .. o.col end, markup.orphan_marks(text, anywhere))
    end
    assert.same({ "**@0" }, marks("**this is nothing"))
    assert.same({ "*@2" }, marks("a *word with no end"))
    assert.same({ "~~@0" }, marks("~~gone"))
    assert.same({ "^{@1" }, marks("x^{2 and the rest"))
    assert.same({ "_{@1" }, marks("H_{2 and the rest"))
    for _, text in ipairs({ "**bold** and *it*", "\\*\\*escaped", "2 * 3 and 2*3", "p < 0.05*", "C++ code", "a == b", "x^{2} closed" }) do
      assert.same({}, marks(text), text)
    end
    assert.same({ "++@1" }, marks("C++ code", true)) -- (anywhere: what the save looks at, see below)
  end)

  for _, ext in ipairs({ "docx", "odt" }) do
    describe(ext, function()
      local path
      before_each(function()
        path = H.copy("empty." .. ext)
        H.open(path)
        api.nvim_buf_set_lines(0, 0, -1, false, { "**this is s**omething that is written", "", "Plain C++ text, a == b." })
        H.reopen()
      end)
      local function line_of(text)
        return row_of(vim.pesc(text))
      end
      local function replace(from, to) -- (as typing: in the line where `from` is)
        local r, line = line_of(from)
        local s = line:find(from, 1, true)
        api.nvim_buf_set_text(0, r - 1, s - 1, r - 1, s - 1 + #from, { to })
        require("neoffice.changes").after_edit(api.nvim_get_current_buf())
      end

      it("doesn't save a word deleted with the mark in it, which would turn the bold into characters, and says where", function()
        local before = H.bytes(path)
        replace("s**omething", "nothing")
        local said = notes(function() H.save() end)
        assert.equals(before, H.bytes(path))
        assert.truthy(said:find("closes nothing any more", 1, true), said)
        assert.truthy(said:find(("line %d, column 1"):format((line_of("this is nothing"))), 1, true), said)
        -- (written as characters on purpose: saved so)
        replace("**this is nothing", "\\*\\*this is nothing")
        H.reopen()
        assert.truthy(line_of("\\*\\*this is nothing that is written"), table.concat(H.lines(), "\n"))
      end)

      it("saves edits that keep the marks, and the characters that were there already", function()
        replace("s**omething", "n**othing")
        replace("a == b", "a == c")
        H.reopen()
        assert.truthy(line_of("**this is n**othing that is written"), table.concat(H.lines(), "\n"))
        assert.truthy(line_of("Plain C++ text, a == c."))
      end)

      it("doesn't save a superscript left open, which would raise all that follows", function()
        api.nvim_buf_set_lines(0, -1, -1, false, { "", "Area in cm^{2} and more." })
        H.reopen()
        local before = H.bytes(path)
        replace("cm^{2}", "cm^{2")
        local said = notes(function() H.save() end)
        assert.equals(before, H.bytes(path))
        assert.truthy(said:find("isn't closed any more", 1, true), said)
      end)

      it("are shown as warnings where they are", function()
        replace("s**omething", "nothing")
        local buf = api.nvim_get_current_buf()
        local found = vim.wait(3000, function()
          return vim.iter(vim.diagnostic.get(buf)):any(function(d) return d.message:find("closes nothing", 1, true) ~= nil end)
        end, 20)
        assert.is_true(found)
        local d = vim.iter(vim.diagnostic.get(buf)):find(function(x) return x.message:find("closes nothing", 1, true) ~= nil end)
        assert.equals(vim.diagnostic.severity.WARN, d.severity)
        assert.equals((line_of("this is nothing")) - 1, d.lnum)
        assert.equals(0, d.col)
      end)
    end)
  end
end)
