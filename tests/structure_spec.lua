-- Titles (% ), bullet and numbered lists (- , 1. ) and numbered headings.
local H = dofile("tests/helpers.lua")
local neo = require("neoffice")
local zip = require("neoffice.zip")
local api = vim.api

local function new(ext, lines) -- a new document with these lines, saved and opened again
  local path = vim.fn.tempname() .. "." .. ext
  H.open(path)
  api.nvim_buf_set_lines(0, 0, -1, false, lines)
  H.reopen()
  return path
end

describe("paragraph layout", function()
  it("starts a paragraph at each heading, title or list item, even without a blank line", function()
    local texts = vim.tbl_map(function(p) return p.text end, neo.paragraphs({ "% T", "# H", "text", "- a", "- b", "  - c", "1. d", "more" }))
    assert.same({ "% T", "# H text", "- a", "- b", "  - c", "1. d more" }, texts)
  end)
end)

for _, ext in ipairs({ "docx", "odt" }) do
  describe(ext, function()
    it("writes a title, and turns it back into a plain paragraph", function()
      local path = new(ext, { "% this is the title", "", "# heading", "", "what?" })
      assert.same({ "% this is the title", "# heading", "what?" }, H.paras())
      if ext == "docx" then
        assert.truthy(zip.read(path, "word/document.xml"):find('<w:pStyle w:val="Title"/>', 1, true))
        assert.truthy(zip.read(path, "word/styles.xml"):find('w:styleId="Title"', 1, true))
      else
        assert.truthy(zip.read(path, "content.xml"):find('text:style-name="Title"', 1, true))
      end
      H.set(1, "not a title any more")
      H.reopen()
      assert.equals("not a title any more", H.paras()[1])
    end)

    it("writes bullet and numbered lists, nested, and shows their numbers", function()
      new(ext, { "Intro.", "", "- apples", "- pears", "  - green ones", "- plums", "", "1. first", "1. second", "  1. inner", "1. third", "", "The end." })
      assert.same({
        "Intro.", "", "- apples", "- pears", "  - green ones", "- plums", "", "1. first", "2. second", "  1. inner", "3. third", "", "The end.",
      }, H.lines())
    end)

    it("adds an item to a list, and takes one out of it", function()
      new(ext, { "- one", "- two", "- three" })
      api.nvim_buf_set_lines(0, 1, 1, false, { "- one and a half" })
      H.set(4, "\nthree, plain now") -- (as in Markdown, a blank line ends the list)
      H.reopen()
      assert.same({ "- one", "- one and a half", "- two", "", "three, plain now" }, H.lines())
    end)

    it("takes an item out of the middle of a list, splitting it", function()
      new(ext, { "1. one", "1. two", "1. three" })
      H.set(2, "\ntwo, plain now\n")
      H.reopen()
      -- (Word goes on with the same list after it, LibreOffice starts again: shown as they will be)
      assert.same({ "1. one", "", "two, plain now", "", ext == "docx" and "2. three" or "1. three" }, H.lines())
    end)

    it("changes the kind and level of items", function()
      new(ext, { "- one", "- two", "- three" })
      H.set(2, "  - two, deeper")
      H.set(3, "1. three, numbered")
      H.reopen()
      assert.same({ "- one", "  - two, deeper", "", "1. three, numbered" }, H.lines())
    end)

    it("keeps lists it didn't change exactly as they were", function()
      local path = H.copy("rich." .. ext)
      H.open(path)
      local before = H.paragraph_xml(path)
      H.replace(12, "Last", "Final")
      H.save()
      local after = H.paragraph_xml(path)
      assert.same({ before[6], before[7] }, { after[6], after[7] })
      local main = zip.read(path, H.MAIN[ext])
      assert.truthy(main:find(ext == "docx" and "<w:numPr>" or "<text:list ", 1, true))
    end)

    it("numbers headings on demand, and stops", function()
      local path = new(ext, { "# One", "", "## One point one", "", "## One point two", "", "# Two" })
      vim.cmd("Neoffice numbers on")
      assert.is_true(vim.bo.modified)
      local ns = api.nvim_get_namespaces().neoffice_decorations
      local numbers = vim.tbl_map(function(m) return m[4].virt_text[1][1] end, api.nvim_buf_get_extmarks(0, ns, 0, -1, { details = true }))
      assert.same({ "1 ", "1.1 ", "1.2 ", "2 " }, numbers)
      H.reopen()
      assert.is_true(H.doc().ctx.numbered)
      if H.has_libreoffice then
        local ok, text = H.libreoffice_reads(path, "1.2 One point two")
        assert.is_true(ok, text)
      end
      vim.cmd("Neoffice numbers off")
      H.reopen()
      assert.is_false(H.doc().ctx.numbered)
      assert.equals(0, #api.nvim_buf_get_extmarks(0, ns, 0, -1, {}))
    end)

    if H.has_libreoffice then
      it("writes lists LibreOffice shows as lists", function()
        local path = new(ext, { "- apples", "  - green", "1. first", "1. second" })
        local ok, text = H.libreoffice_reads(path, "1. first")
        assert.is_true(ok, text)
        assert.truthy(text:find("• apples", 1, true))
        assert.truthy(text:find("2. second", 1, true))
      end)
    end
  end)
end

describe("heading numbers made elsewhere", function()
  it("(docx: on the headings themselves, as LibreOffice writes them) are seen and removed", function()
    local path = H.copy("rich.docx", function(name, s)
      return name == "word/document.xml" and (s:gsub('<w:pStyle w:val="Heading1" />', '<w:pStyle w:val="Heading1" /><w:numPr><w:ilvl w:val="0" /><w:numId w:val="1001" /></w:numPr>', 1)) or s
    end)
    H.open(path)
    assert.is_true(H.doc().ctx.numbered)
    assert.equals("# Introduction", H.paras()[1]) -- (a heading, not a list item)
    vim.cmd("Neoffice numbers off")
    H.reopen()
    assert.is_false(H.doc().ctx.numbered)
    assert.falsy(H.paragraph_xml(path)[1]:find("numPr", 1, true))
    assert.equals("# Introduction", H.paras()[1])
  end)

  it("(odt: by being in a list) are seen and removed", function()
    local path = H.copy("rich.odt", function(name, s)
      return name == "content.xml" and (s:gsub("(<text:h .-</text:h>)", '<text:list text:style-name="WWNum1001"><text:list-item>%1</text:list-item></text:list>', 1)) or s
    end)
    H.open(path)
    assert.is_true(H.doc().ctx.numbered)
    vim.cmd("Neoffice numbers off")
    H.reopen()
    assert.is_false(H.doc().ctx.numbered)
    assert.equals("# Introduction", H.paras()[1])
    assert.falsy(zip.read(path, "content.xml"):find("<text:list[^>]*><text:list%-item><text:h"))
  end)
end)
