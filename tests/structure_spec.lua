-- Titles (% ), bullet and numbered lists (- , #. ) and numbered headings.
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
    local texts = vim.tbl_map(function(p) return p.text end, neo.paragraphs({ "% T", "# H", "text", "- a", "- b", "  - c", "#. d", "more", "1. e", "2.1. f" }))
    assert.same({ "% T", "# H text", "- a", "- b", "  - c", "#. d more 1. e 2.1. f" }, texts) -- (numbers in text: text)
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

    it("writes bullet and numbered lists, nested (the document counts)", function()
      new(ext, { "Intro.", "", "- apples", "- pears", "  - green ones", "- plums", "", "#. first", "#. second", "  #. inner", "#. third", "", "The end." })
      assert.same({
        "Intro.", "", "- apples", "- pears", "  - green ones", "- plums", "", "#. first", "#. second", "  #. inner", "#. third", "", "The end.",
      }, H.lines())
    end)

    it("shows text starting with a number as it is, and #. typed as text escaped", function()
      new(ext, { "1. Introduction", "", "2.1. Methods, as text", "", "49\\. Smith, J. (2020)", "", "\\#. not an item" })
      assert.same({ "1. Introduction", "", "2.1. Methods, as text", "", "49. Smith, J. (2020)", "", "\\#. not an item" }, H.lines())
      assert.same({ "1. Introduction", "2.1. Methods, as text", "49. Smith, J. (2020)", "#. not an item" }, vim.tbl_map(function(p)
        local _, chars = require("neoffice.markup").parse(p)
        return table.concat(vim.tbl_map(function(c) return c.ch end, chars))
      end, H.paras()))
    end)

    it("numbers new lists 1., a., i. by level, each list its own", function()
      new(ext, { "#. first", "#. second", "  #. second sub one", "  #. second sub two", "    #. deeper", "", "Between.", "", "#. another list" })
      assert.same({ "#. first", "#. second", "  #. second sub one", "  #. second sub two", "    #. deeper", "", "Between.", "", "#. another list" }, H.lines())
      if H.has_libreoffice then
        local ok, text = H.libreoffice_reads(H.doc().path, "second sub one")
        assert.is_true(ok, text)
        for _, want in ipairs({ "2%.%s+second", "a%.%s+second sub one", "b%.%s+second sub two", "i%.%s+deeper", "1%.%s+another list" }) do
          assert.truthy(text:find(want), want .. " in " .. text)
        end
      end
    end)

    it("adds an item to a list, and takes one out of it", function()
      new(ext, { "- one", "- two", "- three" })
      api.nvim_buf_set_lines(0, 1, 1, false, { "- one and a half" })
      H.set(4, "\nthree, plain now") -- (as in Markdown, a blank line ends the list)
      H.reopen()
      assert.same({ "- one", "- one and a half", "- two", "", "three, plain now" }, H.lines())
    end)

    it("takes an item out of the middle of a list, splitting it", function()
      new(ext, { "#. one", "#. two", "#. three" })
      H.set(2, "\ntwo, plain now\n")
      H.reopen()
      -- (Word goes on with the same list after it, LibreOffice starts again: shown as they will be)
      assert.same({ "#. one", "", "two, plain now", "", "#. three" }, H.lines())
    end)

    it("changes the kind and level of items", function()
      new(ext, { "- one", "- two", "- three" })
      H.set(2, "  - two, deeper")
      H.set(3, "#. three, numbered")
      H.reopen()
      assert.same({ "- one", "  - two, deeper", "", "#. three, numbered" }, H.lines())
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
      local numbers = {}
      for _, m in ipairs(api.nvim_buf_get_extmarks(0, ns, 0, -1, { details = true })) do
        if m[4].virt_text then
          numbers[#numbers + 1] = m[4].virt_text[1][1]
        end
      end
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
      assert.equals(0, #vim.tbl_filter(function(m) return m[4].virt_text end, api.nvim_buf_get_extmarks(0, ns, 0, -1, { details = true }))) -- (no numbers)
    end)

    if H.has_libreoffice then
      it("writes lists LibreOffice shows as lists", function()
        local path = new(ext, { "- apples", "  - green", "#. first", "#. second" })
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

describe("sub-items numbered their own way under different items of one list", function()
  for _, ext in ipairs({ "docx", "odt" }) do
    it("are counted on as one list: 1., a., b., 2., a., b., 3. (" .. ext .. ")", function()
      local lines = { "#. First, clean the data", "  #. something here", "  #. something else.", "#. Then, fit the model",
        "  #. hey there", "  #. what is this?", "#. Finally, write it up" }
      new(ext, lines)
      assert.same(lines, H.lines())
      if H.has_libreoffice then
        local ok, text = H.libreoffice_reads(H.doc().path, "hey there")
        assert.is_true(ok, text)
        for _, want in ipairs({ "1%.%s+First", "a%.%s+something here", "b%.%s+something else", "2%.%s+Then", "a%.%s+hey there", "b%.%s+what is this", "3%.%s+Finally" }) do
          assert.truthy(text:find(want), want .. " in " .. text)
        end
      end
    end)

    it("join the list above them when no blank line is between, and start another after one (" .. ext .. ")", function()
      new(ext, { "#. one", "", "#. two", "#. three" })
      assert.same({ "#. one", "", "#. two", "#. three" }, H.lines()) -- (two lists: the blank line says so)
      api.nvim_buf_set_lines(0, 1, 2, false, {}) -- (the blank line deleted: one list)
      H.reopen()
      assert.same({ "#. one", "#. two", "#. three" }, H.lines())
      if H.has_libreoffice then
        local ok, text = H.libreoffice_reads(H.doc().path, "three")
        assert.is_true(ok, text)
        assert.truthy(text:find("3%. three"), text)
      end
    end)
  end
end)

describe("a list's numbering as the document has it (roman numerals and \")\", here)", function()
  for _, ext in ipairs({ "docx", "odt" }) do
    it("is kept when its items are edited, and one is added (" .. ext .. ")", function()
      local path = new(ext, { "#. first", "#. second" })
      if ext == "docx" then
        local numbering = zip.read(path, "word/numbering.xml")
        numbering = numbering:gsub('(<w:lvl w:ilvl="0">.-)<w:numFmt w:val="decimal"/><w:lvlText w:val="%%1%."/>', '%1<w:numFmt w:val="upperRoman"/><w:lvlText w:val="%%1)"/>', 1)
        assert.truthy(numbering:find("upperRoman", 1, true))
        zip.write(path, { ["word/numbering.xml"] = numbering })
      else
        local content = zip.read(path, "content.xml")
        content = content:gsub('(<text:list%-level%-style%-number text:level="1") style:num%-suffix="%." style:num%-format="1"', '%1 style:num-suffix=")" style:num-format="I"', 1)
        assert.truthy(content:find('style:num-format="I"', 1, true))
        zip.write(path, { ["content.xml"] = content })
      end
      H.open(path)
      assert.same({ "#. first", "#. second" }, H.lines())
      H.replace(1, "first", "first, edited")
      api.nvim_buf_set_lines(0, 2, 2, false, { "#. third, new" })
      H.reopen()
      assert.same({ "#. first, edited", "#. second", "#. third, new" }, H.lines())
      if ext == "docx" then
        assert.truthy(zip.read(path, "word/numbering.xml"):find("upperRoman", 1, true))
      end
      if H.has_libreoffice then
        local ok, text = H.libreoffice_reads(path, "third, new")
        assert.is_true(ok, text)
        for _, want in ipairs({ "I%)%s+first, edited", "II%)%s+second", "III%)%s+third, new" }) do
          assert.truthy(text:find(want), want .. " in " .. text)
        end
      end
    end)
  end
end)

describe("lists open as they were written, without empty lines between their items", function()
  local SHAPES = {
    { "#. one", "  #. one one", "    #. deeper", "    #. deeper again", "  #. one two", "#. two" },
    { "#. First", "  #. a", "  #. b", "#. Then", "  #. c", "  #. d", "  #. e", "#. Finally" },
    { "#. numbered", "  - a bullet under it", "  - another", "#. numbered again" },
    { "- bullets", "- more", "", "#. then numbers", "#. more" },
    { "#. one list", "", "#. another list", "#. its second" },
    { "#. an item\\", "with a line break", "#. next" },
  }
  -- (as LibreOffice saves it again: its own way of writing lists)
  local function resaved(path, ext)
    local dir = vim.fn.tempname()
    vim.fn.mkdir(dir, "p")
    vim.system({ "soffice", "-env:UserInstallation=file://" .. dir .. "/profile", "--headless", "--convert-to", ext, "--outdir", dir, path }):wait()
    return dir .. "/" .. vim.fn.fnamemodify(path, ":t")
  end
  for _, ext in ipairs({ "docx", "odt" }) do
    for k, lines in ipairs(SHAPES) do
      it(("shape %d (%s), and after LibreOffice saves it again"):format(k, ext), function()
        local path = new(ext, lines)
        assert.same(lines, H.lines())
        if H.has_libreoffice then
          local again = resaved(path, ext)
          assert.truthy(vim.uv.fs_stat(again), again)
          H.open(again)
          assert.same(lines, H.lines())
        end
      end)
    end
  end
end)
