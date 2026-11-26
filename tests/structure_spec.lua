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

    it("writes bullet and numbered lists, nested, and shows their numbering's patterns (the document counts)", function()
      new(ext, { "Intro.", "", "- apples", "- pears", "  - green ones", "- plums", "", "1. first", "1. second", "  1. inner", "1. third", "", "The end." })
      assert.same({
        "Intro.", "", "- apples", "- pears", "  - green ones", "- plums", "", "1. first", "1. second", "  1. inner", "1. third", "", "The end.",
      }, H.lines())
    end)

    it("numbers sub-items as their pattern says (1.1., 1.a.), only in their list", function()
      new(ext, { "1. first", "1. second", "  1.1. second sub one", "  1.1. second sub two", "", "Between.", "", "1. another list" })
      assert.same({ "1. first", "1. second", "  1.1. second sub one", "  1.1. second sub two", "", "Between.", "", "1. another list" }, H.lines())
      local xml = zip.read(H.doc().path, H.MAIN[ext])
      if ext == "docx" then
        assert.truthy(zip.read(H.doc().path, "word/numbering.xml"):find('<w:lvlText w:val="%1.%2."/>', 1, true))
      else
        assert.truthy(zip.read(H.doc().path, "content.xml"):find('text:display%-levels="2"'), xml:sub(1, 400))
      end
      H.replace(3, "1.1.", "1.a.")
      H.replace(4, "1.1.", "1.a.")
      H.reopen()
      assert.same({ "1. first", "1. second", "  1.a. second sub one", "  1.a. second sub two", "", "Between.", "", "1. another list" }, H.lines())
      H.replace(6, "1.", "a)") -- (the other list: its own numbering, the first's as it was)
      H.reopen()
      assert.same({ "1. first", "1. second", "  1.a. second sub one", "  1.a. second sub two", "", "Between.", "", "a) another list" }, H.lines())
      if H.has_libreoffice then
        local ok, text = H.libreoffice_reads(H.doc().path, "second sub one")
        assert.is_true(ok, text)
        assert.truthy(text:find("2%.b%.") or text:find("2%.b"), text) -- (the second's second: 2.b.)
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
      new(ext, { "1. one", "1. two", "1. three" })
      H.set(2, "\ntwo, plain now\n")
      H.reopen()
      -- (Word goes on with the same list after it, LibreOffice starts again: shown as they will be)
      assert.same({ "1. one", "", "two, plain now", "", "1. three" }, H.lines())
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

describe("sub-items numbered their own way under different items of one list", function()
  for _, ext in ipairs({ "docx", "odt" }) do
    it("are counted on as one list: 1., 1.1., 1.2., 2., 2.a., 2.b., 3. (" .. ext .. ")", function()
      local lines = { "1. First, clean the data", "  1.1. something here", "  1.1. something else.", "1. Then, fit the model",
        "  1.a. hey there", "  1.a. what is this?", "1. Finally, write it up" }
      new(ext, lines)
      assert.same(lines, H.lines())
      if H.has_libreoffice then
        local ok, text = H.libreoffice_reads(H.doc().path, "hey there")
        assert.is_true(ok, text)
        local labels = {}
        for label in text:gmatch("(%d[%w.]*%.) ") do
          labels[#labels + 1] = label
        end
        assert.same({ "1.", "1.1.", "1.2.", "2.", "2.a.", "2.b.", "3." }, labels, text)
      end
    end)

    it("join the list above them when no blank line is between, and start another after one (" .. ext .. ")", function()
      new(ext, { "1. one", "", "1. two", "1. three" })
      assert.same({ "1. one", "", "1. two", "1. three" }, H.lines()) -- (two lists: the blank line says so)
      api.nvim_buf_set_lines(0, 1, 2, false, {}) -- (the blank line deleted: one list)
      H.reopen()
      assert.same({ "1. one", "1. two", "1. three" }, H.lines())
      if H.has_libreoffice then
        local ok, text = H.libreoffice_reads(H.doc().path, "three")
        assert.is_true(ok, text)
        assert.truthy(text:find("3%. three"), text)
      end
    end)
  end
end)

describe("docx lists that share their numbering's definition (as Word and pandoc write them)", function()
  it("gives a list its own when its pattern changes: the other keeps its numbering", function()
    local path = new("docx", { "1. first list", "", "Between.", "", "1. second list" })
    -- (the second list made to share the first's definition)
    local numbering, document = zip.read(path, "word/numbering.xml"), zip.read(path, "word/document.xml")
    local ids = {}
    for id in document:gmatch('<w:numId w:val="(%d+)"/>') do
      ids[#ids + 1] = id
    end
    assert.equals(2, #ids)
    assert.are_not.equal(ids[1], ids[2])
    local first = numbering:match('<w:num w:numId="' .. ids[1] .. '"><w:abstractNumId w:val="(%d+)"/>')
    numbering = numbering:gsub('(<w:num w:numId="' .. ids[2] .. '"><w:abstractNumId w:val=")(%d+)("/>)', "%1" .. first
      .. '%3<w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/></w:lvlOverride>') -- (a list of its own: numbered again from 1)
    zip.write(path, { ["word/numbering.xml"] = numbering })
    H.open(path)
    assert.same({ "1. first list", "", "Between.", "", "1. second list" }, H.lines())
    H.replace(3, "1.", "a)")
    H.reopen()
    assert.same({ "1. first list", "", "Between.", "", "a) second list" }, H.lines())
    if H.has_libreoffice then -- (the first as it was: 1.; the second its own way: a))
      local ok, text = H.libreoffice_reads(path, "second list")
      assert.is_true(ok, text)
      assert.truthy(text:find("1%. first list") and text:find("a%) second list"), text)
    end
  end)
end)

describe("lists open as they were written, without empty lines between their items", function()
  local SHAPES = {
    { "1. one", "  1.1. one one", "    1.1.1. deeper", "    1.1.1. deeper again", "  1.1. one two", "1. two" },
    { "1. First", "  1.1. a", "  1.1. b", "1. Then", "  1.a. c", "  1.a. d", "  1.a. e", "1. Finally" },
    { "1. numbered", "  - a bullet under it", "  - another", "1. numbered again" },
    { "- bullets", "- more", "", "1. then numbers", "1. more" },
    { "1. one list", "", "1. another list", "1. its second" },
    { "1. an item\\", "with a line break", "1. next" },
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
