-- Page settings: the document's at its top ("page: {size = A4, margins = 2.5cm}"), and a section's
-- where it starts ("\newpage {orientation = landscape}"), saying what changes.
local H = dofile("tests/helpers.lua")
local zip = require("neoffice.zip")
local neo = require("neoffice")
local pages = require("neoffice.pages")
local markup = require("neoffice.markup")
local api = vim.api

describe("page settings", function()
  it("read attributes (a size by name or its sides, margins as one, two or each), and say which aren't", function()
    assert.same({ size = "A4", orientation = "landscape", ["margin-top"] = "2cm", ["margin-bottom"] = "2cm", ["margin-left"] = "3cm", ["margin-right"] = "3cm" },
      pages.parse("size = a4, landscape, margins = 2cm 3cm"))
    assert.same({ size = "15cm x 20cm" }, pages.parse("size = 20cm x 15cm"))
    assert.same({ ["margin-left"] = "2.54cm" }, pages.parse("margin-left = 1in"))
    assert.is_nil(pages.parse("colour = red"))
    assert.is_nil(pages.parse("orientation = sideways"))
    assert.is_nil(pages.parse("size = huge"))
  end)

  it("write what changes, the margins as briefly as they are", function()
    local a4 = pages.view({ width = 21, height = 29.7, top = 2.5, bottom = 2.5, left = 3, right = 3 })
    assert.equals("size = A4, margins = 2.5cm 3cm", pages.render(a4, true))
    local wide = pages.view({ width = 29.7, height = 21, top = 3, bottom = 3, left = 2.5, right = 2.5 })
    assert.equals("orientation = landscape, margins = 3cm 2.5cm", pages.render(pages.delta(a4, wide)))
    assert.equals("", pages.render(pages.delta(a4, a4)))
    assert.same(a4, pages.apply(wide, pages.delta(wide, a4)))
  end)

  it("keep a page break with braces after it as text: it isn't a section's line", function()
    local spans = { { page_break = true }, { text = " {orientation = landscape}" } }
    local line = markup.render(spans, {})
    assert.is_nil(markup.section_line(line), line)
    assert.is_true(markup.same(line, spans))
  end)
end)

for _, ext in ipairs({ "docx", "odt" }) do
  describe(ext .. " sections", function()
    local path
    before_each(function()
      neo.config.page = true
      path = H.copy("rich." .. ext)
      H.open(path)
    end)
    after_each(function()
      neo.config.page = false
    end)
    local function section_lines()
      return vim.tbl_filter(function(l) return l:find("^\\newpage {") ~= nil end, H.lines())
    end
    -- the pages LibreOffice makes of it: { "portrait" | "landscape" } (consecutive ones once)
    local function rendered()
      local dir = vim.fn.tempname()
      vim.fn.mkdir(dir, "p")
      vim.system({ "soffice", "-env:UserInstallation=file://" .. dir .. "/p", "--headless", "--convert-to", "pdf", "--outdir", dir, path }):wait()
      local pdf = dir .. "/" .. vim.fn.fnamemodify(path, ":t:r") .. ".pdf"
      local info = vim.system({ "pdfinfo", "-f", "1", "-l", "50", pdf }):wait().stdout or ""
      local out = {}
      for w, h in info:gmatch("size:%s+([%d.]+) x ([%d.]+)") do
        local o = tonumber(w) > tonumber(h) and "landscape" or "portrait"
        if out[#out] ~= o then
          out[#out + 1] = o
        end
      end
      return out
    end

    it("shows the document's at its top, and saves it unchanged byte for byte", function()
      assert.equals("---", H.lines()[1])
      assert.truthy(H.lines()[2]:match("^page: {size = %S+"), H.lines()[2])
      assert.same({}, section_lines())
      local before = H.bytes(path)
      vim.bo.modified = true
      H.save()
      assert.equals(before, H.bytes(path))
    end)

    it("starts a section where its line is, and joins it to the one before when the line is deleted", function()
      local r = H.row(3)
      api.nvim_buf_set_lines(0, r - 1, r - 1, false, { "\\newpage {orientation = landscape}", "" })
      H.reopen()
      assert.same({ "\\newpage {orientation = landscape}" }, section_lines())
      assert.is_true(require("neoffice.xml").check(zip.read(path, H.MAIN[ext])))
      if H.has_libreoffice and vim.fn.executable("pdfinfo") == 1 then
        assert.same({ "portrait", "landscape" }, rendered())
      end
      -- (text edited: the sections as they were)
      H.replace(1, "Introduction", "Intro")
      H.reopen()
      assert.same({ "\\newpage {orientation = landscape}" }, section_lines())
      local before = H.bytes(path)
      vim.bo.modified = true
      H.save()
      assert.equals(before, H.bytes(path))
      -- (its line deleted: one section)
      local at = vim.fn.index(H.lines(), "\\newpage {orientation = landscape}")
      api.nvim_buf_set_lines(0, at, at + 2, false, {})
      H.reopen()
      assert.same({}, section_lines())
      if H.has_libreoffice and vim.fn.executable("pdfinfo") == 1 then
        assert.same({ "portrait" }, rendered())
      end
    end)

    it("changes the document's settings, and those of the sections that don't say otherwise", function()
      local r = H.row(3)
      api.nvim_buf_set_lines(0, r - 1, r - 1, false, { "\\newpage {orientation = landscape, margins = 1cm}", "" })
      H.reopen()
      api.nvim_buf_set_lines(0, 1, 2, false, { "page: {size = A4, margins = 2cm}" })
      H.reopen()
      assert.equals("page: {size = A4, margins = 2cm}", H.lines()[2])
      assert.same({ "\\newpage {orientation = landscape, margins = 1cm}" }, section_lines())
      local main = zip.read(path, ext == "docx" and H.MAIN.docx or "styles.xml")
      if ext == "docx" then
        assert.truthy(main:find('<w:pgSz w:w="11906" w:h="16838"/>', 1, true), main:sub(-600))
        assert.truthy(main:find('<w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/>', 1, true))
      else
        assert.truthy(main:find('fo:page%-width="29.7cm"') and main:find('fo:page%-width="21cm"'), main)
      end
    end)

    it("doesn't save settings it can't read, a section's line with nothing after it, or two in a row: says so, changes nothing", function()
      local before = H.bytes(path)
      local errors = {}
      local notify = vim.notify
      vim.notify = function(msg) errors[#errors + 1] = msg end
      local r = H.row(3)
      api.nvim_buf_set_lines(0, r - 1, r - 1, false, { "\\newpage {colour = red}", "" })
      H.save()
      assert.equals(before, H.bytes(path))
      assert.truthy(table.concat(errors, "\n"):find("not a page's attribute: colour", 1, true))
      api.nvim_buf_set_lines(0, r - 1, r, false, { "\\newpage {landscape}", "", "\\newpage {portrait}" })
      H.save()
      assert.equals(before, H.bytes(path))
      assert.truthy(table.concat(errors, "\n"):find("two sections' lines", 1, true))
      api.nvim_buf_set_lines(0, r - 1, r + 2, false, {})
      api.nvim_buf_set_lines(0, -1, -1, false, { "", "\\newpage {landscape}" })
      H.save()
      assert.equals(before, H.bytes(path))
      assert.truthy(table.concat(errors, "\n"):find("nothing after a section's line", 1, true))
      if ext == "odt" then
        api.nvim_buf_set_lines(0, -3, -1, false, {})
        api.nvim_buf_set_lines(0, r - 1, r - 1, false, { "\\newpage {start = continuous}", "" })
        H.save()
        assert.equals(before, H.bytes(path))
        assert.truthy(table.concat(errors, "\n"):find("continuous is Word's", 1, true))
      end
      vim.notify = notify
    end)
  end)
end

describe("odt a first paragraph naming its page style, with a shape holding an empty paragraph (as LibreOffice writes a form)", function()
  local SHAPED = function(name, text)
    if name ~= "content.xml" then
      return text
    end
    text = H.default_styles(name, text)
    local style = '<style:style style:name="Start" style:family="paragraph" style:parent-style-name="Standard" style:master-page-name="Standard"/>'
    local p = '<text:p text:style-name="Start"><draw:custom-shape text:anchor-type="paragraph" svg:width="5cm" svg:height="1cm" svg:x="0cm" svg:y="0cm"><text:p/></draw:custom-shape>A form</text:p>'
    return (text:gsub("<office:automatic%-styles>", "%0" .. style, 1):gsub("<office:text>", "%0" .. p, 1))
  end
  it("opens, with no section of its own, and saves unchanged", function()
    neo.config.page = true
    local path = H.copy("rich.odt", SHAPED)
    H.open(path)
    neo.config.page = false
    assert.truthy(vim.tbl_contains(H.paras(), "A form"), table.concat(H.lines(), "\n"))
    assert.same({}, vim.tbl_filter(function(l) return l:find("^\\newpage {") ~= nil end, H.lines()))
    local before = H.bytes(path)
    vim.bo.modified = true
    H.save()
    assert.equals(before, H.bytes(path))
  end)
end)
