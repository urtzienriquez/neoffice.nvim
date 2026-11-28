-- Tables of contents (and lists of figures, of tables): one line, \tableofcontents, kept as the
-- document has them (their entries aren't the user's text to edit: Word and LibreOffice make them);
-- deleting the line removes one, typing it adds one (its entries the headings).
local H = dofile("tests/helpers.lua")
local zip = require("neoffice.zip")
local api = vim.api

local TOC = { docx = "<w:sdt>.-</w:sdt>", odt = "<text:table%-of%-content .-</text:table%-of%-content>" }
local function said_by(fn)
  local said, notify = {}, vim.notify
  vim.notify = function(msg) said[#said + 1] = msg end
  local ok, err = pcall(fn)
  vim.notify = notify
  assert(ok, err)
  return table.concat(said, "\n")
end
local function row_of(text)
  for r, l in ipairs(H.lines()) do
    if l == text then
      return r
    end
  end
end
local function count(text)
  return #vim.tbl_filter(function(l) return l:find(text, 1, true) ~= nil end, H.lines())
end

for _, case in ipairs({ { "toc.docx", "docx" }, { "toc.odt", "odt" }, { "toc_pandoc.docx", "docx" } }) do
  local name, ext = case[1], case[2]
  describe(name .. ", a table of contents", function()
    it("shows as one line, its entries not as text, and saves unchanged", function()
      local path = H.copy(name)
      H.open(path)
      assert.equals(1, count("\\tableofcontents"), table.concat(H.lines(), "\n"))
      assert.equals(1, count("Introduction"), table.concat(H.lines(), "\n")) -- (the heading only)
      local before = H.bytes(path)
      vim.bo.modified = true
      H.save()
      assert.equals(before, H.bytes(path))
    end)

    it("is kept exactly as it is when the text around it is edited", function()
      local path = H.copy(name)
      H.open(path)
      local xml = require("neoffice.xml")
      local function written(s) return xml.serialize(xml.parse(s)) end -- (as neoffice writes any part: "<a/>" for "<a></a>")
      local toc = written(zip.read(path, H.MAIN[ext]):match(TOC[ext]))
      H.replace(vim.fn.index(H.paras(), "First text.") + 1, "First", "Edited")
      H.reopen()
      assert.equals(toc, written(zip.read(path, H.MAIN[ext]):match(TOC[ext])))
      assert.truthy(vim.tbl_contains(H.paras(), "Edited text."))
    end)

    it("is kept exactly as it is when the default style changes at the top", function()
      local neo = require("neoffice")
      neo.config.styles = true
      local ok, err = pcall(function()
        local path = H.copy(name)
        H.open(path)
        local xml = require("neoffice.xml")
        local function written(s) return xml.serialize(xml.parse(s)) end
        local toc = written(zip.read(path, H.MAIN[ext]):match(TOC[ext]))
        local root = require("neoffice.markup").header(H.lines()).styles[1]
        local line = H.lines()[root.row]
        api.nvim_buf_set_lines(0, root.row - 1, root.row, false, { line:find("size =", 1, true) and (line:gsub("size = [^,}]+", "size = 13pt")) or (line:gsub("{", "{size = 13pt, ", 1)) })
        H.reopen()
        assert.equals(toc, written(zip.read(path, H.MAIN[ext]):match(TOC[ext])))
        assert.equals(1, count("\\tableofcontents"))
      end)
      neo.config.styles = false
      assert(ok, err)
    end)

    it("goes when its line is deleted (a valid file)", function()
      local path = H.copy(name)
      H.open(path)
      local r = row_of("\\tableofcontents")
      api.nvim_buf_set_lines(0, r - 1, r + 1, false, {})
      H.reopen()
      assert.equals(0, count("\\tableofcontents"))
      assert.is_nil(zip.read(path, H.MAIN[ext]):match(TOC[ext]))
      assert.is_true(require("neoffice.xml").check(zip.read(path, H.MAIN[ext])))
      if H.has_libreoffice then
        local ok, text = H.libreoffice_reads(path, "First text.")
        assert.is_true(ok, text)
      end
    end)

    it("typed over (its line replaced by text) goes, and the text is a paragraph", function()
      local path = H.copy(name)
      H.open(path)
      local r = row_of("\\tableofcontents")
      api.nvim_buf_set_lines(0, r - 1, r, false, { "Text where it was." })
      H.reopen()
      assert.equals(0, count("\\tableofcontents"))
      assert.truthy(vim.tbl_contains(H.paras(), "Text where it was."))
      assert.is_nil(zip.read(path, H.MAIN[ext]):match(TOC[ext]))
    end)

    if ext == "docx" then
      it("tells Word to fill it again when a heading it lists changes (and only then)", function()
        local path = H.copy(name)
        H.open(path)
        H.replace(vim.fn.index(H.paras(), "First text.") + 1, "First", "Edited")
        H.save()
        assert.is_nil((zip.read(path, "word/settings.xml") or ""):find("updateFields", 1, true))
        H.replace(vim.fn.index(H.paras(), "# Methods") + 1, "Methods", "Materials and methods")
        H.save()
        assert.truthy((zip.read(path, "word/settings.xml") or ""):find('<w:updateFields w:val="true"/>', 1, true))
      end)
    end
  end)
end

for _, ext in ipairs({ "docx", "odt" }) do
  describe(ext .. " \\tableofcontents typed", function()
    it("adds one, its entries the headings, read back as one line, read by LibreOffice as an index", function()
      local path = H.copy("rich." .. ext)
      H.open(path)
      api.nvim_buf_set_lines(0, 0, 0, false, { "\\tableofcontents", "" })
      H.reopen()
      assert.equals("\\tableofcontents", H.lines()[1])
      assert.equals(1, count("Introduction"))
      local main = zip.read(path, H.MAIN[ext])
      local toc = main:match(TOC[ext]) or main:match(ext == "docx" and "<w:sdt>.-</w:sdt>" or "<text:table%-of%-content.-</text:table%-of%-content>")
      assert.truthy(toc, main:sub(1, 1500))
      assert.truthy(toc:find("Introduction", 1, true), toc)
      if ext == "docx" then
        local anchor = toc:match('w:anchor="([^"]+)"')
        assert.truthy(main:find('<w:bookmarkStart w:id="%d+" w:name="' .. vim.pesc(anchor) .. '"/>'), "the heading's bookmark")
        assert.truthy((zip.read(path, "word/settings.xml") or ""):find("updateFields", 1, true))
      end
      assert.is_true(require("neoffice.xml").check(main))
      local before = H.bytes(path)
      vim.bo.modified = true
      H.save()
      assert.equals(before, H.bytes(path)) -- (as written: kept, not made again)
      if H.has_libreoffice then
        local dir = vim.fn.tempname()
        vim.fn.mkdir(dir, "p")
        vim.system({ "soffice", "-env:UserInstallation=file://" .. dir .. "/p", "--headless", "--convert-to", "odt", "--outdir", dir, path }):wait()
        local odt = dir .. "/" .. vim.fn.fnamemodify(path, ":t:r") .. ".odt"
        assert.truthy((zip.read(odt, "content.xml") or ""):find("<text:table-of-content", 1, true), "LibreOffice reads it as a table of contents")
      end
    end)

    it("is highlighted as a page break is", function()
      H.open(H.copy("rich." .. ext))
      api.nvim_buf_set_lines(0, 0, 0, false, { "\\tableofcontents", "" })
      require("neoffice").decorate(0)
      local groups = vim.tbl_map(function(m) return m[4].hl_group end, api.nvim_buf_get_extmarks(0, -1, { 0, 0 }, { 0, -1 }, { details = true }))
      assert.truthy(vim.tbl_contains(groups, "neofficeBreak"), vim.inspect(groups))
    end)

    it("isn't put in a table cell or a box: the save says so and changes nothing", function()
      local path = H.copy("rich." .. ext)
      H.open(path)
      local before = H.bytes(path)
      local n = api.nvim_buf_line_count(0)
      api.nvim_buf_set_lines(0, n, n, false, { "", "::: {width = 5cm}", "In a box.", "", "\\tableofcontents", ":::" })
      local said = said_by(H.save)
      assert.truthy(said:find("not in a table cell or a box", 1, true), said)
      assert.equals(before, H.bytes(path))
    end)

    it("as text in a paragraph is text", function()
      local path = H.copy("rich." .. ext)
      H.open(path)
      api.nvim_buf_set_lines(0, 0, 0, false, { "See \\tableofcontents here.", "", "\\\\tableofcontents", "" })
      H.reopen()
      assert.equals(0, count("\\tableofcontents") - count("See \\tableofcontents here.") - count("\\\\tableofcontents"))
      assert.is_nil(zip.read(path, H.MAIN[ext]):match(TOC[ext]))
    end)
  end)
end

describe("docx a table of contents not in a content control (a TOC field over its paragraphs, as Word writes a list of figures)", function()
  local function bare(kind)
    return H.copy("toc.docx", function(name, text)
      if name ~= "word/document.xml" then
        return text
      end
      text = text:gsub("<w:sdt><w:sdtPr>.-</w:sdtPr><w:sdtContent>", ""):gsub("</w:sdtContent></w:sdt>", "")
      if kind then
        text = text:gsub("TOC \\z \\o &quot;1%-3&quot; \\u \\h", 'TOC \\h \\z \\c &quot;' .. kind .. '&quot;')
      end
      return text
    end)
  end

  it("shows as one line, keeps all its paragraphs, and deleting the line deletes all of them", function()
    local path = bare()
    H.open(path)
    assert.equals(1, count("\\tableofcontents"), table.concat(H.lines(), "\n"))
    assert.equals(1, count("Introduction"))
    local before = H.bytes(path)
    vim.bo.modified = true
    H.save()
    assert.equals(before, H.bytes(path))
    local r = row_of("\\tableofcontents")
    api.nvim_buf_set_lines(0, r - 1, r + 1, false, {})
    H.reopen()
    local main = zip.read(path, "word/document.xml")
    assert.is_nil(main:find("TOC1", 1, true), main:sub(1, 1500)) -- (none of its entries left)
    assert.is_nil(main:find("fldChar", 1, true))
    assert.is_true(require("neoffice.xml").check(main))
  end)

  it("a paragraph typed right after its line goes after all of it", function()
    local path = bare()
    H.open(path)
    local r = row_of("\\tableofcontents")
    api.nvim_buf_set_lines(0, r, r, false, { "", "Right after it." })
    H.reopen()
    local lines = H.lines()
    assert.truthy(row_of("Right after it.") > row_of("\\tableofcontents"), table.concat(lines, "\n"))
    assert.equals(1, count("\\tableofcontents"))
    local main = zip.read(path, "word/document.xml")
    local after = main:find("Right after it.", 1, true)
    local last_entry = select(2, main:find('fldCharType="end"/></w:r></w:p>', 1, true))
    assert.truthy(after and last_entry and after > last_entry, "after the field's end")
  end)

  it("with \\c \"Figure\", a list of figures: \\listoffigures, kept; a new one can't be made: the save says so", function()
    local path = bare("Figure")
    H.open(path)
    assert.equals(1, count("\\listoffigures"), table.concat(H.lines(), "\n"))
    local before = H.bytes(path)
    api.nvim_buf_set_lines(0, 0, 0, false, { "\\listoftables", "" })
    local said = said_by(H.save)
    assert.truthy(said:find("can't make a new one", 1, true), said)
    assert.equals(before, H.bytes(path))
  end)
end)

if H.has_libreoffice then
  describe("the PDF of a document with a table of contents", function()
    for _, ext in ipairs({ "docx", "odt" }) do
      it("has its entries and page numbers (" .. ext .. ")", function()
        require("neoffice").config.pdf = { auto = false, open = false, quiet = true, timeout = 120 }
        local path = H.copy("rich." .. ext)
        H.open(path)
        local n = api.nvim_buf_line_count(0)
        api.nvim_buf_set_lines(0, n, n, false, { "", "\\newpage", "", "# A later section", "", "Its text." })
        api.nvim_buf_set_lines(0, 0, 0, false, { "\\tableofcontents", "", "\\newpage", "" })
        H.save()
        local pdf = path:gsub("%." .. ext .. "$", ".pdf")
        vim.cmd("Neoffice pdf")
        assert.truthy(vim.wait(30000, function() return vim.uv.fs_stat(pdf) ~= nil end, 50))
        vim.wait(300)
        local text = vim.system({ "pdftotext", "-layout", pdf, "-" }):wait().stdout
        assert.truthy(text:find("A later section%.+%s*%d"), text:sub(1, 600))
      end)
    end
  end)
end

-- a table of contents' title, a paragraph in its heading's style right before it (as when the TOC
-- was made again, or the title typed): shown with it, as its line; and the TOC's styles at the top
for _, ext in ipairs({ "docx", "odt" }) do
  describe(ext .. " a table of contents' title right before it", function()
    local TITLE = {
      docx = '<w:p><w:pPr><w:pStyle w:val="TOCHeading"/></w:pPr><w:r><w:t>My own title</w:t></w:r></w:p>',
      odt = '<text:p text:style-name="Contents_20_Heading">My own title</text:p>',
    }
    local function titled()
      return H.copy("toc." .. ext, function(name, text)
        if ext == "docx" and name == "word/document.xml" then
          return (text:gsub("<w:sdt>", TITLE.docx .. "%0", 1))
        elseif ext == "odt" and name == "content.xml" then
          return (text:gsub("<text:table%-of%-content ", TITLE.odt .. "%0", 1))
        end
        return text
      end)
    end

    it("isn't shown: it is the table of contents', kept, and goes with its line", function()
      local path = titled()
      H.open(path)
      assert.equals(0, count("My own title"), table.concat(H.lines(), "\n"))
      assert.equals(1, count("\\tableofcontents"))
      local before = H.bytes(path)
      vim.bo.modified = true
      H.save()
      assert.equals(before, H.bytes(path))
      local r = row_of("\\tableofcontents")
      api.nvim_buf_set_lines(0, r - 1, r + 1, false, {})
      H.reopen()
      assert.is_nil(zip.read(path, H.MAIN[ext]):find("My own title", 1, true))
    end)

    it("typed before \\tableofcontents is the new one's title, not one more", function()
      local path = H.copy("toc." .. ext) -- (a document with the style)
      H.open(path)
      local style = ext == "docx" and "TOC Heading" or "Contents Heading"
      local r = row_of("\\tableofcontents")
      api.nvim_buf_set_lines(0, r - 1, r + 1, false, {}) -- (its own gone first)
      H.reopen()
      assert.is_nil(zip.read(path, H.MAIN[ext]):match(TOC[ext]))
      api.nvim_buf_set_lines(0, 2, 2, false, { "{style = " .. style .. "}", "Our contents", "", "\\tableofcontents", "" })
      H.reopen()
      assert.equals(0, count("Our contents"), table.concat(H.lines(), "\n"))
      local main = zip.read(path, H.MAIN[ext])
      local _, n = main:gsub("Our contents", "")
      assert.equals(ext == "odt" and 2 or 1, n) -- (.odt: and in its title's template, what LibreOffice writes when it updates it)
      assert.is_nil(main:match(TOC[ext]):find(">Contents<", 1, true)) -- (no second title)
      assert.truthy(main:match(TOC[ext]):find("Our contents", 1, true), "the title in it")
    end)

    it("and the table of contents' styles are at the top, to change how it looks", function()
      local neo = require("neoffice")
      neo.config.styles = true
      local ok, err = pcall(function()
        local path = H.copy("toc." .. ext)
        H.open(path)
        local entry = ext == "docx" and "Toc 1" or "Contents 1"
        local st = vim.iter(require("neoffice.markup").header(H.lines()).styles):find(function(s) return s.name == entry end)
        assert.truthy(st, table.concat(H.lines(), "\n"))
        local line = H.lines()[st.row]
        api.nvim_buf_set_lines(0, st.row - 1, st.row, false, { (line:gsub("}%s*$", ", size = 9pt}")) })
        H.reopen()
        local styles = zip.read(path, ext == "docx" and "word/styles.xml" or "styles.xml")
        assert.truthy(styles:find(ext == "docx" and 'w:sz w:val="18"' or 'fo:font%-size="9pt"'), "the entries' style changed")
      end)
      neo.config.styles = false
      assert(ok, err)
    end)
  end)
end
