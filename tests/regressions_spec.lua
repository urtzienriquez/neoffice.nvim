-- Bugs found by reviewing the plugin, one test each: they must stay fixed.
local H = dofile("tests/helpers.lua")
local neo = require("neoffice")
local changes = require("neoffice.changes")
local markup = require("neoffice.markup")
local xml = require("neoffice.xml")
local zip = require("neoffice.zip")
local api = vim.api

local function edit_xml(name, from, to) -- a copy of a fixture with its main part edited
  return H.copy(name, function(entry, text)
    if entry ~= H.MAIN[name:match("%w+$")] then
      return text
    end
    local new, n = text:gsub(from, to, 1)
    assert(n == 1, "fixture edit didn't apply: " .. from)
    return new
  end)
end
local function typed(r1, c1, r2, c2, text) -- (as typing does: the edit, then TextChanged)
  api.nvim_buf_set_text(0, r1, c1, r2, c2, vim.split(text, "\n"))
  changes.after_edit(api.nvim_get_current_buf())
end
local function owners(row) -- "author type:text" of the changes starting on a row (1-based)
  local out = {}
  for _, m in ipairs(api.nvim_buf_get_extmarks(0, neo.ns.change, { row - 1, 0 }, { row - 1, -1 }, { details = true })) do
    local ch = H.doc().change[m[1]]
    out[#out + 1] = ch.author .. " " .. ch.type .. ":" .. table.concat(api.nvim_buf_get_text(0, m[2], m[3], m[4].end_row, m[4].end_col, {}), "\n")
  end
  return out
end

describe("the text", function()
  for _, name in ipairs({ "empty.docx", "empty.odt" }) do
    it("keeps line breaks before marker-like text, spaces, or at the end (" .. name .. ")", function()
      local path = H.copy(name)
      H.open(path)
      api.nvim_buf_set_lines(0, 0, -1, false, { "Shopping:\\", "\\- milk", "", "Ends with a break\\", "", "Then\\", "\\   " })
      local before = H.paras()
      H.save()
      H.open(path)
      assert.same(before, H.paras())
    end)
  end

  it("keeps backslashes that escape nothing, and marks that never close", function()
    local path = H.copy("empty.docx")
    H.open(path)
    local lines = { "Path C:\\Users\\name", "", "if a == b then", "", "**", "", "|", "", "2 * 3" }
    api.nvim_buf_set_lines(0, 0, -1, false, lines)
    local function texts() -- (the text itself; how it is written may escape more)
      return vim.tbl_map(function(t)
        return table.concat(vim.tbl_map(function(c) return c.ch end, select(2, markup.parse(t))))
      end, H.paras())
    end
    local before = texts()
    H.save()
    H.open(path)
    assert.same(before, texts())
    assert.same({ "Path C:\\Users\\name", "if a == b then", "**", "|", "2 * 3" }, texts())
  end)

  it("writes text next to marks so that it reads back the same", function()
    for _, spans in ipairs({
      { { text = "x=" }, { text = "5", h = true }, { text = " end" } },
      { { text = "_" }, { text = "{x}" } },
      { { text = "a<" }, { text = "/u>" } },
      { { text = "a~" }, { text = "b", s = true } },
    }) do
      local line = markup.render(vim.deepcopy(spans), {})
      assert.is_true(markup.same(line, spans), line)
    end
  end)

  it("writes an equation's LaTeX on one line", function()
    local line = markup.render({ { math = { tex = "\\begin{matrix}\na & b\n\\end{matrix}", display = true } } }, {})
    assert.falsy(line:find("\n"))
  end)

  for action, expected in pairs({ accept = "See *et al.*, 2021; Brown, 2018) here.", reject = "See *et al.*, 2021) here." }) do
    it("tracks a replacement of text that starts with a mark (" .. action .. ")", function()
      local path = H.copy("rich.docx")
      H.open(path)
      H.set(12, "See *et al.*, 2021) here.")
      H.save()
      vim.cmd("Neoffice track on")
      local r = H.row(12)
      local s = H.lines()[r]:find("*et al.", 1, true)
      typed(r - 1, s - 1, r - 1, s - 1 + #"*et al.*, 2021)", "*et al.*, 2021; Brown, 2018)") -- (as ct) does)
      assert.equals("See *et al.*, 2021)*et al.*, 2021; Brown, 2018) here.", H.paras()[12])
      H.reopen()
      assert.equals("See *et al.*, 2021)*et al.*, 2021; Brown, 2018) here.", H.paras()[12])
      changes.resolve(action, true)
      assert.equals(expected, H.paras()[12])
    end)
  end

  it("keeps an escaped ¶ when a deletion of it is rejected", function()
    H.open(H.copy("rich.docx"))
    H.set(12, "See \\¶ 3. Last paragraph.")
    H.save()
    vim.cmd("Neoffice track on")
    local r = H.row(12)
    typed(r - 1, 0, r - 1, #"See \\¶ 3. ", "")
    changes.resolve("reject", true)
    assert.equals("See \\¶ 3. Last paragraph.", H.paras()[12])
  end)
end)

describe("tables", function()
  it("save with one column", function()
    local path = H.copy("rich.docx")
    H.open(path)
    api.nvim_buf_set_lines(0, H.row(8) - 1, H.row(11), false, { "| A |", "|---|", "| cell one |" })
    H.save()
    H.open(path)
    assert.same({ "| A |", "|---|", "| cell one |" }, vim.list_slice(H.lines(), H.row(8), H.row(8) + 2))
  end)

  for _, name in ipairs({ "rich.docx", "rich.odt" }) do
    it("open with spaces at the edges of a cell, and save them unchanged (" .. name .. ")", function()
      local path = H.copy(name, function(_, text) return (text:gsub(">cell one<", ">  cell one <")) end)
      H.open(path)
      assert.equals("cell one", H.paras()[10])
      local bytes = H.bytes(path)
      vim.bo.modified = true
      H.save()
      assert.equals(bytes, H.bytes(path))
    end)
  end

  it("keep a row of dashes, and | inside a footnote", function()
    local path = H.copy("rich.docx")
    H.open(path)
    local r = H.row(10)
    api.nvim_buf_set_lines(0, r - 1, r, false, { "| - | - |", "| a^[x \\| y] | b |" })
    H.save()
    H.open(path)
    assert.same({ "| - | - |", "| a^[x \\| y] | b |" }, vim.list_slice(H.lines(), r, r + 1))
  end)

  it("get no stray empty paragraph after a new table", function()
    local path = H.copy("empty.docx")
    H.open(path)
    api.nvim_buf_set_lines(0, 0, -1, false, { "x", "", "| a | b |", "|---|---|", "", "y" })
    H.save()
    H.open(path)
    assert.same({ "x", "", "| a | b |", "|---|---|", "", "y" }, H.lines())
  end)
end)

describe("saving", function()
  it("works with a footnote typed across two lines", function()
    local path = H.copy("rich.docx")
    H.open(path)
    api.nvim_buf_set_lines(0, -1, -1, false, { "", "Note^[a long", "note] end." })
    H.save()
    assert.is_false(vim.bo.modified)
    H.open(path)
    assert.equals("Note^[a long note] end.", H.paras()[13])
  end)

  it("leaves nothing behind when it fails: the next save writes everything once", function()
    local path = H.copy("rich.odt")
    H.open(path)
    api.nvim_buf_set_lines(0, H.row(12) - 1, H.row(12) - 1, false, { "Added once.", "" })
    local write = zip.write
    zip.write = function() error("disk full") end
    H.save()
    zip.write = write
    assert.is_true(vim.bo.modified)
    H.save()
    H.open(path)
    assert.equals(1, #vim.tbl_filter(function(t) return t == "Added once." end, H.paras()))
  end)

  it("makes an empty paragraph of its own also when the save made it", function()
    local path = H.copy("empty.docx")
    H.open(path)
    api.nvim_buf_set_lines(0, 0, -1, false, { "a", "", "", "b" })
    H.save()
    api.nvim_buf_set_text(0, 2, 0, 2, 0, { "x" })
    assert.same({ "a", "x", "b" }, H.paras())
  end)

  it("takes the new name with :saveas, and refuses another format", function()
    local dir = vim.fn.tempname()
    vim.fn.mkdir(dir, "p")
    H.open(H.copy("rich.docx"))
    H.replace(12, "Last", "Final")
    vim.cmd("silent saveas " .. dir .. "/new.docx")
    assert.is_false(vim.bo.modified)
    assert.equals(dir .. "/new.docx", H.doc().path)
    vim.cmd("silent write " .. dir .. "/out.odt")
    assert.is_nil(vim.uv.fs_stat(dir .. "/out.odt"))
  end)

  it("writes again a footnote deleted by a save and brought back by undo", function()
    local path = H.copy("layout.docx")
    H.open(path)
    H.replace(9, "^[The first note.]", "")
    H.save()
    vim.cmd("silent undo")
    changes.after_edit(api.nvim_get_current_buf())
    H.save()
    H.open(path)
    assert.equals("A claim^[The first note.] and another^[The second note.].", H.paras()[9])
  end)

  it("keeps a section break when its paragraph is deleted, or edited after a deleted one", function()
    local landscape = '<w:sectPr><w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/></w:sectPr>'
    local path = edit_xml("rich.docx", '</w:numPr></w:pPr>(<w:r><w:t xml:space="preserve">second item)', "</w:numPr>" .. landscape .. "</w:pPr>%1")
    H.open(path)
    api.nvim_buf_set_lines(0, H.row(6) - 1, H.row(7), false, { "- second item edited" })
    H.save()
    assert.truthy(zip.read(path, "word/document.xml"):find("landscape.-second item edited"))
    path = edit_xml("layout.docx", '(<w:p><w:pPr><w:pStyle w:val="BodyText" /><w:jc w:val="both"/>)', "<w:p><w:pPr>" .. landscape .. "</w:pPr></w:p>%1")
    H.open(path)
    api.nvim_buf_set_lines(0, H.row(3) - 1, H.row(3), false, {}) -- (the blank line of the empty paragraph)
    H.save()
    assert.truthy(zip.read(path, "word/document.xml"):find("landscape", 1, true))
  end)
end)

describe("tracked changes", function()
  before_each(function()
    H.open(H.copy("rich.docx"))
    vim.cmd("Neoffice track on")
  end)

  it("survive deleting the last line", function()
    local r = H.row(12)
    typed(r - 1, 5, r - 1, 15, "")
    vim.cmd("Neoffice track off")
    api.nvim_win_set_cursor(0, { r, 0 })
    vim.cmd("normal! dd")
    changes.after_edit(api.nvim_get_current_buf())
    typed(0, 0, 0, 0, "x")
    H.save()
    assert.is_false(vim.bo.modified)
  end)

  it("keep deleted text as it was: its formatting and marks", function()
    local r = H.row(2)
    local line = H.lines()[r]
    local s, e = line:find("**bold**, *italic*", 1, true)
    typed(r - 1, s - 1, r - 1, e, "")
    assert.truthy(H.lines()[r]:find("**bold**, *italic*", 1, true))
    H.save()
    assert.truthy(zip.read(H.doc().path, "word/document.xml"):find("<w:del [^>]*><w:r><w:rPr><w:b/>", 1))
    r = H.row(3) -- (a comment's text: the comment keeps its range)
    typed(r - 1, 0, r - 1, #"Commented text ", "")
    local m = api.nvim_buf_get_extmarks(0, neo.ns.comment, 0, -1, { details = true })[1]
    assert.equals("Commented text", H.lines()[m[2] + 1]:sub(m[3] + 1, m[4].end_col))
    changes.resolve("reject", true)
    H.reopen()
    assert.equals("This is **bold**, *italic*, <u>underlined</u> and a [link](https://example.org). Tom & Jerry <3, 2\\*3.", H.paras()[2])
  end)

  it("keep an image deleted with its text as it was", function()
    local path = H.doc().path
    local before = H.entries(path)
    local r = H.row(5)
    typed(r - 1, 3, r - 1, #H.lines()[r] - 1, "")
    H.save()
    assert.same({ "word/document.xml" }, H.changed(before, H.entries(path)))
    assert.truthy(zip.read(path, "word/document.xml"):find('descr=""', 1, true))
  end)

  it("go on deleting leftwards with X, as one deletion", function()
    local r = H.row(12)
    api.nvim_win_set_cursor(0, { r, 4 })
    api.nvim_exec_autocmds("CursorMoved", { buffer = 0 })
    for _ = 1, 4 do
      vim.cmd("normal! X")
      changes.after_edit(api.nvim_get_current_buf())
    end
    assert.same({ neo.config.author .. " delete:Last" }, owners(r))
  end)

  it("don't widen deletions in several places to whole words", function()
    local r = H.row(6)
    api.nvim_buf_set_text(0, r, 2, r, 3, {}) -- ("second" -> "econd")
    api.nvim_buf_set_text(0, r - 1, 2, r - 1, 3, {}) -- ("first" -> "irst"), before one TextChanged
    changes.after_edit(api.nvim_get_current_buf())
    assert.same({ neo.config.author .. " delete:f" }, owners(r))
    assert.same({ neo.config.author .. " delete:s" }, owners(r + 1))
  end)

  it("leave no empty insertion of yours behind, also after a save", function()
    local r = H.row(12)
    typed(r - 1, 0, r - 1, 0, "New ")
    H.save()
    typed(r - 1, 0, r - 1, 4, "")
    assert.same({}, owners(r))
  end)

  it("reject a formatting change over a line break", function()
    local change = '<w:rPrChange w:id="5" w:author="Bob" w:date="2024-01-02T10:00:00Z"><w:rPr/></w:rPrChange>'
    local path = edit_xml("rich.docx", '<w:r><w:t xml:space="preserve">Last paragraph.</w:t></w:r>',
      '<w:r><w:rPr><w:b/>' .. change .. '</w:rPr><w:t>Last</w:t><w:br/><w:t xml:space="preserve">paragraph.</w:t></w:r>')
    H.open(path)
    assert.equals("**Last**\\\n**paragraph.**", H.paras()[12])
    changes.resolve("reject", true)
    assert.equals("Last\\\nparagraph.", H.paras()[12])
  end)

  it("keep another author's deletion without text (a tab) when its paragraph is edited", function()
    local path = edit_xml("rich.docx", '<w:r><w:t xml:space="preserve">inside.</w:t></w:r>',
      '<w:del w:id="7" w:author="Bob" w:date="2024-01-02T10:00:00Z"><w:r><w:tab/></w:r></w:del><w:r><w:t xml:space="preserve">inside.</w:t></w:r>')
    H.open(path)
    H.replace(4, "inside", "inside!")
    H.save()
    assert.truthy(zip.read(path, "word/document.xml"):find('<w:del w:id="7"[^>]*><w:r><w:tab/>'))
  end)

  it("are reviewed in a block selection line by line", function()
    local r = H.row(4)
    local c = H.lines()[r]:find("words and", 1, true)
    vim.fn.setpos("'<", { 0, r, c, 0 })
    vim.fn.setpos("'>", { 0, r, c + 1, 0 })
    vim.fn.histadd("cmd", "")
    vim.cmd("normal! gv\22\27") -- (a block, then <Esc>: as :'<,'>Neoffice reject sees it)
    changes.resolve("reject", false, true)
    assert.equals("Text with inserted rds and deleted words inside.", H.paras()[4])
  end)
end)

describe("comments", function()
  it("are reached from the panel where they are now, and cut whole characters", function()
    H.open(H.copy("rich.docx"))
    vim.cmd("Neoffice comments")
    vim.cmd("wincmd p")
    api.nvim_buf_set_lines(0, 0, 0, false, { "", "", "", "" })
    vim.cmd("wincmd p")
    api.nvim_win_set_cursor(0, { 1, 0 })
    require("neoffice.comments").jump()
    assert.truthy(api.nvim_get_current_line():find("^Commented text"))
    neo.config.comment_text = true
    local style = require("neoffice.comments").style({ author = "A", text = ("a"):rep(49) .. "ñandú", replies = {} })
    neo.config.comment_text = false
    assert.is_true(vim.str_utf_pos(style.virt_text[1][1]) ~= nil and vim.fn.strchars(style.virt_text[1][1]) > 0)
    assert.falsy(style.virt_text[1][1]:find("\195$"))
  end)

  it("write initials as whole characters", function()
    neo.config.author = "Jörg"
    local path = H.copy("rich.docx")
    H.open(path)
    H.cursor(12, "Last")
    H.input = "Why?"
    require("neoffice.comments").add()
    H.save()
    neo.config.author = vim.env.USER or "neoffice"
    assert.truthy(zip.read(path, "word/comments.xml"):find('w:initials="Jö"', 1, true))
  end)
end)

describe("the formats", function()
  it("keep an .odt paragraph's style when it is edited", function()
    local path = H.copy("rich.odt")
    H.open(path)
    H.replace(12, "Last", "Final")
    H.save()
    assert.truthy(zip.read(path, "content.xml"):find('<text:p text:style%-name="[^"]+">Final paragraph'))
  end)

  it("write no control characters, and keep line breaks in attributes", function()
    assert.equals("ab", xml.escape("a\12b"))
    assert.equals("a&#10;b&#9;c", xml.escape("a\nb\tc", true))
    local path = H.copy("rich.docx")
    H.open(path)
    H.replace(12, "Last", "La\12st")
    H.save()
    assert.falsy(zip.read(path, "word/document.xml"):find("\12", 1, true))
  end)

  it("open images whose alt text has a line break", function()
    H.open(edit_xml("rich.docx", '<wp:docPr descr=""', '<wp:docPr descr="A red&#xA;dot"'))
    assert.truthy(H.paras()[5]:find("![A red dot]", 1, true))
  end)

  it("write a tracked paragraph mark first in its properties, and not copy section breaks", function()
    local path = edit_xml("rich.docx", '<w:pStyle w:val="BodyText" /></w:pPr><w:r><w:t xml:space="preserve">Last',
      '<w:pStyle w:val="BodyText" /><w:rPr><w:rFonts w:ascii="Arial"/></w:rPr><w:sectPr><w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/></w:sectPr></w:pPr><w:r><w:t xml:space="preserve">Last')
    H.open(path)
    vim.cmd("Neoffice track on")
    local r = H.row(12)
    typed(r - 1, #H.lines()[r], r - 1, #H.lines()[r], "\n\nBrand new.")
    H.save()
    local main = zip.read(path, "word/document.xml")
    assert.truthy(main:find("<w:rPr><w:ins ", 1, true))
    assert.equals(1, select(2, main:gsub("landscape", "")))
  end)

  it("read w:val=\"off\" as off", function()
    H.open(edit_xml("rich.docx", '<w:r><w:t xml:space="preserve">Last paragraph.', '<w:r><w:rPr><w:b w:val="off"/></w:rPr><w:t xml:space="preserve">Last paragraph.'))
    assert.equals("Last paragraph.", H.paras()[12])
  end)

  it("keep nested .odt spans", function()
    local path = H.copy("rich.odt", function(entry, text)
      if entry ~= "content.xml" then
        return text
      end
      text = text:gsub("</office:automatic%-styles>", '<style:style style:name="T9" style:family="text"><style:text-properties fo:color="#ff0000"/></style:style></office:automatic-styles>', 1)
      return (text:gsub('<text:span text:style%-name="T1">bold</text:span>', '<text:span text:style-name="T9">red <text:span text:style-name="T1">bold</text:span></text:span>', 1))
    end)
    H.open(path)
    H.replace(2, "italic", "italics")
    H.save()
    assert.truthy(zip.read(path, "content.xml"):find('<text:span text:style%-name="T9"><text:span text:style%-name="T1">bold'))
  end)

  it("read whitespace as Word and LibreOffice do", function()
    H.open(edit_xml("rich.docx", '<w:r><w:t xml:space="preserve">Last paragraph.</w:t></w:r>', "<w:r><w:t>Last</w:t></w:r><w:r><w:t>   para</w:t></w:r><w:r><w:t xml:space=\"preserve\">graph.</w:t></w:r>"))
    assert.equals("Lastparagraph.", H.paras()[12])
    H.open(edit_xml("rich.odt", '<text:p text:style%-name="Text_20_body">Last paragraph.', '<text:p text:style-name="Text_20_body">\n    Last <text:span text:style-name="T1"> para</text:span>graph.\n  '))
    assert.equals("Last **para**graph.", H.paras()[12])
  end)

  it("put a paragraph typed after a table of contents after it", function()
    local path = edit_xml("layout.docx", "<w:body>(<w:p>.-</w:p>)", "<w:body><w:sdt><w:sdtPr/><w:sdtContent>%1</w:sdtContent></w:sdt>")
    H.open(path)
    H.set(1, H.paras()[1] .. "\n\nTyped after it.")
    H.save()
    assert.truthy(zip.read(path, "word/document.xml"):find("</w:sdt><w:p>.-Typed after it"))
  end)

  it("write Word's XML declaration (with its carriage return) as it was", function()
    local path = edit_xml("rich.docx", "%?>", "?>\r\n")
    H.open(path)
    H.replace(12, "Last", "Final")
    H.save()
    local main = zip.read(path, "word/document.xml")
    assert.truthy(main:find("^<%?xml[^>]*%?>\r\n<w:document"))
    assert.falsy(main:find("&#13;", 1, true))
    if H.has_libreoffice then
      assert.is_true((H.libreoffice_reads(path, "Final paragraph.")))
    end
  end)

  it("read Word's tracked moves as changes, and write them back as moves", function()
    local D = 'w:author="Fernando" w:date="2026-10-26T19:42:00Z"'
    local path = H.copy("rich.docx", function(entry, text)
      if entry ~= "word/document.xml" then
        return text
      end
      text = text:gsub('<w:r><w:t xml:space="preserve">Last paragraph.</w:t></w:r>',
        '<w:moveTo w:id="71" ' .. D .. '><w:r><w:t xml:space="preserve">Last paragraph.</w:t></w:r></w:moveTo>', 1)
      local n
      text, n = text:gsub('<w:r><w:t xml:space="preserve">and more.</w:t></w:r>',
        '<w:moveFrom w:id="80" ' .. D .. '><w:r><w:t xml:space="preserve">and more.</w:t></w:r></w:moveFrom>', 1)
      assert(n == 1 and text:find("w:moveTo", 1, true), "fixture edit didn't apply")
      return text
    end)
    H.open(path)
    assert.falsy(table.concat(H.lines(), "\n"):find("⟦", 1, true))
    local moved = vim.tbl_filter(function(ch) return ch.moved end, H.changes("Fernando"))
    assert.same({ "delete", "insert" }, vim.fn.sort(vim.tbl_map(function(ch) return ch.type end, moved)))
    H.replace(12, "Last", "Last edited")
    H.replace(3, "Commented", "Commented!")
    H.save()
    local main = zip.read(path, "word/document.xml")
    assert.truthy(main:find('<w:moveTo w:id="71"', 1, true))
    assert.truthy(main:find('<w:moveFrom w:id="80"', 1, true))
    changes.resolve("accept", true)
    H.reopen()
    assert.equals("Last edited paragraph.", H.paras()[12])
    assert.equals("Commented! text ", H.paras()[3]) -- ("and more." moved away: gone once accepted)
  end)

  it("read a deletion inside a move as part of it: accepting both deletes its text once", function()
    local D = 'w:author="Fernando" w:date="2026-10-26T19:42:00Z"'
    local path = H.copy("rich.docx", function(entry, text)
      if entry ~= "word/document.xml" then
        return text
      end
      local n
      text, n = text:gsub('<w:r><w:t xml:space="preserve">and more.</w:t></w:r>',
        '<w:moveFrom w:id="80" ' .. D .. '><w:r><w:t xml:space="preserve">and </w:t></w:r>'
          .. '<w:del w:id="81" ' .. D .. '><w:r><w:delText xml:space="preserve">more</w:delText></w:r></w:del>'
          .. '</w:moveFrom><w:r><w:t xml:space="preserve">. Kept.</w:t></w:r>', 1)
      assert(n == 1, "fixture edit didn't apply")
      return text
    end)
    H.open(path)
    assert.equals(1, #vim.tbl_filter(function(ch) return ch.type == "delete" end, H.changes("Fernando")))
    assert.equals("Commented text and more. Kept.", H.paras()[3])
    H.save() -- (unchanged: written back as it was, the deletion inside the move too)
    assert.truthy(zip.read(path, "word/document.xml"):find('<w:moveFrom w:id="80" ' .. D .. '><w:r><w:t xml:space="preserve">and </w:t></w:r><w:del w:id="81"', 1, true))
    changes.resolve("accept", true)
    H.reopen()
    assert.equals("Commented text . Kept.", H.paras()[3])
  end)

  for action, expected in pairs({
    accept = "Commented text ^{13} and *V. aspis*; Vila vipers.",
    reject = "Commented text ^{14} and *V. aspis*, Vila *V. seoanei*.",
  }) do
    it(action .. " all keep the formatting around changes next to formatted text", function()
      local D = 'w:author="Fernando" w:date="2026-10-26T19:42:00Z"'
      local SUP, I = '<w:rPr><w:vertAlign w:val="superscript"/></w:rPr>', "<w:rPr><w:i/></w:rPr>"
      local path = H.copy("rich.docx", function(entry, text)
        if entry ~= "word/document.xml" then
          return text
        end
        local n
        text, n = text:gsub('<w:r><w:t xml:space="preserve">and more.</w:t></w:r>', table.concat({
          "<w:r>" .. SUP .. "<w:t>1</w:t></w:r>", -- (a superscript: one digit inserted, one deleted)
          '<w:ins w:id="90" ' .. D .. "><w:r>" .. SUP .. "<w:t>3</w:t></w:r></w:ins>",
          '<w:del w:id="91" ' .. D .. "><w:r>" .. SUP .. "<w:delText>4</w:delText></w:r></w:del>",
          '<w:r><w:t xml:space="preserve"> and </w:t></w:r><w:r>' .. I .. "<w:t>V. aspis</w:t></w:r>", -- (after italics)
          '<w:del w:id="92" ' .. D .. '><w:r><w:delText xml:space="preserve">, </w:delText></w:r></w:del>',
          '<w:ins w:id="93" ' .. D .. '><w:r><w:t xml:space="preserve">; </w:t></w:r></w:ins>',
          '<w:r><w:t xml:space="preserve">Vila </w:t></w:r>',
          '<w:del w:id="94" ' .. D .. "><w:r>" .. I .. "<w:delText>V. seoanei</w:delText></w:r></w:del>",
          '<w:ins w:id="95" ' .. D .. '><w:r><w:rPr><w:rPrChange w:id="96" ' .. D .. ">" .. I -- (italics removed)
            .. "</w:rPrChange></w:rPr><w:t>vipers</w:t></w:r></w:ins>",
          "<w:r><w:t>.</w:t></w:r>",
        }), 1)
        assert(n == 1, "fixture edit didn't apply")
        return text
      end)
      H.open(path)
      changes.resolve(action, true)
      assert.equals(expected, H.paras()[3])
      H.reopen()
      assert.equals(expected, H.paras()[3])
    end)
  end

  describe("a formatting change of a paragraph mark (not in the text)", function()
    local D = 'w:author="Fernando" w:date="2026-10-26T19:42:00Z"'
    local CHANGE = '<w:rPrChange w:id="97" ' .. D .. "><w:rPr><w:i/></w:rPr></w:rPrChange>"
    local function open()
      local path = H.copy("rich.docx", function(entry, text)
        if entry ~= "word/document.xml" then
          return text
        end
        local n
        text, n = text:gsub('<w:pPr><w:pStyle w:val="BodyText" /></w:pPr><w:r><w:t xml:space="preserve">Last paragraph.',
          '<w:pPr><w:pStyle w:val="BodyText" /><w:rPr><w:b/>' .. CHANGE .. '</w:rPr></w:pPr><w:r><w:t xml:space="preserve">Last paragraph.', 1)
        assert(n == 1, "fixture edit didn't apply")
        return text
      end)
      H.open(path)
      return path
    end
    local function mark(path)
      return zip.read(path, "word/document.xml"):match('<w:pStyle w:val="BodyText" ?/><w:rPr>(.-)</w:rPr></w:pPr><w:r><w:t xml:space="preserve">Last')
    end

    it("is accepted with all the changes", function()
      local path = open()
      changes.resolve("accept", true)
      H.save()
      assert.equals("<w:b/>", mark(path))
    end)

    it("is rejected with all the changes, and with all of its author's", function()
      local path = open()
      changes.resolve("reject", true, false, "Fernando")
      H.save()
      assert.equals("<w:i/>", mark(path))
    end)

    it("is left by all of someone else's, and by an accept all undone", function()
      local path = open()
      changes.resolve("accept", true, false, "Bob")
      H.save()
      assert.equals("<w:b/>" .. CHANGE, mark(path))
      changes.resolve("accept", true)
      vim.cmd("silent undo")
      changes.after_edit(api.nvim_get_current_buf()) -- (what TextChanged does)
      H.save()
      assert.truthy(mark(path):find("w:rPrChange", 1, true))
    end)
  end)

  it("writes a comment's reference once when a deletion kept in its paragraph holds one", function()
    local D = 'w:author="Fernando" w:date="2026-10-26T19:42:00Z"'
    local REF = '<w:r><w:rPr><w:rStyle w:val="CommentReference" /></w:rPr><w:commentReference w:id="0" /></w:r>'
    local path = H.copy("rich.docx", function(entry, text)
      if entry ~= "word/document.xml" then
        return text
      end
      local n
      text, n = text:gsub(vim.pesc(REF), function() return '<w:del w:id="98" ' .. D .. ">" .. REF .. "</w:del>" end, 1)
      assert(n == 1, "fixture edit didn't apply")
      return text
    end)
    H.open(path)
    H.replace(3, "and more", "and even more") -- (its paragraph written again)
    H.save()
    local _, refs = zip.read(path, "word/document.xml"):gsub("<w:commentReference ", "")
    assert.equals(1, refs)
  end)

  it("save an alt text change of an image that isn't PNG, JPEG or GIF", function()
    local path = H.copy("rich.docx")
    zip.write(path, { ["word/media/rId21.png"] = "\1\0\0\0" .. ("\0"):rep(60) })
    H.open(path)
    H.replace(5, "![]", "![a red dot]")
    H.save()
    assert.is_false(vim.bo.modified)
  end)
end)
