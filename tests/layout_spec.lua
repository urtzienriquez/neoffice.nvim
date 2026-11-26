-- Alignment, quotes, line and page breaks, strikethrough, highlight, footnotes and fields.
local H = dofile("tests/helpers.lua")
local zip = require("neoffice.zip")

local EXPECTED = {
  "<|> Centered text.",
  "->> Right text.",
  "<=> Justified text.",
  "> A quote.",
  "First line\\\nsecond line.",
  "Before the page break.",
  "\\newpageAfter the page break.",
  "Some ~~struck~~ and ==lit== words.",
  "A claim^[The first note.] and another^[The second note.].",
  "See ⟦(Smith, 2020)⟧ for that.",
}

for _, name in ipairs({ "layout.docx", "layout.odt" }) do
  local ext = name:match("%w+$")
  describe(name, function()
    local path
    before_each(function()
      path = H.copy(name)
      H.open(path)
    end)

    it("shows alignment, quotes, breaks, formatting, footnotes and fields", function()
      assert.same(EXPECTED, H.paras())
    end)

    it("highlights the alignment marks (over Markdown's quote highlighting)", function()
      local function hl(row, col)
        for _, x in ipairs(vim.inspect_pos(0, row, col).extmarks) do
          if x.opts.hl_group == "NeofficeAlign" then
            return true
          end
        end
        return false
      end
      assert.equals("NeofficeMark", vim.api.nvim_get_hl(0, { name = "NeofficeAlign" }).link) -- (formatting's gray)
      for _, row in ipairs({ H.row(1), H.row(2), H.row(3) }) do
        assert.same({ true, true, true, false }, { hl(row - 1, 0), hl(row - 1, 1), hl(row - 1, 2), hl(row - 1, 3) })
      end
    end)

    it("saves without edits byte for byte", function()
      local before = H.bytes(path)
      vim.bo.modified = true
      H.save()
      assert.equals(before, H.bytes(path))
    end)

    it("changes alignment, and only that paragraph", function()
      local before = H.paragraph_xml(path)
      H.set(1, "Centered text.")
      H.set(2, "<|> Right text.")
      H.set(3, "->> Justified text.")
      H.set(6, "<=> Before the page break.")
      H.reopen()
      assert.same({ "Centered text.", "<|> Right text.", "->> Justified text." }, vim.list_slice(H.paras(), 1, 3))
      assert.equals("<=> Before the page break.", H.paras()[6])
      assert.equals(4, H.differing(before, H.paragraph_xml(path)))
    end)

    it("aligns headings, titles, quotes and list items", function()
      H.set(4, "> <|> A quote.")
      H.set(6, "# ->> Before the page break.")
      H.reopen()
      assert.equals("> <|> A quote.", H.paras()[4])
      assert.equals("# ->> Before the page break.", H.paras()[6])
    end)

    it("makes and unmakes quotes", function()
      H.set(4, "A quote.")
      H.set(6, "> Before the page break.")
      H.reopen()
      assert.equals("A quote.", H.paras()[4])
      assert.equals("> Before the page break.", H.paras()[6])
    end)

    it("escapes text that looks like an alignment or a quote", function()
      H.set(6, "\\->> not right")
      H.set(7, "\\> not a quote")
      H.reopen()
      assert.equals("\\->> not right", H.paras()[6])
      assert.equals("\\> not a quote", H.paras()[7])
    end)

    it("adds and removes line breaks", function()
      local r = H.row(5)
      vim.api.nvim_buf_set_lines(0, r - 1, r + 1, false, { "First line second line." })
      H.set(6, "Before\\\nthe\\\npage break.")
      H.reopen()
      assert.equals("First line second line.", H.paras()[5])
      assert.equals("Before\\\nthe\\\npage break.", H.paras()[6])
      assert.equals(10, #H.paras())
    end)

    it("adds and removes page breaks", function()
      H.set(7, "After the page break.")
      H.set(8, "\\newpageSome ~~struck~~ and ==lit== words.")
      H.reopen()
      assert.equals("After the page break.", H.paras()[7])
      assert.equals("\\newpageSome ~~struck~~ and ==lit== words.", H.paras()[8])
    end)

    it("adds and removes strikethrough and highlight", function()
      H.set(8, "~~Some~~ struck and lit ==words==.")
      H.reopen()
      assert.equals("~~Some~~ struck and lit ==words==.", H.paras()[8])
    end)

    it("highlights in colours, and underlines with ++", function()
      H.set(8, "Some ==green=={green}, ==dark=={darkblue}, ==hex=={#00ff00} and ++under++ words.")
      H.reopen()
      -- (#00ff00 is green; in .docx an unnamed colour becomes the nearest of Word's)
      assert.equals("Some ==green=={green}, ==dark=={darkblue}, ==hex=={green} and ++under++ words.", H.paras()[8])
      local main = zip.read(H.doc().path, H.MAIN[ext])
      if ext == "docx" then
        assert.truthy(main:find('<w:highlight w:val="darkBlue"/>', 1, true))
      else
        assert.truthy(main:find('fo:background-color="#000080"', 1, true))
      end
      require("neoffice").decorate(0) -- (shown in their colours; underlines by the syntax)
      local shown = {}
      for _, m in ipairs(vim.api.nvim_buf_get_extmarks(0, -1, { H.row(8) - 1, 0 }, { H.row(8) - 1, -1 }, { details = true })) do
        local group = m[4].hl_group
        if group and group:match("^Neoffice[HU]") then
          shown[#shown + 1] = group .. ":" .. H.lines()[H.row(8)]:sub(m[3] + 1, m[4].end_col)
        end
      end
      table.sort(shown)
      assert.same({ "NeofficeHighlight_000080:dark", "NeofficeHighlight_00ff00:green", "NeofficeHighlight_00ff00:hex" }, shown)
      H.set(8, "Some ==green== words.") -- (back to yellow)
      H.reopen()
      assert.equals("Some ==green== words.", H.paras()[8])
    end)

    it("edits, adds and deletes footnotes", function()
      H.set(9, "A claim^[The first note, edited.] and another. New^[A new note].")
      H.reopen()
      assert.equals("A claim^[The first note, edited.] and another. New^[A new note].", H.paras()[9])
      if ext == "docx" then -- (the deleted footnote is gone from footnotes.xml too)
        local notes = zip.read(path, "word/footnotes.xml")
        assert.falsy(notes:find("The second note", 1, true))
        assert.truthy(notes:find("A new note", 1, true))
      end
    end)

    it("adds a footnote to a document that had none", function()
      local other = H.copy("rich." .. ext)
      H.open(other)
      H.set(12, "Last paragraph.^[A first footnote.]")
      H.reopen()
      assert.equals("Last paragraph.^[A first footnote.]", H.paras()[12])
      if H.has_libreoffice then
        assert.is_true((H.libreoffice_reads(other, "A first footnote.", "html")))
      end
    end)

    it("refuses to write a field whose text was edited: the save says so and changes nothing", function()
      local field = ext == "docx" and "ZOTERO_ITEM" or "fieldmark"
      local before = H.bytes(path)
      H.replace(10, "(Smith, 2020)", "(Smith, 2021)")
      H.replace(10, "for that", "for this")
      local errors = {}
      local notify = vim.notify
      vim.notify = function(msg, level) errors[#errors + 1] = { msg, level } end
      H.save()
      vim.notify = notify
      assert.equals(before, H.bytes(path))
      assert.equals(1, #errors)
      assert.equals(vim.log.levels.ERROR, errors[1][2])
      assert.truthy(errors[1][1]:find("⟦(Smith, 2021)⟧ is ⟦(Smith, 2020)⟧", 1, true), errors[1][1])
      H.replace(10, "(Smith, 2021)", "(Smith, 2020)") -- (put back: the other edit is written)
      H.reopen()
      assert.equals("See ⟦(Smith, 2020)⟧ for this.", H.paras()[10])
      assert.truthy(zip.read(path, H.MAIN[ext]):find(field, 1, true))
    end)

    it("deletes a field with its text", function()
      H.replace(10, "⟦(Smith, 2020)⟧ ", "")
      H.reopen()
      assert.equals("See for that.", H.paras()[10])
    end)

    it("with tracking on, keeps footnotes, fields and breaks you delete (and tracks the text)", function()
      vim.cmd("Neoffice track on")
      H.replace(9, " and another^[The second note.]", "")
      H.replace(5, "\\", "")
      H.replace(10, "⟦(Smith, 2020)⟧", "")
      vim.cmd("Neoffice accept all")
      H.reopen()
      assert.same({ "First line\\\nsecond line.", "A claim^[The first note.]^[The second note.].", "See ⟦(Smith, 2020)⟧ for that." },
        { H.paras()[5], H.paras()[9], H.paras()[10] })
    end)

    it("with tracking on, changes alignment and quotes without a tracked change", function()
      vim.cmd("Neoffice track on")
      H.replace(2, "->> ", "<=> ")
      H.replace(4, "> ", "")
      assert.equals(0, #H.changes(require("neoffice").config.author))
      H.reopen() -- (now two justified paragraphs one after the other: under one line of their format)
      assert.same({ "Right text.", "A quote." }, { H.paras()[2], H.paras()[4] })
      assert.equals("{align = justify}", H.lines()[H.row(2) - 1])
    end)

    if H.has_libreoffice then
      it("writes what LibreOffice reads", function()
        H.set(1, "->> Right now.")
        H.set(9, "A claim^[Edited note text.] and another^[The second note.].")
        H.set(5, "One\\\ntwo.")
        H.save()
        local ok, text = H.libreoffice_reads(path, "Right now.")
        assert.is_true(ok, text)
        assert.is_true((H.libreoffice_reads(path, "Edited note text", "html")))
      end)
    end
  end)
end

describe("colours", function()
  it("show formatting (marks, attributes, backslashes, tabs) in one gray, and text as text", function()
    H.open(H.copy("rich.docx"))
    H.set(1, "<|> {first = 1cm} Some **bold**, \\*a star\\* and ==marked== text.")
    local neo = require("neoffice")
    neo.decorate(0)
    local ns = vim.api.nvim_create_namespace("neoffice_decorations")
    local row, line = H.row(1), H.lines()[H.row(1)]
    local gray = {}
    for _, m in ipairs(vim.api.nvim_buf_get_extmarks(0, ns, { row - 1, 0 }, { row - 1, -1 }, { details = true })) do
      if m[4].hl_group == "NeofficeMark" then
        gray[#gray + 1] = line:sub(m[3] + 1, m[4].end_col)
      end
    end
    assert.same({ "<|> {first = 1cm} ", "**", "**", "\\", "\\", "==", "==" }, gray)
    -- what says what a paragraph is: not gray (a heading's #, a list's -, a table's |)
    H.set(1, "# A **heading**")
    H.set(2, "- An item")
    neo.decorate(0)
    for _, m in ipairs(vim.api.nvim_buf_get_extmarks(0, ns, 0, -1, { details = true })) do
      local l = H.lines()[m[2] + 1]
      if m[4].hl_group == "NeofficeMark" then
        assert.is_nil(l:sub(m[3] + 1, m[4].end_col):find("^[#%-|] "), l)
        assert.are_not.equal("|", l:sub(m[3] + 1, m[3] + 1))
      end
    end
    -- (a tab: a → in the same gray, whatever the window's colours; braces with no mark: gray too)
    H.set(1, "{size = 9pt} A\tB **c**.")
    vim.wo.winhighlight = "" -- (as some colour schemes set it)
    neo.decorate(0)
    row, line = H.row(1), H.lines()[H.row(1)]
    local arrows, grays = 0, {}
    for _, m in ipairs(vim.api.nvim_buf_get_extmarks(0, ns, { row - 1, 0 }, { row - 1, -1 }, { details = true })) do
      if m[4].virt_text and m[4].virt_text[1][1] == "→" and m[4].virt_text[1][2] == "NeofficeMark" then
        arrows = arrows + 1
        assert.equals("\t", line:sub(m[3] + 1, m[3] + 1))
      elseif m[4].hl_group == "NeofficeMark" then
        grays[#grays + 1] = line:sub(m[3] + 1, m[4].end_col)
      end
    end
    assert.equals(1, arrows)
    assert.same({ "{size = 9pt} ", "**", "**" }, grays)
    local mark = vim.api.nvim_get_hl(0, { name = "NeofficeMark", link = false })
    local normal = vim.api.nvim_get_hl(0, { name = "Normal" })
    if normal.fg and normal.bg then -- (half way between the text's colour and the background's)
      assert.is_true(mark.fg ~= normal.fg and mark.fg ~= normal.bg)
    end
  end)
end)
