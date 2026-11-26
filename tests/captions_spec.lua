-- Captions: "Table: ..." and "Figure: ...", numbered by the document.
local H = dofile("tests/helpers.lua")
local zip = require("neoffice.zip")
local markup = require("neoffice.markup")
local api = vim.api

local function add_captions() -- a figure's caption under the image, a table's above the table
  local lines = H.lines()
  local t = vim.fn.index(lines, "| A        | B        |") + 1
  t = lines[t - 1]:match("^{") and t - 1 or t
  api.nvim_buf_set_lines(0, t - 1, t - 1, false, { "Table: Two cells.", "" })
  local image = H.row(5)
  api.nvim_buf_set_lines(0, image, image, false, { "", "Figure: A dot." })
end

for _, ext in ipairs({ "docx", "odt" }) do
  describe(ext, function()
    it("writes captions: their word, a number the document counts, and the text", function()
      H.open(H.copy("rich." .. ext))
      add_captions()
      H.reopen()
      assert.truthy(vim.tbl_contains(H.paras(), "Table: Two cells."))
      assert.truthy(vim.tbl_contains(H.paras(), "Figure: A dot."))
      local main = zip.read(H.doc().path, H.MAIN[ext])
      if ext == "docx" then
        assert.truthy(main:find('<w:r><w:t xml:space="preserve">Table </w:t></w:r><w:fldSimple w:instr=" SEQ Table \\* ARABIC ">', 1, true))
        assert.truthy(main:find('<w:pStyle w:val="TableCaption"/><w:keepNext/>', 1, true)) -- (pandoc's style, kept with its table)
        assert.truthy(main:find('<w:pStyle w:val="ImageCaption"/>', 1, true))
      else
        assert.truthy(main:find('Table <text:sequence [^>]*text:name="Table"[^>]*>1</text:sequence>. Two cells.'))
        assert.truthy(main:find('Figure <text:sequence [^>]*text:name="Figure"[^>]*>1</text:sequence>. A dot.'))
      end
    end)

    it("numbers captions in order, again when one is added before the others", function()
      H.open(H.copy("rich." .. ext))
      add_captions()
      H.save()
      local r = H.row(2)
      api.nvim_buf_set_lines(0, r, r, false, { "", "Table: A first one." }) -- (before the other)
      H.reopen()
      local main = zip.read(H.doc().path, H.MAIN[ext])
      local numbers = {}
      for n in main:gmatch(ext == "docx" and 'SEQ Table \\%* ARABIC "><w:r><w:rPr><w:noProof/></w:rPr><w:t>(%d+)<' or 'text:name="Table"[^>]*>(%d+)<') do
        numbers[#numbers + 1] = tonumber(n)
      end
      assert.same({ 1, 2 }, numbers)
      require("neoffice").decorate(0) -- (shown in the buffer too: "Table 1: ...")
      local shown = {}
      for _, m in ipairs(api.nvim_buf_get_extmarks(0, -1, 0, -1, { details = true })) do
        if m[4].virt_text and m[4].virt_text[1][2] == "NeofficeNumber" then
          shown[#shown + 1] = H.lines()[m[2] + 1]:sub(1, m[3]) .. m[4].virt_text[1][1]
        end
      end
      assert.same({ "Table 1", "Figure 1", "Table 2" }, shown)
    end)

    it("turns a caption back into a paragraph", function()
      H.open(H.copy("rich." .. ext))
      add_captions()
      H.save()
      local k = vim.fn.index(H.paras(), "Table: Two cells.") + 1
      H.set(k, "Two cells.")
      H.reopen()
      assert.truthy(vim.tbl_contains(H.paras(), "Two cells."))
      local main = zip.read(H.doc().path, H.MAIN[ext])
      assert.is_nil(main:find(ext == "docx" and "SEQ Table" or '<text:sequence [^>]*text:name="Table"')) -- (its number gone)
    end)
  end)
end

describe("captions in Word documents", function()
  it("reads Word's own (a complex field) and those numbered by hand, and keeps their separator", function()
    local path = H.copy("rich.docx", function(name, text)
      if name ~= "word/document.xml" then
        return text
      end
      local word = '<w:p><w:pPr><w:pStyle w:val="Caption"/></w:pPr><w:r><w:t xml:space="preserve">Table </w:t></w:r>'
        .. '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> SEQ Table \\* ARABIC </w:instrText></w:r>'
        .. '<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>7</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>'
        .. '<w:r><w:t xml:space="preserve">: Made in Word.</w:t></w:r></w:p>'
      local by_hand = '<w:p><w:pPr><w:pStyle w:val="TableCaption"/></w:pPr><w:r><w:t xml:space="preserve">Figure 3 – By hand.</w:t></w:r></w:p>'
      return (text:gsub("<w:body>", "<w:body>" .. word .. by_hand, 1))
    end)
    H.open(path)
    assert.same({ "Table: Made in Word.", "Figure: By hand." }, vim.list_slice(H.paras(), 1, 2))
    H.set(1, "Table: Made in Word, edited.")
    H.reopen()
    assert.equals("Table: Made in Word, edited.", H.paras()[1])
    local main = zip.read(path, "word/document.xml")
    assert.truthy(main:find('<w:t>1</w:t></w:r></w:fldSimple><w:r><w:t xml:space="preserve">: </w:t></w:r><w:r><w:t xml:space="preserve">Made in Word, edited.', 1, true))
  end)
end)

describe("caption markup", function()
  it("reads Table: and Figure: as captions, and writes a paragraph that only looks like one escaped", function()
    assert.same({ caption = "table" }, (markup.block("Table: Sprint speeds.")))
    assert.same({ caption = "figure", align = "center" }, (markup.block("Figure: <|> A plot.")))
    assert.same({}, (markup.block("Note: not a caption.")))
    local line = markup.render({ { text = "Table: just text" } }, {})
    assert.equals("\\Table: just text", line)
    local block, chars = markup.parse(line)
    assert.same({}, block)
    assert.equals("Table: just text", table.concat(vim.tbl_map(function(c) return c.ch end, chars)))
  end)
end)
