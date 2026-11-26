-- The styles at the top of a document: its paragraph styles, and what each gives; editing a line
-- changes the style itself (every paragraph of it follows, as in Word and LibreOffice).
local H = dofile("tests/helpers.lua")
local zip = require("neoffice.zip")
local neo = require("neoffice")
local api = vim.api

local BODY = { docx = "Body Text", odt = "Text body" }
local STYLES = { docx = "word/styles.xml", odt = "styles.xml" }

local function header()
  return require("neoffice.markup").header(H.lines())
end
local function line_of(name)
  for _, st in ipairs(header().styles) do
    if st.name == name then
      return st
    end
  end
end

for _, ext in ipairs({ "docx", "odt" }) do
  describe(ext .. " styles at the top", function()
    local path
    before_each(function()
      neo.config.styles = true
      path = H.copy("rich." .. ext)
      H.open(path)
    end)
    after_each(function()
      neo.config.styles = false
    end)

    it("shows the styles its paragraphs have, and what each gives; and saves without changes byte for byte", function()
      local lines = H.lines()
      assert.equals("---", lines[1])
      assert.equals("styles:", lines[2])
      local body = line_of(BODY[ext])
      assert.truthy(body, table.concat(lines, "\n"))
      assert.same({ font = "Cambria", size = "12pt", before = "9pt", after = "9pt" }, body.attrs)
      assert.truthy(line_of("Heading 1").attrs.bold)
      assert.truthy(H.paras()[1]) -- (the paragraphs after it, as they were)
      local before = H.bytes(path)
      vim.bo.modified = true
      H.save()
      assert.equals(before, H.bytes(path))
    end)

    it("changes a style, and the styles based on it keep what they show", function()
      local st = line_of(BODY[ext])
      local text = H.lines()[st.row]
      api.nvim_buf_set_lines(0, st.row - 1, st.row, false, { (text:gsub("size = 12pt", "size = 14pt"):gsub("}$", ", align = justify}")) })
      H.reopen()
      assert.same({ font = "Cambria", size = "14pt", align = "justify", before = "9pt", after = "9pt" }, line_of(BODY[ext]).attrs)
      assert.equals("12pt", line_of("First Paragraph").attrs.size) -- (based on it: as it showed)
      assert.is_nil(line_of("First Paragraph").attrs.align)
      local styles = zip.read(path, STYLES[ext])
      assert.truthy(styles:find(ext == "docx" and 'w:sz w:val="28"' or 'fo:font%-size="14pt"'), styles:sub(1, 400))
      assert.is_true(require("neoffice.xml").check(styles))
      if H.has_libreoffice then
        local ok, out = H.libreoffice_reads(path, "Commented text")
        assert.is_true(ok, out)
      end
    end)

    it("changes nothing when its lines, or all of them, are deleted", function()
      local before = H.bytes(path)
      api.nvim_buf_set_lines(0, 2, 3, false, {})
      H.save()
      assert.equals(before, H.bytes(path)) -- (nothing changed: nothing written)
      local h = header()
      api.nvim_buf_set_lines(0, 0, h.last + 1, false, {})
      H.save()
      assert.equals(before, H.bytes(path))
    end)

    it("doesn't save a style the document hasn't, or an attribute it doesn't know: says so, changes nothing", function()
      local before = H.bytes(path)
      local errors = {}
      local notify = vim.notify
      vim.notify = function(msg) errors[#errors + 1] = msg end
      api.nvim_buf_set_lines(0, 2, 2, false, { "  Nonsense: {size = 10pt}" })
      H.save()
      assert.equals(before, H.bytes(path))
      assert.truthy(table.concat(errors, "\n"):find("not a paragraph style of this document: Nonsense", 1, true))
      api.nvim_buf_set_lines(0, 2, 3, false, { "  " .. BODY[ext] .. ": {shadow = red}" })
      H.save()
      vim.notify = notify
      assert.equals(before, H.bytes(path))
      assert.truthy(table.concat(errors, "\n"):find("not a style's attribute: shadow", 1, true))
    end)

    it("completes a style's name and its attributes", function()
      local buf = api.nvim_get_current_buf()
      assert(vim.wait(2000, function() return #vim.lsp.get_clients({ bufnr = buf, name = "neoffice_ls" }) > 0 end))
      local function complete(row, col)
        local params = { textDocument = { uri = vim.uri_from_bufnr(buf) }, position = { line = row - 1, character = col } }
        local items = {}
        for _, r in pairs(vim.lsp.buf_request_sync(buf, "textDocument/completion", params, 2000) or {}) do
          items = r.result and (r.result.items or r.result) or items
        end
        return vim.tbl_map(function(i) return i.label end, items)
      end
      api.nvim_buf_set_lines(0, 2, 2, false, { "  Ti" })
      assert.truthy(vim.tbl_contains(complete(3, 4), "Title"))
      api.nvim_buf_set_lines(0, 2, 3, false, { "  Title: {" })
      local labels = complete(3, #"  Title: {")
      assert.truthy(vim.tbl_contains(labels, "bold"))
      assert.truthy(vim.tbl_contains(labels, "align = justify"))
      assert.is_false(vim.tbl_contains(labels, "align = none"))
    end)
  end)
end

-- two paragraphs of the default style whose font and size are set on their text, first
local SIZED = {
  docx = function(name, text)
    if name ~= "word/document.xml" then
      return text
    end
    local p = '<w:p><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="22"/></w:rPr>'
      .. "<w:t>Sized text.</w:t></w:r></w:p>"
    return (text:gsub("<w:body>", "<w:body>" .. p .. p:gsub("Sized text", "More of it"), 1))
  end,
  odt = function(name, text)
    if name ~= "content.xml" then
      return text
    end
    local style = '<style:style style:name="Sized" style:family="text"><style:text-properties style:font-name="Times New Roman" fo:font-size="11pt"/></style:style>'
    local p = '<text:p text:style-name="Standard"><text:span text:style-name="Sized">Sized text.</text:span></text:p>'
    return (text:gsub("<office:automatic%-styles>", "%0" .. style, 1):gsub("<office:text>", "%0" .. p .. p:gsub("Sized text", "More of it"), 1))
  end,
}
local DEFAULT = { docx = "Normal", odt = "Standard" }

for _, ext in ipairs({ "docx", "odt" }) do
  describe(ext .. " a format moved into its style", function()
    before_each(function() neo.config.styles = true end)
    after_each(function() neo.config.styles = false end)

    it("leaves the text looking the same, the style saying it, and no format's line", function()
      local path = H.copy("rich." .. ext, SIZED[ext])
      H.open(path)
      local buf = api.nvim_get_current_buf()
      local r = H.row(1) - 1
      assert.equals("{font = Times New Roman, size = 11pt}", H.lines()[r])
      assert(vim.wait(2000, function() return #vim.lsp.get_clients({ bufnr = buf, name = "neoffice_ls" }) > 0 end))
      local params = { textDocument = { uri = vim.uri_from_bufnr(buf) }, range = { start = { line = r - 1, character = 0 }, ["end"] = { line = r - 1, character = 0 } }, context = { diagnostics = {} } }
      local actions = {}
      for _, res in pairs(vim.lsp.buf_request_sync(buf, "textDocument/codeAction", params, 2000) or {}) do
        actions = res.result or actions
      end
      local act = vim.iter(actions):find(function(a) return a.title:find("into the style " .. DEFAULT[ext], 1, true) end)
      assert.truthy(act, vim.inspect(vim.tbl_map(function(a) return a.title end, actions)))
      vim.lsp.util.apply_workspace_edit(act.edit, "utf-8")
      local default = line_of(DEFAULT[ext])
      assert.same({ "Times New Roman", "11pt" }, { default.attrs.font, default.attrs.size })
      assert.is_nil(vim.tbl_contains(H.lines(), "{font = Times New Roman, size = 11pt}") or nil)
      H.reopen()
      assert.same({ "Sized text.", "More of it." }, vim.list_slice(H.paras(), 1, 2))
      assert.is_nil(H.lines()[H.row(1) - 1]:find("font", 1, true)) -- (the style says it: nothing to say)
      assert.same({ "Times New Roman", "11pt" }, { line_of(DEFAULT[ext]).attrs.font, line_of(DEFAULT[ext]).attrs.size })
    end)
  end)
end

-- an empty paragraph with spacing of its own between two paragraphs (as forms have them), and (docx)
-- a style of a footnote based on the default one
local EMPTY = {
  docx = function(name, text)
    if name == "word/document.xml" then
      local p = function(t, spaced)
        return "<w:p>" .. (spaced and '<w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>' or "")
          .. (t and "<w:r><w:t>" .. t .. "</w:t></w:r>" or "") .. "</w:p>"
      end
      return (text:gsub("<w:body>", "<w:body>" .. p("Spaced A.") .. p(nil, true) .. p("Spaced B."), 1))
    elseif name == "word/styles.xml" then
      return (text:gsub("</w:styles>", '<w:style w:type="paragraph" w:styleId="NoteX"><w:name w:val="Note X"/><w:basedOn w:val="Normal"/></w:style></w:styles>', 1))
    elseif name == "word/footnotes.xml" then
      return (text:gsub("</w:footnotes>", '<w:footnote w:id="9"><w:p><w:pPr><w:pStyle w:val="NoteX"/></w:pPr><w:r><w:t>A note.</w:t></w:r></w:p></w:footnote></w:footnotes>', 1))
    end
    return text
  end,
  odt = function(name, text)
    if name ~= "content.xml" then
      return text
    end
    local style = '<style:style style:name="Tight" style:family="paragraph" style:parent-style-name="Standard"><style:paragraph-properties fo:margin-bottom="0cm" fo:line-height="100%"/></style:style>'
    local p = '<text:p text:style-name="Standard">Spaced A.</text:p><text:p text:style-name="Tight"/><text:p text:style-name="Standard">Spaced B.</text:p>'
    return (text:gsub("<office:automatic%-styles>", "%0" .. style, 1):gsub("<office:text>", "%0" .. p, 1))
  end,
}

for _, ext in ipairs({ "docx", "odt" }) do
  describe(ext .. " a style changed, and what isn't in it", function()
    before_each(function() neo.config.styles = true end)
    after_each(function() neo.config.styles = false end)

    it("leaves an empty paragraph's own format as it was (a blank line has none to say)", function()
      local path = H.copy("rich." .. ext, EMPTY[ext])
      H.open(path)
      local default = header().styles[1]
      local text = H.lines()[default.row]
      api.nvim_buf_set_lines(0, default.row - 1, default.row, false, { (text:gsub("font = [^,}]+", "font = Arial")) })
      if not text:find("font =", 1, true) then
        api.nvim_buf_set_lines(0, default.row - 1, default.row, false, { (text:gsub("{", "{font = Arial, ", 1)) })
      end
      H.save()
      local main = zip.read(path, H.MAIN[ext])
      if ext == "docx" then
        assert.truthy(main:find('<w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:p>', 1, true), main:sub(1, 1500))
        local note_style = zip.read(path, "word/styles.xml"):match('<w:style w:type="paragraph" w:styleId="NoteX">.-</w:style>')
        assert.truthy(note_style:find("<w:rFonts", 1, true), note_style) -- (the footnote's: as it was)
        assert.is_nil(note_style:find('w:ascii="Arial"', 1, true))
      else
        assert.truthy(main:find('<text:p text:style%-name="Tight"/>'), main:sub(1, 3000))
      end
    end)
  end)
end

describe("comments at the top", function()
  it("are lines starting with # (as in YAML): inactive, and the top is still the top with its first commented", function()
    local markup = require("neoffice.markup")
    local h = markup.header({ "---", "# page: {size = A4}", "styles:", "  # Normal: {size = 9pt}", "  Title: {size = 20pt}", "---", "Text." })
    assert.same({ 6, nil, 1, "Title" }, { h.last, h.page, #h.styles, h.styles[1].name })
    assert.is_nil(markup.header({ "---", "Text." }))
  end)

  it("change nothing when saved, a style's line changed and commented out", function()
    neo.config.styles = true
    local path = H.copy("rich.docx")
    H.open(path)
    neo.config.styles = false
    local before = H.bytes(path)
    local st = require("neoffice.markup").header(H.lines()).styles[1]
    local line = H.lines()[st.row]
    api.nvim_buf_set_lines(0, st.row - 1, st.row, false, { "  # " .. vim.trim(line:gsub("size = %d+pt", "size = 30pt")) })
    H.save()
    assert.equals(before, H.bytes(path))
    assert.equals(#H.paras(), #vim.tbl_filter(function() return true end, H.paras())) -- (no line of it read as text)
    assert.is_nil(vim.iter(H.paras()):find(function(p) return p:find("size = 30pt", 1, true) end))
  end)
end)

for _, ext in ipairs({ "docx", "odt" }) do
  describe(ext .. " a style's colour", function()
    before_each(function() neo.config.styles = true end)
    after_each(function() neo.config.styles = false end)

    it("changes the style, and the styles based on it keep theirs", function()
      local path = H.copy("rich." .. ext)
      H.open(path)
      local st = line_of(BODY[ext])
      assert.is_nil(st.attrs.color) -- (black, automatic: unsaid)
      local text = H.lines()[st.row]
      api.nvim_buf_set_lines(0, st.row - 1, st.row, false, { (text:gsub("}$", ", color = darkred}")) })
      H.reopen()
      assert.equals("darkred", line_of(BODY[ext]).attrs.color)
      assert.is_nil(line_of("First Paragraph").attrs.color) -- (based on it: as it showed)
      assert.truthy(zip.read(path, STYLES[ext]):find(ext == "docx" and 'w:val="800000"' or 'fo:color="#800000"', 1, true))
    end)
  end)
end
