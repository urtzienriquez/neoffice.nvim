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
      api.nvim_buf_set_lines(0, 2, 3, false, { "  " .. BODY[ext] .. ": {colour = red}" })
      H.save()
      vim.notify = notify
      assert.equals(before, H.bytes(path))
      assert.truthy(table.concat(errors, "\n"):find("not a style's attribute: colour", 1, true))
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
