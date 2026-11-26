-- Fonts and sizes: a paragraph's in its format ("<=> {font = Times New Roman, size = 11pt}"), said
-- only where they aren't the document's; a few words in another in a span ("[words]{font = Arial}").
local H = dofile("tests/helpers.lua")
local zip = require("neoffice.zip")
local markup = require("neoffice.markup")
local api = vim.api

-- two paragraphs whose font and size are set on their text (as Word and LibreOffice often do), first
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
    local p = '<text:p><text:span text:style-name="Sized">Sized text.</text:span></text:p>'
    return (text:gsub("<office:automatic%-styles>", "%0" .. style, 1):gsub("<office:text>", "%0" .. p .. p:gsub("Sized text", "More of it"), 1))
  end,
}
local LINE = "{font = Times New Roman, size = 11pt}"

describe("fonts in the markup", function()
  it("read a span of text in another font, and write it so", function()
    local line = markup.render({ { text = "the " }, { text = "Latin name", font = "Arial", size = "9pt" }, { text = " of it" } })
    assert.equals("the [Latin name]{font = Arial, size = 9pt} of it", line)
    local _, chars = markup.parse(line)
    assert.same({ "Arial", "9pt" }, { chars[5].font, chars[5].size })
    assert.is_nil(chars[1].font)
    -- (with marks inside, and a name that needs quotes)
    line = markup.render({ { text = "x", b = true, font = "Times New Roman;serif" } })
    assert.equals('[**x**]{font = "Times New Roman;serif"}', line)
    _, chars = markup.parse(line)
    assert.same({ "Times New Roman;serif", true }, { chars[1].font, chars[1].b })
  end)

  it("write text that only looks like a span escaped, and say what is wrong in one", function()
    local line = markup.render({ { text = "a [b]{c} d" } })
    assert.equals("a \\[b]{c} d", line)
    assert.is_true(markup.same(line, { { text = "a [b]{c} d" } }))
    local _, chars = markup.parse("[x]{colour = red}")
    assert.truthy(chars[1].span_bad:find("font, size", 1, true))
    assert.same({ font = "Times New Roman", size = "11pt" }, markup.span_attrs("font = Times New Roman, size = 11pt"))
  end)
end)

for _, ext in ipairs({ "docx", "odt" }) do
  describe(ext .. " fonts", function()
    local path
    local function main() return zip.read(path, H.MAIN[ext]) end
    local function at(text) return vim.fn.index(H.paras(), text) + 1 end
    before_each(function()
      path = H.copy("rich." .. ext, SIZED[ext])
      H.open(path)
    end)

    it("shows the font and size of text that has the document's no more than once, and saves it unchanged", function()
      assert.equals(LINE, H.lines()[H.row(1) - 1])
      assert.same({ "Sized text.", "More of it." }, vim.list_slice(H.paras(), 1, 2))
      for _, p in ipairs(vim.list_slice(H.paras(), 3)) do -- (the others: in the document's font, unsaid)
        assert.is_nil(p:find("font =", 1, true), p)
      end
      local before = H.bytes(path)
      vim.bo.modified = true
      H.save()
      assert.equals(before, H.bytes(path))
    end)

    it("changes the font and size of the paragraphs under a format's line", function()
      api.nvim_buf_set_lines(0, H.row(1) - 2, H.row(1) - 1, false, { "{font = Arial, size = 14pt}" })
      H.reopen()
      assert.equals("{font = Arial, size = 14pt}", H.lines()[H.row(1) - 1])
      local xml = main()
      assert.is_true(require("neoffice.xml").check(xml)) -- (a part LibreOffice can read)
      if H.has_libreoffice then
        local ok, text = H.libreoffice_reads(path, "More of it.")
        assert.is_true(ok, text)
      end
      if ext == "docx" then
        local _, n = xml:gsub('<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="28"/><w:szCs w:val="28"/>', "")
        assert.equals(2, n, xml:sub(1, 1500))
      else
        assert.truthy(xml:find('<style:font%-face style:name="Arial" svg:font%-family="Arial"/>'), xml:sub(1, 1500))
        assert.truthy(xml:find('style:font%-name="Arial"[^>]*fo:font%-size="14pt"') or xml:find('fo:font%-size="14pt"[^>]*style:font%-name="Arial"'), xml:sub(1, 3000))
      end
    end)

    it("writes a few words in another font, and takes them back to the paragraph's", function()
      H.replace(at("More of it."), "of it", "[of it]{font = Courier New}")
      H.reopen()
      assert.equals("More [of it]{font = Courier New}.", H.paras()[2])
      assert.truthy(main():find("Courier New", 1, true))
      H.replace(2, "[of it]{font = Courier New}", "of it")
      H.reopen()
      assert.equals("More of it.", H.paras()[2])
      assert.equals(LINE, H.lines()[H.row(1) - 1])
    end)

    it("takes text back to the document's font when its format's line says none", function()
      api.nvim_buf_set_lines(0, H.row(1) - 2, H.row(1) - 1, false, {})
      H.reopen()
      assert.same({ "Sized text.", "More of it." }, vim.list_slice(H.paras(), 1, 2))
      assert.is_nil(H.lines()[1]:find("font =", 1, true)) -- (the document's: unsaid)
    end)

    it("doesn't write a span's attribute it doesn't know: the save says so and changes nothing", function()
      local before = H.bytes(path)
      H.replace(at("More of it."), "of it", "[of it]{colour = red}")
      local errors = {}
      local notify = vim.notify
      vim.notify = function(msg) errors[#errors + 1] = msg end
      H.save()
      vim.notify = notify
      assert.equals(before, H.bytes(path))
      assert.truthy(table.concat(errors, "\n"):find("font, size", 1, true))
    end)
  end)
end

describe("the language server, on the document's fonts", function()
  it("offers the fonts the document has first", function()
    H.open(H.copy("rich.docx", SIZED.docx))
    local r = H.row(3)
    vim.api.nvim_buf_set_lines(0, r - 1, r - 1, false, { "Some [words]{font = " })
    local buf = vim.api.nvim_get_current_buf()
    assert(vim.wait(2000, function() return #vim.lsp.get_clients({ bufnr = buf, name = "neoffice_ls" }) > 0 end))
    local params = { textDocument = { uri = vim.uri_from_bufnr(buf) }, position = { line = r - 1, character = #"Some [words]{font = " } }
    local items = {}
    for _, res in pairs(vim.lsp.buf_request_sync(buf, "textDocument/completion", params, 2000) or {}) do
      items = res.result and (res.result.items or res.result) or items
    end
    table.sort(items, function(x, y) return (x.sortText or x.label) < (y.sortText or y.label) end)
    local first = vim.tbl_map(function(i) return i.label end, vim.list_slice(items, 1, 2))
    assert.truthy(vim.tbl_contains(first, "Times New Roman"), vim.inspect(first))
    -- (its text's fonts, even said nowhere else: on the line of its format, being written)
    r = H.row(1) - 1
    vim.api.nvim_buf_set_lines(0, r - 1, r, false, { "<<- {font = " })
    params.position = { line = r - 1, character = #"<<- {font = " }
    for _, res in pairs(vim.lsp.buf_request_sync(buf, "textDocument/completion", params, 2000) or {}) do
      items = res.result and (res.result.items or res.result) or items
    end
    table.sort(items, function(x, y) return (x.sortText or x.label) < (y.sortText or y.label) end)
    first = vim.tbl_map(function(i) return i.label end, vim.list_slice(items, 1, 2))
    assert.truthy(vim.tbl_contains(first, "Times New Roman"), vim.inspect(first))
  end)
end)

describe("the language server, on fonts", function()
  local function complete(row, col)
    local buf = api.nvim_get_current_buf()
    assert(vim.wait(2000, function() return #vim.lsp.get_clients({ bufnr = buf, name = "neoffice_ls" }) > 0 end))
    local params = { textDocument = { uri = vim.uri_from_bufnr(buf) }, position = { line = row - 1, character = col } }
    local result
    for _, r in pairs(vim.lsp.buf_request_sync(buf, "textDocument/completion", params, 2000) or {}) do
      result = r.result
    end
    return result and (result.items or result) or {}
  end
  it("completes a font's name with those installed, and sizes", function()
    local fonts = require("neoffice.lsp").fonts()
    if #fonts == 0 then
      return pending("no fonts known (fontconfig)")
    end
    H.open(H.copy("rich.docx"))
    local family = fonts[1]
    H.set(1, "<=> {font = " .. family:sub(1, 3))
    local items = complete(H.row(1), #H.lines()[H.row(1)])
    assert.truthy(vim.iter(items):find(function(i) return i.label == markup.attr_value(family) end), vim.inspect(vim.tbl_map(function(i) return i.label end, items)))
    -- (a word from the middle of a name finds it too)
    local multi = vim.iter(fonts):find(function(f) return f:find(" ", 1, true) end)
    if multi then
      local word = multi:match("%s(%S+)")
      H.set(1, "<=> {font = " .. word:lower())
      items = complete(H.row(1), #H.lines()[H.row(1)])
      assert.truthy(vim.iter(items):find(function(i) return i.label == markup.attr_value(multi) end), word)
    end
    for _, mark in ipairs({ "<=>", "<|>", "->>", "<<-" }) do -- (after any mark)
      H.set(1, mark .. " {size = ")
      items = complete(H.row(1), #H.lines()[H.row(1)])
      assert.truthy(vim.iter(items):find(function(i) return i.label == "11pt" end), mark)
    end
    H.set(1, "Some [words]{size = ")
    items = complete(H.row(1), #H.lines()[H.row(1)])
    assert.truthy(vim.iter(items):find(function(i) return i.label == "11pt" end))
  end)
end)
