-- A plain paragraph's style in its format ({style = First Paragraph}, see paras.lua): shown where it
-- isn't the one plain paragraphs in its place have (first ones: first in the document, after a
-- heading, title or quote; and others), written where it is said, the usual one where it isn't.
local H = dofile("tests/helpers.lua")
local zip = require("neoffice.zip")
local api = vim.api

local FIRST = { docx = "FirstParagraph", odt = "First_20_Paragraph" }
local BODY = { docx = "BodyText", odt = "Text_20_body" }

-- the style (id) of the paragraph whose text has `text`, as the document now has it
local function style_of(text)
  local doc = H.doc()
  for k, t in ipairs(doc.texts) do
    if t:find(text, 1, true) then
      return doc.fmt.style_of(doc.ctx, doc.paras[k])
    end
  end
end
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
    if l:find(text, 1, true) then
      return r
    end
  end
end

for _, ext in ipairs({ "docx", "odt" }) do
  describe(ext .. " a paragraph's style", function()
    it("isn't shown where it is the usual one in its place (pandoc's First Paragraph and Body Text)", function()
      H.open(H.copy("rich." .. ext))
      assert.same({}, vim.tbl_filter(function(l) return l:find("style =", 1, true) ~= nil end, H.lines()))
      assert.equals(FIRST[ext], H.doc().first_style)
      assert.equals(BODY[ext], H.doc().plain_style)
    end)

    it("is written where a paragraph says it, shown so, and taken back when that is deleted", function()
      local path = H.copy("rich." .. ext)
      H.open(path)
      local r = row_of("Commented text")
      local line = H.lines()[r]
      api.nvim_buf_set_lines(0, r - 1, r, false, { "{style = First Paragraph} " .. line })
      H.reopen()
      assert.equals(FIRST[ext], style_of("Commented text"))
      assert.equals("{style = First Paragraph} " .. line, H.lines()[row_of("Commented text")])
      assert.equals(BODY[ext], style_of("Text with inserted")) -- (the one after it as it was)
      if ext == "docx" then -- (its style first in its properties, as the schema wants)
        local p = vim.iter(H.paragraph_xml(path)):find(function(x) return x:find("Commented text", 1, true) end)
        assert.truthy(p:find('^<w:p[^>]*><w:pPr><w:pStyle w:val="FirstParagraph" ?/>'), p)
      end
      -- (said again unchanged: saved as it is)
      local before = H.bytes(path)
      vim.bo.modified = true
      H.save()
      assert.equals(before, H.bytes(path))
      -- (taken away: the usual one there again)
      r = row_of("Commented text")
      api.nvim_buf_set_lines(0, r - 1, r, false, { line })
      H.reopen()
      assert.equals(BODY[ext], style_of("Commented text"))
      assert.equals(line, H.lines()[row_of("Commented text")])
    end)

    it("is given by a format's line to the paragraphs after it, until one says none", function()
      local path = H.copy("rich." .. ext)
      H.open(path)
      local r = row_of("Commented text")
      api.nvim_buf_set_lines(0, r - 1, r - 1, false, { "{style = first paragraph}" }) -- (whatever its case)
      r = row_of("An image")
      api.nvim_buf_set_lines(0, r - 1, r - 1, false, { "{style = none}" })
      H.reopen()
      assert.equals(FIRST[ext], style_of("Commented text"))
      assert.equals(FIRST[ext], style_of("Text with inserted"))
      assert.equals(BODY[ext], style_of("An image"))
      assert.equals(BODY[ext], style_of("Last paragraph"))
      assert.equals(FIRST[ext], style_of("This is ")) -- (the first, after the heading: as it was)
      -- (shown as said: the two paragraphs under a line of it, the others not)
      local shown = vim.tbl_filter(function(l) return l:find("style =", 1, true) ~= nil end, H.lines())
      assert.truthy(#shown >= 1, table.concat(H.lines(), "\n"))
      local before = H.bytes(path)
      vim.bo.modified = true
      H.save()
      assert.equals(before, H.bytes(path))
    end)

    it("isn't written when the document hasn't it: the save says so and changes nothing", function()
      local path = H.copy("rich." .. ext)
      H.open(path)
      local before = H.bytes(path)
      local r = row_of("Commented text")
      api.nvim_buf_set_lines(0, r - 1, r, false, { "{style = No Such Style} " .. H.lines()[r] })
      local said = said_by(H.save)
      assert.truthy(said:find("not a paragraph style of this document: No Such Style", 1, true), said)
      assert.equals(before, H.bytes(path))
    end)

    it("isn't given to a paragraph in a box: the save says so and changes nothing", function()
      local path = H.copy("rich." .. ext)
      H.open(path)
      local r = row_of("Last paragraph")
      api.nvim_buf_set_lines(0, r, r, false, { "", "::: {width = 5cm}", "{style = First Paragraph} In a box.", ":::" })
      local before = H.bytes(path)
      local said = said_by(H.save)
      assert.truthy(said:find("not one in a table cell or a box", 1, true), said)
      assert.equals(before, H.bytes(path))
    end)

    it("of a document all of one style shows nowhere; a paragraph given another keeps it", function()
      local path = H.copy("rich." .. ext, H.default_styles)
      H.open(path)
      assert.same({}, vim.tbl_filter(function(l) return l:find("style =", 1, true) ~= nil end, H.lines()))
      local plain = style_of("Commented text")
      local r = row_of("Commented text")
      api.nvim_buf_set_lines(0, r - 1, r, false, { "{style = First Paragraph} " .. H.lines()[r] })
      H.reopen()
      assert.equals(FIRST[ext], style_of("Commented text"))
      local shown = vim.tbl_filter(function(l) return l:find("style =", 1, true) ~= nil end, H.lines())
      assert.equals(1, #shown, table.concat(H.lines(), "\n"))
      assert.truthy(shown[1]:find("^{style = First Paragraph} Commented text"), shown[1])
      -- (a paragraph typed after it: the usual one, not its)
      r = row_of("Commented text")
      api.nvim_buf_set_lines(0, r, r, false, { "", "Typed after it." })
      H.reopen()
      assert.equals(plain, style_of("Typed after it."))
      -- (the only first paragraph given it: that is the usual one there now, not shown)
      r = row_of("This is ")
      api.nvim_buf_set_lines(0, r - 1, r, false, { "{style = First Paragraph} " .. H.lines()[r] })
      H.reopen()
      assert.equals(FIRST[ext], style_of("This is "))
      assert.is_nil(H.lines()[row_of("This is ")]:find("style =", 1, true))
    end)

    it("of a paragraph typed where an empty caption's was, after a table, is a plain one", function()
      local path = H.copy("rich." .. ext, function(name, text)
        if name == "word/document.xml" then
          return (text:gsub("</w:tbl>", '</w:tbl><w:p><w:pPr><w:pStyle w:val="TableCaption"/></w:pPr></w:p>', 1))
        elseif name == "content.xml" then
          return (text:gsub("</table:table>", '</table:table><text:p text:style-name="Table"/>', 1))
        end
        return text
      end)
      H.open(path)
      local r = row_of("Last paragraph") - 1 -- (the blank line before it: the empty caption's)
      api.nvim_buf_set_lines(0, r - 1, r, false, { "Typed after the table.", "" })
      H.reopen()
      assert.equals(BODY[ext], style_of("Typed after the table."))
      assert.equals(BODY[ext], style_of("Last paragraph"))
    end)

    it("of a style the document doesn't define (a file broken so) is kept when its paragraph is edited", function()
      local path = H.copy("rich." .. ext, function(name, text)
        if name == "word/document.xml" then
          return (text:gsub("<w:sectPr", '<w:p><w:pPr><w:pStyle w:val="Ghost"/></w:pPr><w:r><w:t>Haunted text.</w:t></w:r></w:p>%0', 1))
        elseif name == "content.xml" then
          return (text:gsub("</office:text>", '<text:p text:style-name="Ghost">Haunted text.</text:p>%0', 1))
        end
        return text
      end)
      H.open(path)
      local r = row_of("Haunted text.")
      assert.truthy(H.lines()[r]:find("style = Ghost", 1, true), H.lines()[r])
      H.replace(vim.fn.index(H.paras(), H.lines()[r]) + 1, "Haunted", "Still haunted")
      local said = said_by(H.save)
      assert.is_nil(said:find("failed", 1, true), said)
      H.reopen()
      assert.equals("Ghost", style_of("Still haunted text."))
    end)

    it("is completed after style =, from the document's", function()
      local path = H.copy("rich." .. ext)
      H.open(path)
      local buf = api.nvim_get_current_buf()
      assert(vim.wait(2000, function() return #vim.lsp.get_clients({ bufnr = buf, name = "neoffice_ls" }) > 0 end))
      local r = row_of("Commented text")
      api.nvim_buf_set_lines(0, r - 1, r - 1, false, { "{style = fir" })
      local params = { textDocument = { uri = vim.uri_from_bufnr(buf) }, position = { line = r - 1, character = #"{style = fir" } }
      local labels = {}
      for _, res in pairs(vim.lsp.buf_request_sync(buf, "textDocument/completion", params, 2000) or {}) do
        for _, item in ipairs(res.result and (res.result.items or res.result) or {}) do
          labels[#labels + 1] = item.label
        end
      end
      assert.truthy(vim.tbl_contains(labels, "First Paragraph"), vim.inspect(labels))
      assert.is_false(vim.tbl_contains(labels, "Heading 1"), vim.inspect(labels))
    end)
  end)
end

describe("odt a paragraph's style, with formatting of its own shared with another", function()
  it("changes that paragraph's only, its own formatting kept", function()
    local path = H.copy("rich.odt", function(name, text)
      if name ~= "content.xml" then
        return text
      end
      local style = '<style:style style:name="Centred" style:family="paragraph" style:parent-style-name="Text_20_body"><style:paragraph-properties fo:text-align="center"/></style:style>'
      local p = '<text:p text:style-name="Centred">One centred.</text:p><text:p text:style-name="Centred">Two centred.</text:p>'
      return (text:gsub("<office:automatic%-styles>", "%0" .. style, 1):gsub("</office:text>", p .. "%0", 1))
    end)
    H.open(path)
    local r = row_of("One centred.")
    api.nvim_buf_set_lines(0, r - 1, r, false, { "{style = First Paragraph} One centred." })
    H.reopen()
    assert.equals("First_20_Paragraph", style_of("One centred."))
    assert.equals("Text_20_body", style_of("Two centred."))
    local main = zip.read(path, "content.xml")
    local name = main:match('<text:p text:style%-name="([^"]+)">One centred%.')
    local own = main:match('<style:style style:name="' .. vim.pesc(name) .. '".-</style:style>')
    assert.truthy(own:find('style:parent-style-name="First_20_Paragraph"', 1, true), own)
    assert.truthy(own:find('fo:text-align="center"', 1, true), own) -- (still centred)
    assert.truthy(main:find('<text:p text:style-name="Centred">Two centred.', 1, true))
    assert.truthy(zip.read(path, "content.xml"):find('style:name="Centred"', 1, true)) -- (the other's, as it was)
  end)
end)
