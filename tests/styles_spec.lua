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

    it("shows the styles its paragraphs have, under those they are based on, each with what it says itself; and saves without changes byte for byte", function()
      local lines = H.lines()
      assert.equals("---", lines[1])
      assert.equals("styles:", lines[2])
      local body = line_of(BODY[ext])
      assert.truthy(body, table.concat(lines, "\n"))
      assert.same({ before = "9pt", after = "9pt" }, body.attrs) -- (its own: the rest from the one above it)
      local root = header().styles[1] -- (the default one: all it gives)
      assert.same({ "Cambria", "12pt" }, { root.attrs.font, root.attrs.size })
      local function depth(st) return #H.lines()[st.row]:match("^( *)") / 2 - 1 end
      assert.same({ 0, 1, 2 }, { depth(root), depth(body), depth(line_of("First Paragraph")) })
      -- (highlighted by neoffice, each line as tree-sitter shows YAML: read together, the indented
      -- ones would be wrong, a key with braces having none under it)
      api.nvim_buf_set_lines(0, 2, 2, false, { "  # a comment" })
      require("neoffice").decorate(0)
      local ns = api.nvim_get_namespaces().neoffice_decorations
      local function groups(row)
        local out = {}
        for _, m in ipairs(api.nvim_buf_get_extmarks(0, ns, { row - 1, 0 }, { row - 1, -1 }, { details = true })) do
          if m[4].hl_group then
            out[m[4].hl_group .. ":" .. H.lines()[row]:sub(m[3] + 1, m[4].end_col)] = true
          end
        end
        return out
      end
      local g = groups(line_of(BODY[ext]).row)
      local ok, q = pcall(vim.treesitter.query.get, "yaml", "highlights")
      if ok and q then -- (as an R Markdown file's YAML is shown: each line read alone, without its indentation)
        assert.truthy(g["@property.yaml:" .. BODY[ext]] and g["@string.yaml:before = 9pt"], vim.inspect(g))
        assert.truthy(groups(3)["@comment.yaml:# a comment"], vim.inspect(groups(3)))
      else -- (without a YAML parser: the names, comments, the rest in gray)
        local gray = vim.iter(vim.tbl_keys(g)):any(function(k) return k:match("^NeofficeMark::%s*{before = 9pt, after = 9pt}$") ~= nil end)
        assert.truthy(g["NeofficeStyleName:" .. BODY[ext]] and gray, vim.inspect(g))
        assert.truthy(groups(3)["NeofficeHeaderComment:  # a comment"], vim.inspect(groups(3)))
      end
      api.nvim_buf_set_lines(0, 2, 3, false, {})
      assert.truthy(line_of("Heading 1").attrs.bold)
      assert.truthy(H.paras()[1]) -- (the paragraphs after it, as they were)
      local before = H.bytes(path)
      vim.bo.modified = true
      H.save()
      assert.equals(before, H.bytes(path))
    end)

    it("changes a style, and the styles based on it follow where they say nothing themselves", function()
      local st = line_of(BODY[ext])
      local text = H.lines()[st.row]
      api.nvim_buf_set_lines(0, st.row - 1, st.row, false, { (text:gsub("}$", ", size = 14pt, align = justify}")) })
      H.reopen()
      assert.same({ size = "14pt", align = "justify", before = "9pt", after = "9pt" }, line_of(BODY[ext]).attrs)
      assert.is_nil(line_of("First Paragraph").attrs.size) -- (based on it: says nothing of its own)
      local doc = H.doc()
      local first = doc.fmt.style_id(doc.ctx, "First Paragraph")
      assert.same({ "14pt", "justify" }, { doc.fmt.style_view(doc.ctx, first).size, doc.fmt.style_view(doc.ctx, first).align }) -- (it follows)
      local styles = zip.read(path, STYLES[ext])
      assert.truthy(styles:find(ext == "docx" and 'w:sz w:val="28"' or 'fo:font%-size="14pt"'), styles:sub(1, 400))
      assert.is_true(require("neoffice.xml").check(styles))
      if H.has_libreoffice then
        local ok, out = H.libreoffice_reads(path, "Commented text")
        assert.is_true(ok, out)
      end
    end)

    it("gives a style what the one above it gives when a value is taken off its line", function()
      local st = line_of(BODY[ext])
      api.nvim_buf_set_lines(0, st.row - 1, st.row, false, { (H.lines()[st.row]:gsub("before = 9pt, ", "")) })
      H.reopen()
      assert.same({ after = "9pt" }, line_of(BODY[ext]).attrs)
      local doc = H.doc()
      local id = doc.fmt.style_id(doc.ctx, BODY[ext])
      local root = header().styles[1]
      assert.is_true(require("neoffice.paras").same_style_value("before", root.attrs.before, doc.fmt.style_view(doc.ctx, id).before)) -- (the default style's)
      local styles = zip.read(path, STYLES[ext])
      local own = ext == "docx" and styles:match('<w:style [^>]*w:styleId="BodyText".-</w:style>') or styles:match('<style:style style:name="Text_20_body".-</style:style>')
      assert.is_nil(own:find(ext == "docx" and "w:before=" or "fo:margin%-top"), own) -- (none of its own: not a 0 said)
      assert.is_true(require("neoffice.xml").check(styles))
    end)

    it("changes the styles under the default one where they say nothing, when it changes", function()
      local root = header().styles[1]
      api.nvim_buf_set_lines(0, root.row - 1, root.row, false, { (H.lines()[root.row]:gsub("font = [^,}]+", "font = Arial")) })
      H.reopen()
      assert.equals("Arial", header().styles[1].attrs.font)
      local doc = H.doc()
      local id = doc.fmt.style_id(doc.ctx, BODY[ext])
      assert.equals("Arial", doc.fmt.style_view(doc.ctx, id).font) -- (it says no font: the default style's)
      assert.is_nil(line_of(BODY[ext]).attrs.font)
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
    text = H.default_styles(name, text)
    if name ~= "word/document.xml" then
      return text
    end
    local p = '<w:p><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="22"/></w:rPr>'
      .. "<w:t>Sized text.</w:t></w:r></w:p>"
    return (text:gsub("<w:body>", "<w:body>" .. p .. p:gsub("Sized text", "More of it"), 1))
  end,
  odt = function(name, text)
    text = H.default_styles(name, text)
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
      assert.equals("{font = Times New Roman, size = 11pt}+", H.lines()[r])
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
        -- (the footnote's: as it was, nothing of its own; it follows the default style, as in Word)
        assert.equals('<w:style w:type="paragraph" w:styleId="NoteX"><w:name w:val="Note X"/><w:basedOn w:val="Normal"/></w:style>', note_style)
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

describe("a colour with # at the top", function()
  it("is written in quotes (\" #\" starts a comment in YAML), read either way, and quoted by formatting", function()
    local paras = require("neoffice.paras")
    assert.equals('font = Calibri, color = "#4f81bd", bold', paras.render_style({ font = "Calibri", color = "#4f81bd", bold = true }))
    assert.same({ color = "#4f81bd", bold = true }, paras.parse_style('color = "#4f81bd", bold'))
    assert.same({ color = "#4f81bd", bold = true }, paras.parse_style("color = #4f81bd, bold"))
    neo.config.styles = true
    H.open(H.copy("rich.docx"))
    neo.config.styles = false
    local st = require("neoffice.markup").header(H.lines()).styles[1]
    local line = H.lines()[st.row]
    api.nvim_buf_set_lines(0, st.row - 1, st.row, false, { (line:gsub("}$", ", color = #c00000}")) })
    local edits = require("neoffice.format").edits(api.nvim_get_current_buf())
    vim.lsp.util.apply_text_edits(edits, api.nvim_get_current_buf(), "utf-8")
    assert.truthy(H.lines()[st.row]:find('color = "#c00000"', 1, true), H.lines()[st.row])
  end)
end)

describe(":Neoffice styles reset", function()
  before_each(function() neo.config.styles = true end)
  after_each(function() neo.config.styles = false end)
  local MAIN = { docx = "word/document.xml", odt = "content.xml" }
  local function top()
    local lines = H.lines()
    return vim.list_slice(lines, 1, header().last)
  end
  local function restyle_all()
    for k, l in ipairs(H.lines()) do
      if k > 2 and l:match("^  %S") and l:match("{") then
        api.nvim_buf_set_lines(0, k - 1, k, false, { (l:gsub("{.*}", "{font = Courier New, size = 20pt, align = center, italic, after = 30pt}")) })
      end
      if l == "---" and k > 1 then
        break
      end
    end
  end

  for _, ext in ipairs({ "docx", "odt" }) do
    it("makes the styles a new document's, after they were changed and saved (" .. ext .. ")", function()
      local path = H.copy("rich." .. ext) -- (made from the same empty document as neoffice's)
      H.open(path)
      local original = top()
      restyle_all()
      H.save()
      H.reopen()
      assert.are_not.same(original, top())
      local text = zip.read(path, MAIN[ext])
      vim.cmd("Neoffice styles reset")
      assert.same(original, top())
      assert.is_true(vim.bo.modified)
      H.save()
      assert.equals(text, zip.read(path, MAIN[ext]), "the text written again")
      H.open(path)
      assert.same(original, top())
    end)
  end

  it("leaves a style the new document hasn't nothing of its own: it follows its base (docx)", function()
    local path = H.copy("rich.docx", function(name, text)
      if name == "word/document.xml" then -- (a paragraph in a style of its own, based on Body Text)
        return (text:gsub('<w:pStyle w:val="BodyText" />', '<w:pStyle w:val="MyOwn" />', 1))
      end
      return text
    end)
    local styles = zip.read(path, "word/styles.xml")
    styles = styles:gsub("</w:styles>", '<w:style w:type="paragraph" w:customStyle="1" w:styleId="MyOwn"><w:name w:val="My Own" /><w:basedOn w:val="BodyText" />'
      .. '<w:pPr><w:jc w:val="center" /></w:pPr><w:rPr><w:rFonts w:ascii="Courier New" w:hAnsi="Courier New" /><w:sz w:val="40" /></w:rPr></w:style></w:styles>')
    zip.write(path, { ["word/styles.xml"] = styles })
    H.open(path)
    assert.truthy(line_of("My Own").text:find("Courier New", 1, true), line_of("My Own").text)
    vim.cmd("Neoffice styles reset")
    assert.equals("", line_of("My Own").text) -- (nothing of its own: as Body Text, what it is based on)
    H.save()
    local written = zip.read(path, "word/styles.xml"):match('<w:style [^>]*w:styleId="MyOwn".-</w:style>')
    assert.truthy(written:find('w:basedOn w:val="BodyText"', 1, true), written)
    assert.is_nil(written:find("Courier", 1, true), written)
  end)

  it("keeps the headings numbered (docx)", function()
    local path = H.copy("rich.docx")
    H.open(path)
    vim.cmd("Neoffice numbers on")
    H.save()
    H.reopen()
    vim.cmd("Neoffice styles reset")
    H.save()
    local heading = zip.read(path, "word/styles.xml"):match('<w:style [^>]*w:styleId="Heading1".-</w:style>')
    assert.truthy(heading:find("<w:numPr>", 1, true), heading)
  end)

  it("writes a valid file where the document doesn't declare the namespaces the new styles use (odt)", function()
    local path = H.copy("rich.odt")
    local styles = zip.read(path, "styles.xml")
    styles = styles:gsub(' xmlns:loext="[^"]*"', ""):gsub(' loext:[%w%-]+="[^"]*"', ""):gsub("<loext:[^>]*/>", "")
    zip.write(path, { ["styles.xml"] = styles })
    H.open(path)
    vim.cmd("Neoffice styles reset")
    local said, notify = {}, vim.notify
    vim.notify = function(msg) said[#said + 1] = msg end
    H.save()
    vim.notify = notify
    assert.is_nil(table.concat(said, "\n"):find("failed", 1, true), table.concat(said, "\n"))
    assert.truthy(zip.read(path, "styles.xml"):find('xmlns:loext="', 1, true))
  end)
end)

-- a style that says its first line isn't indented (0), based on one that indents it, and a paragraph
-- that says so of its own: a 0 said is said, not "nothing said" (else the base's indent shows, and
-- is written when the style's line is edited)
for _, ext in ipairs({ "docx", "odt" }) do
  describe(ext .. " an indent of 0 said by a style or a paragraph", function()
    before_each(function() neo.config.styles = true end)
    after_each(function() neo.config.styles = false end)
    local function indented(path)
      if ext == "docx" then
        local styles = zip.read(path, "word/styles.xml")
        styles = styles:gsub('(<w:style [^>]*w:styleId="BodyText".-<w:pPr>)', '%1<w:ind w:firstLine="567"/>', 1)
        styles = styles:gsub('(<w:style [^>]*w:styleId="FirstParagraph".-)</w:style>', '%1<w:pPr><w:ind w:firstLine="0"/></w:pPr></w:style>', 1)
        zip.write(path, { ["word/styles.xml"] = styles })
      else
        local styles = zip.read(path, "styles.xml")
        styles = styles:gsub('(<style:style style:name="Text_20_body".-<style:paragraph%-properties)', '%1 fo:text-indent="1cm"', 1)
        styles = styles:gsub('(<style:style style:name="First_20_Paragraph"[^>]-)/>', '%1><style:paragraph-properties fo:text-indent="0cm"/></style:style>', 1)
        zip.write(path, { ["styles.xml"] = styles })
      end
    end
    local function style_line(name)
      for _, l in ipairs(H.lines()) do
        if l:match("^%s+" .. vim.pesc(name) .. ":") then
          return l
        end
      end
    end

    it("shows the style without the base's indent, and keeps its 0 when its line is edited", function()
      local path = H.copy("rich." .. ext)
      indented(path)
      H.open(path)
      assert.truthy(style_line(BODY[ext]):find("first = 1cm", 1, true), style_line(BODY[ext]))
      local line = style_line("First Paragraph")
      assert.truthy(line:find("first = 0cm", 1, true), line) -- (its own 0, over the base's 1cm)
      local doc = H.doc()
      assert.is_nil(doc.fmt.style_view(doc.ctx, doc.fmt.style_id(doc.ctx, "First Paragraph")).first) -- (as it comes out: none)
      local r = vim.fn.index(H.lines(), line) + 1
      api.nvim_buf_set_lines(0, r - 1, r, false, { (line:gsub("}%s*$", ", size = 14pt}")) })
      H.reopen()
      line = style_line("First Paragraph")
      assert.truthy(line:find("size = 14pt", 1, true), line)
      assert.truthy(line:find("first = 0cm", 1, true), line)
      doc = H.doc()
      assert.is_nil(doc.fmt.style_view(doc.ctx, doc.fmt.style_id(doc.ctx, "First Paragraph")).first)
    end)
  end)
end

-- a paragraph of a style based on the default one, whose text says the default style's font (as
-- Word writes text often): it shows none of its own, so when the default style's font changes it
-- follows, as the style it is in does (its paragraph written again, not only its style's)
describe("docx the default style's font changed", function()
  before_each(function() neo.config.styles = true end)
  after_each(function() neo.config.styles = false end)
  it("changes the text of the styles under it that showed none of their own", function()
    local path = H.copy("rich.docx", function(name, text)
      if name ~= "word/document.xml" then
        return text
      end
      return (text:gsub('(<w:commentRangeStart w:id="0" />)<w:r><w:t xml:space="preserve">Commented', '%1<w:r><w:rPr><w:rFonts w:ascii="Cambria" w:hAnsi="Cambria"/></w:rPr><w:t xml:space="preserve">Commented', 1))
    end)
    H.open(path)
    local function above(text)
      for r, l in ipairs(H.lines()) do
        if l:find(text, 1, true) then
          return H.lines()[r - 1]
        end
      end
    end
    assert.equals("", above("Commented text")) -- (the default style's font: nothing to say)
    local root = header().styles[1]
    assert.equals("Cambria", root.attrs.font)
    api.nvim_buf_set_lines(0, root.row - 1, root.row, false, { (H.lines()[root.row]:gsub("font = Cambria", "font = Arial")) })
    H.reopen()
    assert.equals("", above("Commented text")) -- (in Arial now, as its style: still nothing to say)
    local p = vim.iter(H.paragraph_xml(path)):find(function(x) return x:find("Commented", 1, true) end)
    assert.is_nil(p:find("Cambria", 1, true), p)
  end)
end)
