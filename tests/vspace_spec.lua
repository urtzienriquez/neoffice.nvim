-- Vertical space: \vspace{1cm}, a line of its own between paragraphs, is an empty paragraph of
-- exactly that height with no space before or after it; such a paragraph opens as \vspace again.
local H = dofile("tests/helpers.lua")
local zip = require("neoffice.zip")
local api = vim.api

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
local function shown()
  return vim.tbl_filter(function(l) return l:find("^\\vspace") ~= nil end, H.lines())
end
-- the empty paragraphs' heights (cm), in the document's order
local function heights(path, ext)
  local main, out = zip.read(path, H.MAIN[ext]), {}
  if ext == "docx" then
    for p in main:gmatch("<w:p>.-</w:p>") do
      local line = not p:find("<w:t", 1, true) and p:match('w:line="(%d+)" w:lineRule="exact"')
      if line then
        out[#out + 1] = ("%.2f"):format(tonumber(line) / 1440 * 2.54)
      end
    end
  else
    for name in main:gmatch('<text:p text:style%-name="([^"]+)"%s*/>') do
      local st = main:match('<style:style style:name="' .. vim.pesc(name) .. '".-</style:style>') or ""
      local line = st:match('fo:line%-height="([%d.]+)cm"')
      if line then
        out[#out + 1] = ("%.2f"):format(tonumber(line))
      end
    end
  end
  return out
end

for _, ext in ipairs({ "docx", "odt" }) do
  describe(ext .. " \\vspace", function()
    it("is written as an empty paragraph of that height, opens as it was, and saves unchanged", function()
      local path = H.copy("rich." .. ext)
      H.open(path)
      local r = row_of("An image")
      api.nvim_buf_set_lines(0, r - 1, r - 1, false, { "\\vspace{1cm}", "" })
      r = row_of("An image")
      api.nvim_buf_set_lines(0, r, r, false, { "", "\\vspace{ 12pt }" })
      H.reopen()
      assert.same({ "\\vspace{1cm}", "\\vspace{12pt}" }, shown())
      assert.same({ "1.00", "0.42" }, heights(path, ext))
      assert.equals("\\vspace{1cm}", H.lines()[row_of("An image") - 2])
      if ext == "docx" then -- (no space before or after it, the usual style: the paragraphs around as they were)
        local p = vim.iter(H.paragraph_xml(path)):find(function(x) return x:find('w:line="567"', 1, true) end)
        assert.truthy(p:find('<w:pStyle w:val="BodyText" ?/><w:spacing w:before="0" w:after="0" w:line="567" w:lineRule="exact" ?/>'), p)
      end
      local before = H.bytes(path)
      vim.bo.modified = true
      H.save()
      assert.equals(before, H.bytes(path))
    end)

    it("alone is 1cm, highlighted as a page break is", function()
      local path = H.copy("rich." .. ext)
      H.open(path)
      local r = row_of("An image")
      api.nvim_buf_set_lines(0, r - 1, r - 1, false, { "\\vspace", "" })
      require("neoffice").decorate(0)
      local groups = vim.tbl_map(function(m) return m[4].hl_group end,
        api.nvim_buf_get_extmarks(0, -1, { r - 1, 0 }, { r - 1, -1 }, { details = true }))
      assert.truthy(vim.tbl_contains(groups, "neofficeBreak"), vim.inspect(groups))
      H.reopen()
      assert.same({ "\\vspace{1cm}" }, shown())
      assert.same({ "1.00" }, heights(path, ext))
    end)

    it("is a paragraph of its own, without blank lines around it", function()
      local path = H.copy("rich." .. ext)
      H.open(path)
      local r = row_of("An image")
      api.nvim_buf_set_lines(0, r, r, false, { "\\vspace{5mm}", "Right after it." })
      H.reopen()
      assert.same({ "\\vspace{0.5cm}" }, shown())
      assert.truthy(row_of("Right after it."))
      assert.truthy(H.lines()[row_of("An image")]:find("^An image !%[%]"), H.lines()[row_of("An image")])
    end)

    it("changes height, and goes when its line is deleted", function()
      local path = H.copy("rich." .. ext)
      H.open(path)
      local r = row_of("An image")
      api.nvim_buf_set_lines(0, r - 1, r - 1, false, { "\\vspace{1cm}", "" })
      H.reopen()
      r = row_of("\\vspace")
      api.nvim_buf_set_lines(0, r - 1, r, false, { "\\vspace{2cm}" })
      H.reopen()
      assert.same({ "\\vspace{2cm}" }, shown())
      assert.same({ "2.00" }, heights(path, ext))
      local paras = #H.paras()
      r = row_of("\\vspace")
      api.nvim_buf_set_lines(0, r - 1, r + 1, false, {})
      H.reopen()
      assert.same({}, shown())
      assert.same({}, heights(path, ext))
      assert.equals(paras - 1, #H.paras())
    end)

    it("isn't written when it isn't a length: the save says so and changes nothing", function()
      local path = H.copy("rich." .. ext)
      H.open(path)
      local before = H.bytes(path)
      for _, bad in ipairs({ "\\vspace{lots}", "\\vspace{-1cm}" }) do
        local r = row_of("An image")
        api.nvim_buf_set_lines(0, r - 1, r - 1, false, { bad, "" })
        local said = said_by(H.save)
        assert.truthy(said:find("not a length above 0", 1, true), said)
        assert.equals(before, H.bytes(path))
        api.nvim_buf_set_lines(0, r - 1, r + 1, false, {})
      end
    end)

    it("as text is text: in a paragraph as it is, starting one escaped", function()
      local path = H.copy("rich." .. ext)
      H.open(path)
      local r = row_of("Last paragraph")
      api.nvim_buf_set_lines(0, r - 1, r, false, { "Some \\vspace{1cm} here.", "", "\\\\vspace{1cm}" })
      H.reopen()
      assert.same({}, shown())
      assert.same({}, heights(path, ext))
      assert.truthy(row_of("Some \\vspace{1cm} here."), table.concat(H.lines(), "\n"))
      assert.truthy(vim.tbl_contains(H.lines(), "\\\\vspace{1cm}"), table.concat(H.lines(), "\n"))
    end)

    it("isn't put in a box: the save says so and changes nothing", function()
      local path = H.copy("rich." .. ext)
      H.open(path)
      local before = H.bytes(path)
      local r = row_of("Last paragraph")
      api.nvim_buf_set_lines(0, r, r, false, { "", "::: {width = 5cm}", "In a box.", "", "\\vspace{1cm}", "", "Also.", ":::" })
      local said = said_by(H.save)
      assert.truthy(said:find("not in a table cell or a box", 1, true), said)
      assert.equals(before, H.bytes(path))
    end)

    it("between a heading and a paragraph leaves that paragraph a first one (its style: First Paragraph)", function()
      local path = H.copy("rich." .. ext)
      H.open(path)
      local n = api.nvim_buf_line_count(0)
      api.nvim_buf_set_lines(0, n, n, false, { "", "# Methods", "", "\\vspace{1cm}", "", "New text after it." })
      H.reopen()
      local doc = H.doc()
      local k = vim.fn.index(doc.texts, "New text after it.") + 1
      assert.equals(ext == "docx" and "FirstParagraph" or "First_20_Paragraph", doc.fmt.style_of(doc.ctx, doc.paras[k]))
      assert.equals("New text after it.", H.lines()[row_of("New text after it.")]) -- (nothing to say)
    end)

    it("with changes tracked, is saved and opens as it was", function()
      local path = H.copy("rich." .. ext)
      H.open(path)
      vim.cmd("Neoffice track on")
      local r = row_of("An image")
      api.nvim_buf_set_lines(0, r - 1, r - 1, false, { "\\vspace{1cm}", "" })
      H.reopen()
      assert.same({ "\\vspace{1cm}" }, shown())
      assert.truthy(H.lines()[row_of("An image")]:find("^An image !%[%]"), H.lines()[row_of("An image")])
    end)
  end)
end

-- a \vspace typed over (the line replaced, "cc"): its paragraph, kept, is written as the new kind,
-- and nothing of the space stays on it (no exact height, no 0 before and after over its style's)
for _, ext in ipairs({ "docx", "odt" }) do
  describe(ext .. " a \\vspace typed over", function()
    local function own(path, text) -- (the paragraph's own spacing, as the file has it)
      local main = zip.read(path, H.MAIN[ext])
      if ext == "docx" then
        local p = vim.iter(H.paragraph_xml(path)):find(function(x) return x:find(text, 1, true) end)
        return p:match("<w:spacing[^>]*/>") or ""
      end
      local name = main:match('<text:[ph] text:style%-name="([^"]+)"[^>]*>' .. vim.pesc(text))
        or main:match('<text:[ph] [^>]*text:style%-name="([^"]+)"[^>]*>' .. vim.pesc(text))
      local st = name and main:match('<style:style style:name="' .. vim.pesc(name) .. '".-</style:style>') or ""
      return (st:match("fo:line%-height=\"[^\"]+cm\"") or "") .. (st:match("fo:margin%-top=\"0cm\"") or "")
    end
    for _, case in ipairs({ { "as a heading", "# A new heading", "A new heading" }, { "as text", "Plain text now.", "Plain text now." },
      { "as a list item", "- An item now", "An item now" }, { "as a quote", "> A quote now.", "A quote now." },
      { "as a caption", "Figure: A caption now", "A caption now" } }) do
      it(case[1] .. ": nothing of the space stays on it", function()
        local path = H.copy("rich." .. ext)
        H.open(path)
        local r
        for i, l in ipairs(H.lines()) do
          if l:find("^Last paragraph") then
            r = i
          end
        end
        vim.api.nvim_buf_set_lines(0, r - 1, r - 1, false, { "\\vspace{0.5cm}", "" })
        H.reopen()
        for i, l in ipairs(H.lines()) do
          if l:find("^\\vspace") then
            r = i
          end
        end
        vim.api.nvim_buf_set_lines(0, r - 1, r, false, { case[2] }) -- (cc: the line typed again)
        H.reopen()
        assert.truthy(vim.tbl_contains(H.lines(), case[2]), table.concat(H.lines(), "\n"))
        assert.same({}, heights(path, ext))
        assert.equals("", own(path, case[3]))
      end)
    end
  end)
end

for _, ext in ipairs({ "docx", "odt" }) do
  describe(ext .. " a paragraph typed right after a \\vspace", function()
    it("isn't formatted like it (no exact height, no 0 before and after)", function()
      local path = H.copy("rich." .. ext)
      H.open(path)
      local n = vim.api.nvim_buf_line_count(0)
      vim.api.nvim_buf_set_lines(0, n, n, false, { "", "\\vspace{0.5cm}", "", "# Right after it" })
      H.reopen()
      assert.same({ "0.50" }, heights(path, ext))
      local main = zip.read(path, H.MAIN[ext])
      if ext == "docx" then
        local p = vim.iter(H.paragraph_xml(path)):find(function(x) return x:find("Right after it", 1, true) end)
        assert.is_nil(p:find('w:lineRule="exact"', 1, true), p)
      else
        local name = main:match('<text:h [^>]*text:style%-name="([^"]+)"[^>]*>Right after it')
        assert.truthy(name, main:sub(-1500))
        local st = main:match('<style:style style:name="' .. vim.pesc(name) .. '".-</style:style>') or ""
        assert.is_nil(st:find('fo:line%-height="0.500cm"'), st)
      end
    end)
  end)
end

describe("a heading of exactly a height, nothing before or after it, of its own (as a form has)", function()
  it("keeps its spacing when its text is edited: it was never vertical space (docx)", function()
    local tight = '<w:p><w:pPr><w:pStyle w:val="Heading1"/><w:spacing w:before="0" w:after="0" w:line="300" w:lineRule="exact"/></w:pPr><w:r><w:t>A tight line.</w:t></w:r></w:p>'
    local path = H.copy("rich.docx", function(name, text)
      return name == "word/document.xml" and (text:gsub("<w:sectPr", tight .. "%0", 1)) or text
    end)
    H.open(path)
    H.replace(vim.fn.index(H.paras(), "# A tight line.") + 1, "tight", "still tight")
    H.reopen()
    local p = vim.iter(H.paragraph_xml(path)):find(function(x) return x:find("still tight", 1, true) end)
    assert.truthy(p:find('w:line="300" w:lineRule="exact"', 1, true), p)
  end)
end)

describe("an empty paragraph of exactly a height in a document", function()
  it("opens as \\vspace when nothing is before or after it, else as a blank line", function()
    local path = H.copy("rich.docx", function(name, text)
      if name ~= "word/document.xml" then
        return text
      end
      local exact = '<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="360" w:lineRule="exact"/></w:pPr></w:p>'
      local spaced = '<w:p><w:pPr><w:spacing w:after="200" w:line="360" w:lineRule="exact"/></w:pPr></w:p>'
      return (text:gsub("<w:sectPr", exact .. "<w:p><w:r><w:t>Between.</w:t></w:r></w:p>" .. spaced .. "%0", 1))
    end)
    H.open(path)
    assert.same({ "\\vspace{18pt}" }, shown())
    local before = H.bytes(path)
    vim.bo.modified = true
    H.save()
    assert.equals(before, H.bytes(path))
  end)
end)
