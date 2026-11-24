-- Editing text, formatting, links, headings and paragraphs, in both formats.
local H = dofile("tests/helpers.lua")
local zip = require("neoffice.zip")
local api = vim.api

local ORIGINAL = {
  "# Introduction",
  "This is **bold**, *italic*, <u>underlined</u> and a [link](https://example.org). Tom & Jerry <3, 2\\*3.",
  "Commented text and more.",
  "Text with inserted words and deleted words inside.",
  "An image ![](IMAGE) here.",
  "- first item",
  "- second item",
  "A",
  "B",
  "cell one",
  "cell two",
  "Last paragraph.",
}

for _, ext in ipairs({ "docx", "odt" }) do
  describe(ext, function()
    before_each(function()
      H.open(H.copy("rich." .. ext))
    end)

    local function paras()
      return vim.tbl_map(function(t) return (t:gsub("{width=[^}]*}", ""):gsub("!%[%]%(.-%)", "![](IMAGE)")) end, H.paras())
    end

    it("opens as text with marks, paragraphs separated by blank lines", function()
      assert.same(ORIGINAL, paras())
      assert.equals("", H.lines()[2])
      assert.equals(#ORIGINAL * 2 - 6, #H.lines()) -- (none between list items; a table row per line)
    end)

    it("keeps run properties when editing inside a run", function()
      H.replace(2, "link", "linky")
      H.replace(2, "bold", "strong")
      H.reopen()
      assert.equals("This is **strong**, *italic*, <u>underlined</u> and a [linky](https://example.org). Tom & Jerry <3, 2\\*3.", H.paras()[2])
    end)

    it("text typed at the end of a word takes its formatting", function()
      H.replace(2, "bold**", "boldest**")
      H.reopen()
      assert.truthy(H.paras()[2]:find("**boldest**", 1, true))
    end)

    it("writes formatting typed as marks, and removes it when the marks are deleted", function()
      H.set(3, "Plain **bold *both*** <u>under</u> [x](https://x.org/a) \\*literal\\*")
      H.replace(2, "**bold**", "bold")
      H.replace(2, "*italic*", "italic")
      H.reopen()
      assert.equals("Plain **bold *both*** <u>under</u> [x](https://x.org/a) \\*literal\\*", H.paras()[3])
      assert.truthy(H.paras()[2]:find("This is bold, italic, <u>", 1, true))
    end)

    it("changes, adds and removes links", function()
      H.replace(2, "https://example.org", "https://other.org")
      H.set(12, "See [the site](https://new.example.org/path?a=1&b=2) now.")
      H.replace(3, "Commented", "[Commented](https://x.org)")
      H.reopen()
      assert.truthy(H.paras()[2]:find("[link](https://other.org)", 1, true))
      assert.equals("See [the site](https://new.example.org/path?a=1&b=2) now.", H.paras()[12])
      H.replace(12, "[the site](https://new.example.org/path?a=1&b=2)", "the site")
      H.reopen()
      assert.equals("See the site now.", H.paras()[12])
    end)

    it("adds, changes and removes headings", function()
      H.set(1, "## Overview")
      H.set(12, "### Last heading")
      H.set(6, "first item, no longer a heading")
      H.reopen()
      assert.equals("## Overview", H.paras()[1])
      assert.equals("### Last heading", H.paras()[12])
      H.set(1, "Overview")
      H.reopen()
      assert.equals("Overview", H.paras()[1])
    end)

    it("adds, removes, splits and joins paragraphs", function()
      -- from the bottom up, so the rows above stay where they are
      H.set(12, "Last\n\nparagraph.") -- split
      H.replace(9, "B", "") -- (its table cell keeps an empty paragraph)
      local r = H.row(4) -- join paragraph 4 with 3
      api.nvim_buf_set_lines(0, r - 3, r, false, { "Commented text and more. Text with inserted words and  inside." })
      api.nvim_buf_set_lines(0, H.row(2) - 1, H.row(2) - 1, false, { "A new paragraph.", "" })
      H.reopen()
      local p = H.paras()
      assert.equals("A new paragraph.", p[2])
      assert.equals("Commented text and more. Text with inserted words and  inside.", p[4])
      assert.equals("", p[9]) -- the emptied cell
      assert.same({ "Last", "paragraph." }, { p[#p - 1], p[#p] })
    end)

    it("joins lines without a blank line between them into one paragraph", function()
      H.set(12, "Last\nparagraph,\nwrapped by hand.")
      H.reopen()
      assert.equals("Last paragraph, wrapped by hand.", H.paras()[12])
    end)

    it("keeps, adds and removes empty paragraphs", function()
      api.nvim_buf_set_lines(0, H.row(12) - 1, H.row(12) - 1, false, { "", "" }) -- two empty paragraphs
      H.reopen()
      assert.same({ "cell two", "", "", "Last paragraph." }, vim.list_slice(H.paras(), 11, 14))
      api.nvim_buf_set_lines(0, H.row(12) - 1, H.row(12), false, {}) -- remove one
      H.reopen()
      assert.same({ "cell two", "", "Last paragraph." }, vim.list_slice(H.paras(), 11, 13))
    end)

    it("keeps list items and table cells what they are", function()
      H.replace(6, "first", "First")
      H.replace(10, "cell one", "Cell One")
      H.save()
      local main = zip.read(H.doc().path, H.MAIN[ext])
      if ext == "docx" then
        assert.truthy(main:find("<w:numPr>.-First item"))
        assert.truthy(main:find("<w:tc>.-Cell One"))
      else
        assert.truthy(main:find("<text:list%-item><text:p[^>]*>First item"))
        assert.truthy(main:find("<table:table%-cell[^>]*><text:p[^>]*>Cell One"))
      end
    end)

    it("keeps any text: unicode, symbols, tabs, runs of spaces", function()
      local text = "Ñandú — 日本語 😀 & < > \" ' \\* \\[x\\] # a\tb  c   d"
      H.set(12, text)
      H.reopen()
      assert.equals(text, H.paras()[12])
    end)

    it("keeps an empty paragraph you type in a paragraph of its own", function()
      api.nvim_buf_set_lines(0, H.row(12) - 1, H.row(12) - 1, false, { "" }) -- (an empty paragraph)
      H.reopen()
      local r = H.row(12) -- (its line, right before "Last paragraph.")
      assert.same({ "", "Last paragraph." }, { H.lines()[r], H.lines()[r + 1] })
      api.nvim_buf_set_text(0, r - 1, 0, r - 1, 0, { "Typed in it." })
      assert.same({ "Typed in it.", "Last paragraph." }, vim.list_slice(H.paras(), 12, 13))
      H.reopen()
      assert.same({ "Typed in it.", "Last paragraph." }, vim.list_slice(H.paras(), 12, 13))
    end)

    it("keeps a paragraph of only spaces, and a literal # or ! before a link", function()
      H.set(12, "\\   ")
      H.replace(11, "cell two", "\\# 1 [x](https://x.org)")
      H.replace(10, "cell one", "Wow\\![x](https://x.org)")
      H.reopen()
      assert.equals("\\   ", H.paras()[12])
      assert.equals("\\# 1 [x](https://x.org)", H.paras()[11])
      assert.equals("Wow\\![x](https://x.org)", H.paras()[10])
    end)

    it("leaves a valid document when everything is deleted", function()
      api.nvim_buf_set_lines(0, 0, -1, false, {})
      H.reopen()
      assert.same({ "" }, H.lines())
      api.nvim_buf_set_lines(0, 0, -1, false, { "Again." })
      H.reopen()
      assert.same({ "Again." }, H.paras())
    end)
  end)
end

describe("docx headings", function()
  it("are found by style name, whatever their id", function()
    H.open(H.copy("rich.docx", function(_, s) return (s:gsub("Heading(%d)", "Berschrift%1")) end))
    assert.equals("# Introduction", H.paras()[1])
    api.nvim_buf_set_lines(0, 1, 1, false, { "", "## Sub" })
    H.reopen()
    assert.equals("## Sub", H.paras()[2])
    assert.truthy(zip.read(H.doc().path, "word/document.xml"):find('w:pStyle w:val="Berschrift2"'))
  end)

  it("get a style added when the document has none for that level", function()
    H.open(H.copy("rich.docx", function(name, s)
      return name == "word/styles.xml" and (s:gsub('<w:style w:styleId="Heading4".-</w:style>', "")) or s
    end))
    api.nvim_buf_set_lines(0, 1, 1, false, { "", "#### Deep" })
    H.reopen()
    assert.equals("#### Deep", H.paras()[2])
    assert.truthy(zip.read(H.doc().path, "word/styles.xml"):find('w:val="heading 4"'))
  end)
end)

describe("odt", function()
  it("writes runs of spaces and leading spaces so that they survive", function()
    H.open(H.copy("rich.odt"))
    H.set(12, " lead  two   three")
    H.save()
    assert.truthy(zip.read(H.doc().path, "content.xml"):find("<text:s/>lead <text:s/>two <text:s/><text:s/>three", 1, true))
  end)

  it("makes a style when bold is added to text with another style", function()
    H.open(H.copy("rich.odt"))
    H.replace(2, "<u>underlined</u>", "**<u>underlined</u>**")
    H.reopen()
    assert.truthy(H.paras()[2]:find("**<u>underlined</u>**", 1, true))
  end)
end)

describe("titles", function()
  for _, ext in ipairs({ "docx", "odt" }) do
    it("show as % title in " .. ext .. " (highlighted as NeofficeTitle), and stay titles when edited", function()
      local path = H.copy("rich." .. ext, function(name, s)
        if name == "word/document.xml" then
          return (s:gsub('<w:pStyle w:val="Heading1" />', '<w:pStyle w:val="Title" />', 1))
        elseif name == "content.xml" then
          return (s:gsub('<text:h text:style%-name="Heading_20_1" text:outline%-level="1">(.-)</text:h>', '<text:p text:style-name="Title">%1</text:p>', 1))
        end
        return s
      end)
      H.open(path)
      assert.equals("% Introduction", H.paras()[1])
      local ns = api.nvim_get_namespaces().neoffice_decorations
      local marks = api.nvim_buf_get_extmarks(0, ns, 0, -1, { details = true })
      assert.equals(1, #marks)
      assert.equals("NeofficeTitle", marks[1][4].hl_group)
      assert.same({ 0, 0, 0, #"% Introduction" }, { marks[1][2], marks[1][3], marks[1][4].end_row, marks[1][4].end_col })
      H.replace(1, "Introduction", "An introduction")
      H.reopen()
      assert.equals("% An introduction", H.paras()[1])
    end)
  end
end)
