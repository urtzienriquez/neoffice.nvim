-- Paragraphs' formats: a line with a mark (and attributes) and "+" sets the format of the paragraphs
-- after it, until the next one (as in LaTeX); without "+", of the paragraph right below it only. A
-- paragraph's text never holds its format.
local H = dofile("tests/helpers.lua")
local zip = require("neoffice.zip")
local neo = require("neoffice")
local paras = require("neoffice.paras")
local api = vim.api

describe("paragraph formats", function()
  it("read attributes, and say which aren't", function()
    assert.same({ ["indent-left"] = "1cm", first = "-0.5cm", line = "1.5" }, paras.parse("indent-left = 1cm, first = -0.5cm, line = 1.5"))
    assert.is_nil(paras.parse("shadow = red"))
    assert.is_nil(paras.parse("before = lots"))
    assert.equals("indent-left = 1cm, before = 6pt", paras.render({ before = "6pt", ["indent-left"] = "1cm" }))
  end)

  it("are a run's with what a paragraph says, and a paragraph says only what differs", function()
    local run = { align = "justify", attrs = { first = "0.5cm" } }
    assert.same({ align = "center", attrs = { first = "0.5cm" } }, paras.apply(run, "center", nil))
    assert.same({ attrs = {} }, paras.apply(run, "none", { first = "none" })) -- (none: its style's)
    assert.same({ align = "left", attrs = { first = "0.5cm" } }, paras.apply(run, "left", nil)) -- (left: left, said)
    local align, attrs = paras.diff(run, { align = "justify", attrs = { first = "0.5cm", before = "6pt" } })
    assert.is_nil(align) -- (the run's: unsaid)
    assert.same({ before = "6pt" }, attrs)
    align, attrs = paras.diff(run, { attrs = { first = "0.5cm" } })
    assert.same({ "none" }, { align, attrs })
    assert.is_true(paras.same({ attrs = { first = "5mm" } }, { attrs = { first = "0.5cm" } })) -- (as measured)
  end)

  it("say what changes: what a line doesn't say goes on, none goes back to the style", function()
    local list = neo.paragraphs({ "{align = justify, size = 11pt}+", "One.", "", "{first = 0.5cm}+", "Two.", "",
      "{size = none} +", "Three.", "", "{none, before = 6pt}+", "Four." })
    assert.same({ align = "justify", attrs = { size = "11pt" } }, { align = list[1].format.align, attrs = list[1].format.attrs })
    assert.same({ align = "justify", attrs = { size = "11pt", first = "0.5cm" } }, { align = list[2].format.align, attrs = list[2].format.attrs })
    assert.same({ align = "justify", attrs = { first = "0.5cm" } }, { align = list[3].format.align, attrs = list[3].format.attrs })
    assert.same({ attrs = { before = "6pt" } }, { align = list[4].format.align, attrs = list[4].format.attrs })
    local markup = require("neoffice.markup")
    assert.equals("{first = 0.5cm}", markup.para_line_for(list[1].format, list[2].format))
    assert.equals("{none}", markup.para_line_for(list[2].format, { attrs = {} }))
  end)

  it("go on from a format's line until the next one", function()
    local list = neo.paragraphs({ "One.", "", "<=> {first = 0.5cm}+", "Two.", "", "Three.", "", "<|>+", "Four." })
    local function format(k) return list[k].format end
    assert.same({ attrs = {} }, format(1))
    assert.equals("justify", format(2).align)
    assert.same({ first = "0.5cm" }, format(2).attrs)
    assert.equals(format(2), format(3))
    assert.equals("center", format(4).align)
    assert.same({ "One.", "Two.", "Three.", "Four." }, vim.tbl_map(function(p) return p.text end, list))
  end)

  it("without +, are for the paragraph right below only, over its run's", function()
    local list = neo.paragraphs({ "{align = justify}+", "One.", "", "{size = 9pt}", "<|>", "Two.", "", "Three." })
    assert.same({ align = "justify", attrs = {} }, { align = list[1].format.align, attrs = list[1].format.attrs })
    assert.same({ align = "center", attrs = { size = "9pt" } }, { align = list[2].format.align, attrs = list[2].format.attrs })
    assert.equals(list[1].format, list[3].format) -- (the run's again)
    assert.equals(list[1].format, list[2].run)
    assert.is_nil(list.stray_format)
  end)

  it("without +, for none when no paragraph is right below, or one that isn't plain", function()
    assert.equals(1, neo.paragraphs({ "{size = 9pt}", "", "One." }).stray_format) -- (a blank line between)
    assert.equals(3, neo.paragraphs({ "One.", "", "{size = 9pt}" }).stray_format) -- (nothing after it)
    assert.equals(1, neo.paragraphs({ "{size = 9pt}", "\\newpage {orientation = landscape}", "One." }).stray_format) -- (a section's line between)
    local list = neo.paragraphs({ "{size = 9pt}", "# A heading" })
    assert.truthy(list[1].format.bad and list[1].format.bad:find("no plain paragraph", 1, true))
    assert.is_nil(neo.paragraphs({ "{size = 9pt}", "{before = 6pt}+", "One." }).stray_format) -- (lines of formats together)
  end)
end)

for _, ext in ipairs({ "docx", "odt" }) do
  describe(ext, function()
    local function main() return zip.read(H.doc().path, H.MAIN[ext]) end

    it("sets the format of the paragraphs after a format's line, shows it once, and stops at the next", function()
      H.open(H.copy("rich." .. ext))
      local r = H.row(3)
      api.nvim_buf_set_lines(0, r - 1, r - 1, false, { "<=> {indent-left = 1cm, first = 0.5cm}+" }) -- (a mark: as align =)
      r = H.row(5)
      api.nvim_buf_set_lines(0, r - 1, r - 1, false, { "{}+" }) -- (from here: as their style says)
      H.reopen()
      local lines = H.lines()
      local at = vim.fn.index(lines, "<=> {indent-left = 1cm, first = 0.5cm}+") + 1
      assert.truthy(at > 0, table.concat(lines, "\n"))
      assert.equals("Commented text and more.", lines[at + 1]) -- (no mark of its own)
      assert.is_true(vim.tbl_contains(lines, "{none}+")) -- (back to the style: written so)
      local xml = main()
      if ext == "docx" then
        assert.truthy(xml:find('<w:ind w:left="567" w:firstLine="283"/>', 1, true), xml)
        assert.truthy(xml:find('<w:jc w:val="both"/>', 1, true))
      else
        assert.truthy(xml:find('fo:margin%-left="1cm"') and xml:find('fo:text%-indent="0.5cm"'))
      end
    end)

    it("lets a paragraph differ (only that one), on a line of its own, and taking it away goes back to the run's", function()
      H.open(H.copy("rich." .. ext))
      local r = H.row(3)
      api.nvim_buf_set_lines(0, r - 1, r - 1, false, { "<=>+" })
      r = H.row(4)
      api.nvim_buf_set_lines(0, r - 1, r - 1, false, { "<|>" })
      H.reopen()
      local lines = H.lines()
      r = H.row(4)
      assert.equals("Text with inserted words and deleted words inside.", lines[r]) -- (its text: only text)
      assert.equals("<|>", lines[r - 1])
      assert.equals("Commented text and more.", lines[H.row(3)])
      assert.equals("", lines[H.row(5) - 1]) -- (the one after it: the run's, nothing to say)
      api.nvim_buf_set_lines(0, r - 2, r - 1, false, {})
      H.reopen()
      lines = H.lines()
      assert.equals("", lines[H.row(4) - 1])
      -- (said in its text, as before: read, and shown on a line of its own)
      H.replace(4, "Text with", "<|> Text with")
      H.reopen()
      assert.equals("<|>", H.lines()[H.row(4) - 1])
    end)

    it("doesn't write an attribute it doesn't know, and says which it knows", function()
      local path = H.copy("rich." .. ext)
      H.open(path)
      local before = H.bytes(path)
      local r = H.row(3)
      api.nvim_buf_set_lines(0, r - 1, r - 1, false, { "<=> {shadow = red}" })
      local errors = {}
      local notify = vim.notify
      vim.notify = function(msg) errors[#errors + 1] = msg end
      H.save()
      vim.notify = notify
      assert.equals(before, H.bytes(path))
      assert.truthy(table.concat(errors, "\n"):find("indent-left", 1, true))
    end)
  end)
end

describe("the lines that show paragraphs' formats", function()
  local markup = require("neoffice.markup")
  local function lines_of(blocks) -- (the lines shown, as text)
    local out = {}
    local run = { attrs = {} }
    for i, line in ipairs(neo.format_lines(blocks)) do
      if line.run then
        out[#out + 1] = { i, markup.para_line_for(run, line.run) .. "+" }
        run = line.run
      end
      if line.own then
        out[#out + 1] = { i, markup.render_para_line(line.own) }
      end
    end
    return out
  end

  it("say an attribute where it changes, not again because another one does (the centred names, then text)", function()
    -- (two centred lines, a justified abstract, then a paragraph of the usual style and two of Body Text)
    local shown = lines_of({ { align = "center" }, { align = "center" }, { align = "justify" }, {}, { para = { style = "Body Text" } }, { para = { style = "Body Text" } } })
    local text = table.concat(vim.tbl_map(function(l) return l[2] end, shown), "\n")
    assert.equals(4, #shown, text)
    assert.is_nil(text:find("align = none", 1, true), text)
    assert.is_nil(vim.iter(shown):find(function(l) return l[1] == 4 end), text) -- (the first after them: nothing to say)
    assert.truthy(text:find("{style = Body Text}+", 1, true), text)
  end)

  it("keep a run's alignment once when the fonts change from one paragraph to the next", function()
    local blocks = {}
    for k = 1, 6 do
      blocks[k] = { align = "justify", para = { font = k % 2 == 1 and "Arial" or "Cambria" } }
    end
    local shown = lines_of(blocks)
    local text = table.concat(vim.tbl_map(function(l) return l[2] end, shown), "\n")
    assert.truthy(#shown <= 4, text)
    local _, aligned = text:gsub("<=>", "")
    assert.equals(1, aligned, text)
  end)

  it("end a run once, and say only what a paragraph has besides (not what it no longer has of the run)", function()
    -- (indented paragraphs, then one of another style, then plain ones: as in a document with a list)
    local shown = lines_of({ { align = "justify" }, { para = { first = "0cm" } }, { para = { first = "0cm" } }, { para = { first = "0cm" } },
      { para = { style = "First Paragraph" } }, {}, {} })
    local text = table.concat(vim.tbl_map(function(l) return l[2] end, shown), "\n")
    assert.is_nil(text:find("first = none", 1, true), text)
    assert.equals(4, #shown, text)
  end)

  it("are never more than a whole format at a time would make (random paragraphs)", function()
    local formats = { {}, { align = "center" }, { align = "justify" }, { para = { style = "Body Text" } },
      { align = "justify", para = { style = "Body Text" } }, { para = { font = "Arial" } }, { align = "center", para = { size = "10pt" } } }
    local function old_count(blocks) -- (the rule before: a run of a whole format, or the paragraph's own)
      local pf, run, n = require("neoffice.paras"), { attrs = {} }, 0
      for i, b in ipairs(blocks) do
        local own = { align = b.align, attrs = b.para or {} }
        local function shared(j)
          local o = blocks[i + j]
          return o and pf.same(own, { align = o.align, attrs = o.para or {} })
        end
        if not pf.same(own, run) then
          n = n + 1
          if shared(1) or shared(2) then
            run = own
          end
        end
      end
      return n
    end
    math.randomseed(7)
    for _ = 1, 300 do
      local blocks = {}
      for i = 1, math.random(1, 9) do
        blocks[i] = vim.deepcopy(formats[math.random(#formats)])
      end
      local shown = lines_of(blocks)
      assert.truthy(#shown <= old_count(blocks), vim.inspect(blocks))
      -- (and they say each paragraph's format: what the lines give is what it has)
      local run, pf = { attrs = {} }, require("neoffice.paras")
      for i, line in ipairs(neo.format_lines(blocks)) do
        run = line.run or run
        local got = line.own and markup.apply_para_line(run, line.own) or run
        assert.is_true(pf.same(got, { align = blocks[i].align, attrs = blocks[i].para or {} }), vim.inspect({ blocks, i }))
      end
    end
  end)
end)

for _, ext in ipairs({ "docx", "odt" }) do
  describe(ext .. " formats shown after a centred run, a heading, and paragraphs of two styles", function()
    it("say no alignment again where none changes, and the style where it does", function()
      local path = H.copy("rich." .. ext)
      H.open(path)
      local n = api.nvim_buf_line_count(0)
      api.nvim_buf_set_lines(0, n, n, false, { "", "<|>+", "Someone.", "", "Somewhere.", "", "# A new section", "", "{none, style = First Paragraph}+",
        "One after the heading.", "", "{style = none}+", "Two.", "", "Three." })
      H.reopen()
      local lines = H.lines()
      local at = vim.fn.index(lines, "# A new section") + 1
      assert.truthy(at > 0, table.concat(lines, "\n"))
      local after = vim.tbl_filter(function(l) return l:find("^[{<]") ~= nil end, vim.list_slice(lines, at))
      assert.is_nil(table.concat(after, "\n"):find("align", 1, true), table.concat(lines, "\n"))
      assert.truthy(#after <= 1, table.concat(after, "\n"))
      local before = H.bytes(path)
      vim.bo.modified = true
      H.save()
      assert.equals(before, H.bytes(path))
    end)
  end)
end
