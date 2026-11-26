local markup = require("neoffice.markup")
local neo = require("neoffice")

-- parse() back into spans of equal formatting, for comparison
local function spans_of(line)
  local level, chars = markup.parse(line)
  local spans = {}
  for _, c in ipairs(chars) do
    local last = spans[#spans]
    if c.image then
      spans[#spans + 1] = { image = c.image }
    elseif last and last.text and last.b == c.b and last.i == c.i and last.u == c.u and last.link == c.link then
      last.text = last.text .. c.ch
    else
      spans[#spans + 1] = { text = c.ch, b = c.b, i = c.i, u = c.u, link = c.link }
    end
  end
  return spans, level -- (the block)
end

local function span(text, attrs)
  return vim.tbl_extend("force", { text = text, b = false, i = false, u = false }, attrs or {})
end

local function merged(spans) -- (neighbours with the same formatting read back as one)
  local out = {}
  for _, s in ipairs(vim.deepcopy(spans)) do
    local last = out[#out]
    if last and last.text and s.text and last.b == s.b and last.i == s.i and last.u == s.u and last.link == s.link then
      last.text = last.text .. s.text
    else
      out[#out + 1] = s
    end
  end
  return out
end

local function roundtrip(spans, block)
  local line = markup.render(vim.deepcopy(spans), block)
  local back, back_block = spans_of(line)
  assert.same(merged(spans), back, line)
  assert.same(block or {}, back_block)
  return line
end

describe("markup", function()
  it("renders formatting as marks", function()
    local line = markup.render({ span("a "), span("bold", { b = true }), span(" "), span("it", { i = true }), span(" "), span("u", { u = true }), span(" "), span("site", { link = "https://x.org" }) })
    assert.equals("a **bold** *it* ++u++ [site](https://x.org)", line)
  end)

  it("nests marks that open together", function()
    assert.equals("***++all++***", markup.render({ span("all", { b = true, i = true, u = true }) }))
    assert.equals("**++x++**", markup.render({ span("x", { b = true, u = true }) }))
  end)

  it("round-trips overlapping formatting, links and escapes", function()
    roundtrip({
      span("plain *star* [x] \\ <u> "),
      span("bold", { b = true }),
      span("bold italic", { b = true, i = true }),
      span("italic", { i = true }),
      span("link (with) bold", { link = "https://x.org/a_(b)", b = true }),
      span("under", { u = true }),
    }, { heading = 2 })
  end)

  it("round-trips every combination of formatting next to each other", function()
    local combos = {}
    for b = 0, 1 do
      for i = 0, 1 do
        for u = 0, 1 do
          combos[#combos + 1] = { b = b == 1, i = i == 1, u = u == 1 }
        end
      end
    end
    for _, a in ipairs(combos) do
      for _, b in ipairs(combos) do
        roundtrip({ span("x", a), span("y", b), span("z", a) })
      end
    end
  end)

  it("round-trips adjacent links and images", function()
    roundtrip({ span("a", { link = "https://a.org" }), span("b", { link = "https://b.org" }), { image = { alt = "pic [1]", path = "media/i (2).png" } }, span("!") })
  end)

  it("escapes what would be read as marks", function()
    roundtrip({ span("# 1 fan, 2*3, a [b], <u>, \\, end!") })
    roundtrip({ span("- not a list") })
    roundtrip({ span("12. not a list either") })
    roundtrip({ span("% not a title") })
    roundtrip({ span("Wow!"), span("x", { link = "https://x.org" }) })
    assert.equals("\\# 1", markup.render({ span("# 1") }))
    assert.equals("#1", markup.render({ span("#1") })) -- (not a heading: no need)
  end)

  it("writes brackets as they are, unless they would start a link, an image or a footnote", function()
    assert.equals("J Comp Physiol [B] 173, refs [1, 2]", markup.render({ span("J Comp Physiol [B] 173, refs [1, 2]") }))
    assert.equals("not a \\[link](x)", markup.render({ span("not a [link](x)") }))
    assert.equals("x^\\[n]", markup.render({ span("x^[n]") }))
  end)

  it("keeps a paragraph of only a tab from looking blank, and reads it back as a tab", function()
    local line = markup.render({ span("\t") })
    assert.equals("\\\t", line)
    local _, chars = markup.parse(line)
    assert.equals(1, #chars)
    assert.equals("\t", chars[1].ch)
    local _, typed = markup.parse(line .. "text") -- (typed after it: still no backslash)
    assert.equals("\ttext", table.concat(vim.tbl_map(function(c) return c.ch end, typed)))
  end)

  it("keeps a paragraph of only spaces from looking blank", function()
    local line, cols = markup.render({ span("   ") })
    assert.equals("\\   ", line)
    assert.same({ { 1, 4 } }, cols)
    assert.equals("   ", spans_of(line)[1].text)
  end)

  it("reads images", function()
    local _, chars = markup.parse("see ![a \\] b](dir/x).png) end")
    assert.same({ alt = "a ] b", path = "dir/x" }, chars[5].image)
  end)

  it("reads an image's width with or without spaces, and writes it with them", function()
    for _, text in ipairs({ "![a](x.png){width=3cm}", "![a](x.png){width = 3cm}", "![a](x.png){ width= 3cm }" }) do
      local _, chars = markup.parse(text)
      assert.equals("3cm", chars[1].image.width, text)
    end
    assert.equals("![a](x.png){width = 3cm}", markup.render({ { image = { alt = "a", path = "x.png", width = "3cm" } } }))
  end)

  it("reads attributes: name = value or name=value, separated by commas, or values alone", function()
    local list, tidy = markup.attributes("width=full,height = tall  grid", function(v) return v == "grid" and "rules" or nil end)
    assert.same({ "width", "height", nil }, { list[1].name, list[2].name, list[3].name })
    assert.same({ "full", "tall", "grid" }, { list[1].value, list[2].value, list[3].value })
    assert.equals("width = full, height = tall, rules = grid", tidy)
  end)

  it("reports the columns of each span", function()
    local _, cols = markup.render({ span("a"), span("b", { b = true }) })
    assert.same({ { 0, 1 }, { 3, 4 } }, cols)
  end)

  it("highlights in colours: ==text== is yellow, ==text=={green} green", function()
    local line = markup.render({ span("a"), span("lit", { h = "yellow" }), span("b"), span("g", { h = "green" }), span("r", { h = "#a1b2c3" }) })
    assert.equals("a==lit==b==g=={green}==r=={#a1b2c3}", line)
    local _, chars = markup.parse(line)
    assert.same({ false, "yellow", "yellow", "yellow", false, "green", "#a1b2c3" }, vim.tbl_map(function(c) return c.h end, chars))
    local _, upper = markup.parse("==x=={Green}")
    assert.equals("green", upper[1].h)
    local _, unknown = markup.parse("==x=={nope}") -- (not a colour: yellow, then text)
    assert.same({ "yellow", false, false }, { unknown[1].h, unknown[2].h, unknown[#unknown].h })
    -- text that reads as a colour after a highlight is escaped
    local escaped = markup.render({ span("x", { h = "yellow" }), span("{green}") })
    local _, back = markup.parse(escaped)
    assert.same({ "yellow", false }, { back[1].h, back[2].h })
    assert.equals("x{green}", table.concat(vim.tbl_map(function(c) return c.ch end, back)))
  end)

  it("underlines with ++, and keeps a lone ++ as text", function()
    assert.equals("C++ is fine", markup.render({ span("C++ is fine") }))
    assert.equals("++fine++, in C++", markup.render({ span("fine", { u = true }), span(", in C++") }))
    roundtrip({ span("C++ is "), span("fine", { u = true }) }) -- (escaped: C\+\+)
    roundtrip({ span("C++ and C++ again") })
    roundtrip({ span("a "), span("u", { u = true }), span(" ++ b") })
  end)

  it("reads unbalanced marks without failing", function()
    for _, line in ipairs({ "**", "*", "++", "==x=={", "<u>", "[x](", "[x]", "![", "![a](", "\\", "# ", "#", "a **b", "[a](b" }) do
      assert.has_no.errors(function() markup.parse(line) end)
    end
  end)
end)

describe("paragraph layout", function()
  local function texts(lines)
    return vim.tbl_map(function(p) return p.text end, neo.paragraphs(lines))
  end

  it("separates paragraphs by blank lines and joins wrapped lines", function()
    assert.same({ "one two", "three" }, texts({ "one", "two", "", "three" }))
  end)

  it("reads further blank lines as empty paragraphs", function()
    assert.same({ "", "a", "", "b", "" }, texts({ "", "a", "", "", "b", "" }))
    assert.same({}, texts({}))
  end)

  it("treats lines of spaces as blank", function()
    assert.same({ "a", "b" }, texts({ "a", "  ", "b" }))
  end)

  it("gives each paragraph the rows whose marks belong to it", function()
    local p = neo.paragraphs({ "a", "b", "", "c", "", "" })
    for _, para in ipairs(p) do -- (their format: see "paragraph formats")
      para.format = nil
    end
    assert.same({ first = 1, last = 2, owns = 3, text = "a b" }, p[1]) -- (and the blank line after it)
    assert.same({ first = 4, last = 4, owns = 4, text = "c" }, p[2])
    assert.same({ first = 5, last = 5, owns = 5, text = "" }, p[3]) -- blank lines at the end: empty paragraphs
    assert.same({ first = 6, last = 6, owns = 6, text = "" }, p[4])
  end)
end)

describe("what is written reads back as it was", function()
  local markup = require("neoffice.markup")
  local TEXTS = {
    "\t", " ", "  ", "\t x", " x", "1. x", "12. x", "- x", "  - x", "# x", "% x", "> x", "<=> x", "<|> x", "->> x",
    "<<- x", "<=>", "{x} y", "{first = 1cm} y", "Table: x", "Figure: x", ":::", "::: {float = right}", "|a|", "| a | b |",
    "{a}", "\\", "\\\\", "==x==", "++x++", "~~x~~", "*x*", "x\ny", "x\n- y", "x\n# y", "a ==b=={green} c", "![a](b)",
    "x\\", "[x](y)", "$x$", "^[n]", "⟦f⟧", "C++ and C++", "n = 20", "a_{b}", "x^{2}",
    "J Physiol [B] 173", "[1]", "[1, 2] and [3]", "[a] [b](c)", "a]b[c", "[x]", "![x](y)", "x^[n]", "[[x]](y)", "]",
    "[", "[x] (y)", "\\[ x \\]",
  }
  local BLOCKS = {
    {}, { align = "justify" }, { align = "center" }, { align = "left" }, { align = "justify", para = { first = "1cm" } },
    { heading = 2 }, { title = true }, { quote = true }, { list = { kind = "bullet", level = 1 } }, { caption = "table" },
  }
  it("for any text in any kind of paragraph", function()
    for _, block in ipairs(BLOCKS) do
      for _, text in ipairs(TEXTS) do
        local spans = {}
        for k, part in ipairs(vim.split(text, "\n", { plain = true })) do
          if k > 1 then
            spans[#spans + 1] = { line_break = true }
          end
          if part ~= "" then
            spans[#spans + 1] = { text = part }
          end
        end
        local line = markup.render(vim.deepcopy(spans), vim.deepcopy(block))
        local back, chars = markup.parse(line)
        local got = table.concat(vim.tbl_map(function(c) return c.ch end, chars))
        local what = vim.inspect(block) .. " " .. vim.inspect(text) .. " -> " .. vim.inspect(line)
        assert.equals(text, got, what)
        for _, key in ipairs({ "heading", "title", "quote", "caption", "align" }) do
          assert.equals(block[key], back[key], what .. " (" .. key .. ")")
        end
        assert.same(block.para, back.para, what .. " (para)")
        assert.equals(block.list and block.list.kind, back.list and back.list.kind, what .. " (list)")
      end
    end
  end)
end)
