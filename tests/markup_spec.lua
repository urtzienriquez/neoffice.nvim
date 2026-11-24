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
    assert.equals("a **bold** *it* <u>u</u> [site](https://x.org)", line)
  end)

  it("nests marks that open together", function()
    assert.equals("***<u>all</u>***", markup.render({ span("all", { b = true, i = true, u = true }) }))
    assert.equals("**<u>x</u>**", markup.render({ span("x", { b = true, u = true }) }))
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

  it("reports the columns of each span", function()
    local _, cols = markup.render({ span("a"), span("b", { b = true }) })
    assert.same({ { 0, 1 }, { 3, 4 } }, cols)
  end)

  it("reads unbalanced marks without failing", function()
    for _, line in ipairs({ "**", "*", "<u>", "[x](", "[x]", "![", "![a](", "\\", "# ", "#", "a **b", "[a](b" }) do
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
    assert.same({ first = 1, last = 2, owns = 3, text = "a b" }, p[1]) -- (and the blank line after it)
    assert.same({ first = 4, last = 4, owns = 4, text = "c" }, p[2])
    assert.same({ first = 5, last = 5, owns = 5, text = "" }, p[3]) -- blank lines at the end: empty paragraphs
    assert.same({ first = 6, last = 6, owns = 6, text = "" }, p[4])
  end)
end)
