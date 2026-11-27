-- Equations ($...$, $$...$$), superscripts and subscripts.
local H = dofile("tests/helpers.lua")
local markup = require("neoffice.markup")
local zip = require("neoffice.zip")
local api = vim.api

describe("markup", function()
  it("reads superscripts, subscripts and equations", function()
    local _, chars = markup.parse("cm^{2} W_{f} $x^2$ $$y$$")
    assert.is_true(chars[3].sup)
    assert.is_true(chars[6].sub)
    assert.same({ tex = "x^2", display = false }, chars[8].math)
    assert.same({ tex = "y", display = true }, chars[10].math)
  end)

  it("leaves dollar signs that aren't equations as text", function()
    for _, line in ipairs({ "$5 and $10", "costs $ 3", "a $ b $ c", "lone $", "$x $" }) do
      local _, chars = markup.parse(line)
      assert.is_false(vim.iter(chars):any(function(c) return c.math ~= nil end), line)
    end
  end)

  it("escapes $, ^{, _{ (and } inside them) in text, and writes them as in LaTeX", function()
    local line = markup.render({ { text = "$5 a_{b} snake_case 2^3 " }, { text = "}2", sup = true }, { text = "x", sub = true, i = true } })
    assert.equals("\\$5 a\\_{b} snake_case 2^3 ^{\\}2}*_{x}*", line)
    local _, chars = markup.parse(line)
    assert.equals("$5 a_{b} snake_case 2^3 }2x", table.concat(vim.tbl_map(function(c) return c.ch end, chars)))
    assert.is_true(chars[#chars - 1].sup)
    assert.is_true(chars[#chars].sub and chars[#chars].i)
  end)
end)

for _, ext in ipairs({ "docx", "odt" }) do
  describe(ext .. " superscripts and subscripts", function()
    before_each(function()
      H.open(H.copy("math." .. ext))
    end)

    it("show as ^{...} and _{...}", function()
      assert.equals("Area in cm^{2} and a rate in mL h^{-1} of W_{food}.", H.paras()[1])
    end)

    it("are kept, added and removed", function()
      H.replace(1, "cm^{2}", "m^{3}")
      H.replace(1, "W_{food}", "W_{meat} and CO_{2} and x^{n}")
      H.replace(1, "h^{-1}", "h-1")
      H.reopen()
      assert.equals("Area in m^{3} and a rate in mL h-1 of W_{meat} and CO_{2} and x^{n}.", H.paras()[1])
    end)

    it("combine with other formatting", function()
      H.set(4, "**bold^{sup}** *W_{i}*") -- (its {style = ...} gone with the line: the usual one there)
      H.reopen()
      assert.equals("**bold^{sup}** *W_{i}*", H.paras()[4])
    end)
  end)
end

describe("docx equations", function()
  local path

  before_each(function()
    path = H.copy("math.docx")
    H.open(path)
  end)

  if not H.has_pandoc then
    it("show as ￼ without pandoc, and are kept as they are", function()
      assert.equals("Inline ￼ here.", H.paras()[2])
      assert.equals("{style = First Paragraph} Prices: \\$5 and \\$10.", H.paras()[4]) -- (as the fixture has it: not where it goes)
      local before = H.paragraph_xml(path)
      H.replace(2, "Inline", "An inline")
      H.reopen()
      assert.equals("An inline ￼ here.", H.paras()[2])
      assert.equals(before[2]:match("<m:oMath>.-</m:oMath>"), H.paragraph_xml(path)[2]:match("<m:oMath>.-</m:oMath>"))
    end)
  end

  if H.has_pandoc then
    it("show as LaTeX, inline and on their own line", function()
      assert.equals("Inline $x^{2} + y_{1}$ here.", H.paras()[2])
      assert.equals("$$B = W_{food} + W_{met} - E$$", H.paras()[3])
      assert.equals("{style = First Paragraph} Prices: \\$5 and \\$10.", H.paras()[4]) -- (as the fixture has it: not where it goes)
    end)
  end

  it("are kept exactly as they were when not edited", function()
    local before = H.paragraph_xml(path)
    H.replace(2, "Inline", "An inline")
    H.save()
    local after = H.paragraph_xml(path)
    assert.equals(before[3], after[3])
    assert.equals(before[2]:match("<m:oMath>.-</m:oMath>"), after[2]:match("<m:oMath>.-</m:oMath>"))
  end)

  if H.has_pandoc then -- (these need pandoc: LaTeX <-> Word's equations)
    it("are rewritten from edited LaTeX", function()
      H.replace(3, "- E", "- E - \\alpha^{2}")
      H.replace(2, "y_{1}", "\\frac{a}{b}")
      H.reopen()
      assert.equals("$$B = W_{food} + W_{met} - E - \\alpha^{2}$$", H.paras()[3])
      assert.equals("Inline $x^{2} + \\frac{a}{b}$ here.", H.paras()[2])
      local main = zip.read(path, "word/document.xml")
      assert.truthy(main:find("<m:oMathPara>", 1, true))
      assert.truthy(main:find("<m:f>", 1, true)) -- (the fraction)
    end)

    it("are added by typing LaTeX", function()
      api.nvim_buf_set_lines(0, H.row(4) - 1, H.row(4) - 1, false, { "Energy is $E = mc^{2}$.", "", "$$\\sum_{i=1}^{n} x_i$$", "" })
      H.save()
      local first = H.bytes(path)
      H.save() -- (saving again changes nothing)
      assert.equals(first, H.bytes(path))
      H.open(path)
      assert.equals("Energy is $E = mc^{2}$.", H.paras()[4])
      assert.truthy(H.paras()[5]:match("^%$%$.*sum"))
    end)
  end

  it("are removed by deleting their text", function()
    H.set(3, "No equation any more.")
    H.reopen()
    assert.equals("No equation any more.", H.paras()[3])
    assert.equals(1, select(2, zip.read(path, "word/document.xml"):gsub("<m:oMath>", "")))
  end)

  if H.has_pandoc then
    it("that can't be converted make the save fail, changing nothing", function()
      local before = H.bytes(path)
      H.replace(3, "B =", "\\frac{")
      H.save()
      assert.equals(before, H.bytes(path))
      assert.is_true(vim.bo.modified)
    end)
  end

  if H.has_libreoffice and H.has_pandoc then
    it("LibreOffice reads the rewritten equations", function()
      H.replace(3, "- E", "- E - F")
      H.save()
      local dir = vim.fn.tempname() -- (as .odt: LibreOffice turns each equation into a formula object)
      vim.system({ "soffice", "-env:UserInstallation=file://" .. dir .. "/p", "--headless", "--convert-to", "odt", "--outdir", dir, path }):wait()
      local formulas = {}
      for name in pairs(zip.list(dir .. "/math.odt")) do
        if name:match("^Object %d+/content.xml$") then
          formulas[#formulas + 1] = zip.read(dir .. "/math.odt", name)
        end
      end
      assert.equals(2, #formulas)
      assert.is_true(vim.iter(formulas):any(function(f) return f:find("<mi>F</mi>", 1, true) ~= nil end))
    end)
  end
end)

describe("odt formulas", function()
  it("are kept as objects (and not mistaken for images)", function()
    local path = H.copy("math.odt")
    H.open(path)
    assert.equals("Inline ⟦equation⟧ here.", H.paras()[2]) -- (a formula: shown by name, kept)
    local before = H.paragraph_xml(path)[2]:match("<draw:frame.-</draw:frame>")
    H.replace(2, "Inline", "An inline")
    H.save()
    assert.equals(before, H.paragraph_xml(path)[2]:match("<draw:frame.-</draw:frame>"))
  end)

  it("can't be typed: the save says so and changes nothing", function()
    local path = H.copy("math.odt")
    local before = H.bytes(path)
    H.open(path)
    H.set(4, "Now $x$ is here.")
    H.save()
    assert.equals(before, H.bytes(path))
  end)
end)
