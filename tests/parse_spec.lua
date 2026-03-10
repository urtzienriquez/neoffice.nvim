-- tests/parse_spec.lua

local assert = require("luassert")
local direct = require("neoffice.direct")

local FIXTURES = "tests/fixtures/"

describe("to_text", function()
  describe("simple file", function()
    local file_text

    before_each(function()
      file_text = direct.to_text(FIXTURES .. "simple-file.odt")
    end)

    it("parses simple text correctly", function()
      assert.equals("Add a simple line of text. And something else.", file_text)
    end)
  end)

  describe("medium file", function()
    local file_text
    local subset1
    local subset2
    local subset3
    local subset4

    before_each(function()
      file_text = direct.to_text(FIXTURES .. "medium-file.odt")
      local lines = vim.split(file_text, "\n")

      local function slice(tbl, first, last)
        local result = {}
        for i = first, last do
          result[#result + 1] = tbl[i]
        end
        return result
      end

      subset1 = table.concat(slice(lines, 1, 1), "\n")
      subset2 = table.concat(slice(lines, 4, 4), "\n")
      subset3 = table.concat(slice(lines, 10, 15), "\n")
      subset4 = table.concat(slice(lines, 23, 29), "\n")
    end)

    it("parses complex titles correctly", function()
      assert.equals(
        "Expro idea on the counter-symmetric negotiation of hypothetical umbrellas – procedural management in intermittently reversible and diagonally flavored environments",
        subset1
      )
    end)

    it("parses long paragraphs correctly", function()
      assert.equals(
        "Confusion is inevitable. The global committee of miscellaneous intentions has either misplaced the instructions or decided that instructions are optional. Consequently, we must prepare to operate within circumstances that fluctuate between sideways mornings and triangular afternoons, frequently disturbing our supplies of ordinary objects such as chairs, calendars, and moderately enthusiastic sandwiches. Fortunately, several inanimate processes demonstrate how such uncertainty can be enthusiastically ignored.",
        subset2
      )
    end)

    it("parses multiline text correctly", function()
      assert.equals(
        [[Project Title
The influence of speculative budgeting on the adaptive organization of improbable systems

Abstract

This project addresses a fundamental puzzle in theoretical miscellaneous studies: how the allocation of imaginary resources—particularly enthusiasm, paperwork, and partially understood diagrams—shapes the adaptability of complex but poorly defined systems.]],
        subset3
      )
    end)

    it("parses lines with bold-face correctly", function()
      assert.equals(
        [[Research goals

1. Quantify large-scale patterns in the distribution of hypothetical resources across irregularly scheduled environments.

2. Identify procedural and philosophical mechanisms responsible for the emergence of organized disorder.

3. Evaluate the consequences and trade-offs of allocating attention, paperwork, and speculative diagrams under conditions of moderate uncertainty.]],
        subset4
      )
    end)
  end)
end)
