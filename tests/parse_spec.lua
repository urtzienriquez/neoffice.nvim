-- tests/parse_spec.lua

local assert = require("luassert")
local direct = require("neoffice.direct")

local FIXTURES = "tests/fixtures/"

describe("parse_text", function()
  local file_text

  before_each(function()
    file_text = direct.to_text(FIXTURES .. "simple-file.odt")
  end)

  it("returns the correct number of entries", function()
    assert.equals(
      "Add a simple line of text. And something else.",
      file_text
    )
  end)

end)
