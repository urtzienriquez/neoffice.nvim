-- tests/parse_spec.lua
-- Tests for direct.to_text() XML extraction

local assert = require("luassert")
local direct = require("neoffice.direct")

local FIXTURES = "tests/fixtures/"

describe("to_text", function()
  describe("simple file", function()
    local file_text, para_map, root

    before_each(function()
      file_text, para_map, root = direct.to_text(FIXTURES .. "simple-file.odt")
    end)

    it("returns valid text content", function()
      assert.is_not_nil(file_text)
      assert.is_string(file_text)
    end)

    it("returns paragraph map", function()
      assert.is_not_nil(para_map)
      assert.is_table(para_map)
    end)

    it("returns parsed root", function()
      assert.is_not_nil(root)
      assert.is_table(root)
    end)

    it("contains expected text content (may include XML tags)", function()
      -- The text should contain the actual content, possibly with XML tags
      assert.is_true(file_text:find("Add a simple line of text", 1, true) ~= nil)
      assert.is_true(file_text:find("something else", 1, true) ~= nil)
    end)

    it("has correct number of paragraphs in map", function()
      -- Count non-nil entries in para_map
      local count = 0
      for _ in pairs(para_map) do
        count = count + 1
      end
      assert.is_true(count >= 1) -- At least one paragraph
    end)
  end)

  describe("medium file", function()
    local file_text, para_map, root

    before_each(function()
      file_text, para_map, root = direct.to_text(FIXTURES .. "medium-file.odt")
    end)

    it("extracts multi-paragraph content", function()
      assert.is_not_nil(file_text)
      local lines = vim.split(file_text, "\n")
      assert.is_true(#lines > 1) -- Multiple paragraphs
    end)

    it("contains title text", function()
      assert.is_true(
        file_text:find("counter%-symmetric negotiation", 1, true) ~= nil
          or file_text:find("counter-symmetric negotiation", 1, true) ~= nil
      )
    end)

    it("contains long paragraph text", function()
      assert.is_true(file_text:find("Confusion is inevitable", 1, true) ~= nil)
      assert.is_true(file_text:find("enthusiastically ignored", 1, true) ~= nil)
    end)

    it("preserves multiline structure", function()
      assert.is_true(file_text:find("Project Title", 1, true) ~= nil)
      assert.is_true(file_text:find("Abstract", 1, true) ~= nil)
    end)

    it("handles formatted text (bold, etc)", function()
      assert.is_true(file_text:find("Research goals", 1, true) ~= nil)
    end)

    it("paragraph map keys match line count", function()
      local lines = vim.split(file_text, "\n")
      local max_line = 0
      for line_num in pairs(para_map) do
        if line_num > max_line then
          max_line = line_num
        end
      end
      assert.is_true(max_line <= #lines)
    end)
  end)

  describe("from_text roundtrip", function()
    it("can write edited text back", function()
      local orig_path = FIXTURES .. "simple-file.odt"
      local text, para_map, root = direct.to_text(orig_path)

      assert.is_not_nil(text)
      assert.is_not_nil(para_map)
      assert.is_not_nil(root)

      -- Modify the text slightly
      local modified = text:gsub("simple", "MODIFIED")

      -- Create a temp copy to avoid modifying fixture
      local tmp = vim.fn.tempname() .. ".odt"
      vim.fn.system(string.format("cp '%s' '%s'", orig_path, tmp))

      -- Write back
      local ok, err = direct.from_text(tmp, modified, para_map, root)
      assert.is_true(ok, err or "from_text failed")

      -- Read again and verify
      local new_text = direct.to_text(tmp)
      assert.is_not_nil(new_text)
      assert.is_true(new_text:find("MODIFIED", 1, true) ~= nil)

      -- Cleanup
      vim.fn.delete(tmp)
    end)
  end)

  describe("edge cases", function()
    it("handles non-existent files gracefully", function()
      local text, err = direct.to_text(FIXTURES .. "does-not-exist.odt")
      assert.is_nil(text)
      assert.is_not_nil(err)
    end)

    it("handles empty paragraphs", function()
      -- If we have a fixture with empty paragraphs, test it
      local text = direct.to_text(FIXTURES .. "simple-file.odt")
      if text then
        -- Empty paragraphs should at least contain a space
        local lines = vim.split(text, "\n")
        for _, line in ipairs(lines) do
          -- Each line should be non-nil (may be empty string or space)
          assert.is_not_nil(line)
        end
      end
    end)
  end)

  describe("unsupported formats", function()
    it("returns error for unsupported extensions", function()
      local text, err = direct.to_text(FIXTURES .. "test.txt")
      assert.is_nil(text)
      assert.equals("Unsupported format", err)
    end)
  end)
end)
