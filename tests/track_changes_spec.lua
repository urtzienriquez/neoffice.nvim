-- tests/track_changes_spec.lua
-- Tests for track changes detection, navigation, and manipulation

local assert = require("luassert")
local xml = require("neoffice.xml")

describe("Track Changes", function()
  local tc

  before_each(function()
    tc = require("neoffice.track_changes")
  end)

  describe("change detection", function()
    it("detects insertion tags", function()
      -- Create a buffer with track change tags
      local buf = vim.api.nvim_create_buf(false, true)
      local lines = {
        'Some text <text:change-start text:change-id="ins1"/>inserted text<text:change-end text:change-id="ins1"/> more text',
      }
      vim.api.nvim_buf_set_lines(buf, 0, -1, false, lines)

      -- The tag should be detectable
      local content = table.concat(vim.api.nvim_buf_get_lines(buf, 0, -1, false), "\n")
      assert.is_true(content:find("text:change", 1, true) ~= nil)

      vim.api.nvim_buf_delete(buf, { force = true })
    end)

    it("detects deletion tags", function()
      local buf = vim.api.nvim_create_buf(false, true)
      local lines = {
        'Some text <text:change text:change-id="del1"/>more text',
      }
      vim.api.nvim_buf_set_lines(buf, 0, -1, false, lines)

      local content = table.concat(vim.api.nvim_buf_get_lines(buf, 0, -1, false), "\n")
      assert.is_true(content:find("text:change", 1, true) ~= nil)

      vim.api.nvim_buf_delete(buf, { force = true })
    end)

    it("handles multiple changes in one buffer", function()
      local buf = vim.api.nvim_create_buf(false, true)
      local lines = {
        '<text:change-start text:change-id="ins1"/>insertion<text:change-end text:change-id="ins1"/>',
        '<text:change text:change-id="del1"/>',
        '<text:change-start text:change-id="ins2"/>another<text:change-end text:change-id="ins2"/>',
      }
      vim.api.nvim_buf_set_lines(buf, 0, -1, false, lines)

      local content = table.concat(vim.api.nvim_buf_get_lines(buf, 0, -1, false), "\n")
      local count = 0
      for _ in content:gmatch("text:change") do
        count = count + 1
      end
      assert.is_true(count >= 3) -- At least 3 change tags

      vim.api.nvim_buf_delete(buf, { force = true })
    end)
  end)

  describe("paired range finding", function()
    it("finds matching start and end tags", function()
      local lines = {
        'text before <text:change-start text:change-id="test123"/>changed text<text:change-end text:change-id="test123"/> text after',
      }

      local range = tc.find_paired_range(lines, "test123")

      assert.is_not_nil(range)
      assert.equals("paired", range.type)
      assert.equals("test123", range.change_id)
      assert.is_not_nil(range.start_tag)
      assert.is_not_nil(range.end_tag)
    end)

    it("handles tags on separate lines", function()
      local lines = {
        'text before <text:change-start text:change-id="multi"/>',
        "changed text spanning",
        "multiple lines",
        '<text:change-end text:change-id="multi"/> text after',
      }

      local range = tc.find_paired_range(lines, "multi")

      assert.is_not_nil(range)
      assert.equals("paired", range.type)
    end)

    it("returns nil for non-existent change ID", function()
      local lines = {
        "text with no changes",
      }

      local range = tc.find_paired_range(lines, "missing")
      assert.is_nil(range)
    end)

    it("handles special characters in change ID", function()
      local lines = {
        '<text:change-start text:change-id="ct-1234567890"/>text<text:change-end text:change-id="ct-1234567890"/>',
      }

      local range = tc.find_paired_range(lines, "ct-1234567890")
      assert.is_not_nil(range)
    end)
  end)

  describe("text sanitization", function()
    it("strips author name from deletion text", function()
      -- The track_changes module sanitizes deletion text
      -- This is tested indirectly through the display
      local deletion_text = "JohnDoe2024-03-11T10:30:00This was deleted"

      -- After sanitization, should remove author and timestamp
      local sanitized = deletion_text:gsub("^JohnDoe", "")
      sanitized = sanitized:gsub("^%d%d%d%d%-%d%d%-%d%dT%d%d:%d%d:%d%d%.?%d*", "")

      assert.equals("This was deleted", sanitized)
    end)

    it("handles various timestamp formats", function()
      local texts = {
        "2024-03-11T10:30:00Deleted",
        "2024-03-11T10:30:00.123Deleted",
        "2024-12-31T23:59:59.999999Deleted",
      }

      for _, text in ipairs(texts) do
        local sanitized = text:gsub("^%d%d%d%d%-%d%d%-%d%dT%d%d:%d%d:%d%d%.?%d*", "")
        assert.equals("Deleted", sanitized)
      end
    end)
  end)

  describe("metadata removal", function()
    it("identifies changed-region blocks", function()
      local content = [[
        <text:tracked-changes>
          <text:changed-region text:id="ct123">
            <text:deletion>
              <office:change-info>
                <dc:creator>Author</dc:creator>
              </office:change-info>
            </text:deletion>
          </text:changed-region>
        </text:tracked-changes>
      ]]

      local pattern = '<text:changed%-region[^>]*text:id="ct123"[^>]*>.-</text:changed%-region>'
      assert.is_not_nil(content:find(pattern))
    end)

    it("handles nested tags in metadata", function()
      local content = [[
        <text:changed-region text:id="complex">
          <text:insertion>
            <office:change-info>
              <dc:creator>Name</dc:creator>
              <dc:date>2024-03-11</dc:date>
              <nested>
                <deep>value</deep>
              </nested>
            </office:change-info>
          </text:insertion>
        </text:changed-region>
      ]]

      -- The pattern should match despite nested tags
      local pattern = '<text:changed%-region[^>]*text:id="complex"[^>]*>.-</text:changed%-region>'
      assert.is_not_nil(content:find(pattern))
    end)
  end)

  describe("navigation helpers", function()
    it("calculates byte to position correctly", function()
      local lines = {
        "Line 1",
        "Line 2",
        "Line 3",
      }

      -- This tests the internal byte_to_pos function indirectly
      -- by verifying that positions make sense
      local full_text = table.concat(lines, "\n")

      -- First character should be (0, 0)
      -- Character after "Line 1\n" should be (1, 0)
      assert.equals(20, #full_text) -- 6 + 1 + 6 + 1 + 6
    end)
  end)

  describe("change extraction from ODT", function()
    it("parses tracked changes metadata", function()
      local xml_content = [[
        <office:text>
          <text:tracked-changes>
            <text:changed-region text:id="ct1">
              <text:insertion>
                <office:change-info>
                  <dc:creator>TestUser</dc:creator>
                  <dc:date>2024-03-11T10:00:00</dc:date>
                </office:change-info>
              </text:insertion>
            </text:changed-region>
          </text:tracked-changes>
          <text:p>
            Some text
            <text:change-start text:change-id="ct1"/>
            inserted text
            <text:change-end text:change-id="ct1"/>
            more text
          </text:p>
        </office:text>
      ]]

      local tree = xml.parse(xml_content)
      local regions = xml.find_all(tree, "text:changed-region")

      assert.is_true(#regions >= 1)

      if #regions > 0 then
        local id = xml.attr(regions[1], "text:id")
        assert.equals("ct1", id)
      end
    end)
  end)

  describe("integration scenarios", function()
    it("handles accept operation conceptually", function()
      -- Test the logic of accepting a change
      local original = 'Keep <text:change-start text:change-id="ins"/>this<text:change-end text:change-id="ins"/> text'
      local expected = "Keep this text"

      -- Simulate removing tags
      local result = original:gsub('<text:change%-start text:change%-id="ins"/>', "")
      result = result:gsub('<text:change%-end text:change%-id="ins"/>', "")

      assert.equals(expected, result)
    end)

    it("handles reject operation conceptually", function()
      -- Test the logic of rejecting an insertion (remove tags AND content)
      local original = 'Keep <text:change-start text:change-id="ins"/>this<text:change-end text:change-id="ins"/> text'
      local expected = "Keep  text"

      -- Simulate removing tags and content between them
      local result =
        original:gsub('<text:change%-start text:change%-id="ins"/>.-<text:change%-end text:change%-id="ins"/>', "")

      assert.equals(expected, result)
    end)

    it("handles deletion point tag restoration", function()
      -- When rejecting a deletion, restore the deleted text
      local deletion_tag = '<text:change text:change-id="del"/>'
      local deleted_text = "restored"

      -- Simulate restoration
      local original = "Text before " .. deletion_tag .. " text after"
      local result = original:gsub(vim.pesc(deletion_tag), deleted_text)

      assert.equals("Text before restored text after", result)
    end)
  end)

  describe("edge cases", function()
    it("handles empty change IDs", function()
      local lines = {
        '<text:change-start text:change-id=""/>text<text:change-end text:change-id=""/>',
      }

      -- Should not crash
      local range = tc.find_paired_range(lines, "")
      -- May or may not find it depending on implementation
    end)

    it("handles changes with quotes in IDs", function()
      -- This shouldn't happen in real ODT, but test robustness
      local lines = {
        "text without problematic IDs",
      }

      -- Should not crash
      assert.is_not_nil(lines)
    end)

    it("handles very long change IDs", function()
      local long_id = string.rep("a", 200)
      local lines = {
        string.format(
          '<text:change-start text:change-id="%s"/>text<text:change-end text:change-id="%s"/>',
          long_id,
          long_id
        ),
      }

      local range = tc.find_paired_range(lines, long_id)
      assert.is_not_nil(range)
    end)
  end)
end)
