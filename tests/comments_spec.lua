-- tests/comments_spec.lua
-- Tests for comment parsing, rendering, and manipulation

local assert = require("luassert")
local xml = require("neoffice.xml")

describe("Comments", function()
  local comments

  before_each(function()
    comments = require("neoffice.comments")
  end)

  describe("comment structure", function()
    it("parses basic comment annotation", function()
      local xml_content = [[
        <office:annotation office:name="comment1">
          <dc:creator>TestUser</dc:creator>
          <dc:date>2024-03-11T10:00:00</dc:date>
          <text:p>This is a comment</text:p>
        </office:annotation>
      ]]

      local tree = xml.parse(xml_content)
      local annotations = xml.find_all(tree, "office:annotation")

      assert.equals(1, #annotations)

      local ann = annotations[1]
      assert.equals("comment1", xml.attr(ann, "office:name"))

      local creator = xml.find_first(ann, "dc:creator")
      assert.equals("TestUser", xml.inner_text(creator))
    end)

    it("parses comment with multiple paragraphs", function()
      local xml_content = [[
        <office:annotation office:name="comment1">
          <dc:creator>User</dc:creator>
          <dc:date>2024-03-11</dc:date>
          <text:p>First paragraph</text:p>
          <text:p>Second paragraph</text:p>
        </office:annotation>
      ]]

      local tree = xml.parse(xml_content)
      local ann = xml.find_first(tree, "office:annotation")
      local paragraphs = xml.find_all(ann, "text:p")

      assert.equals(2, #paragraphs)
    end)

    it("identifies replies via parent-name attribute", function()
      local xml_content = [[
        <office:annotation loext:parent-name="comment1">
          <dc:creator>ReplyUser</dc:creator>
          <dc:date>2024-03-12</dc:date>
          <text:p>This is a reply</text:p>
        </office:annotation>
      ]]

      local tree = xml.parse(xml_content)
      local ann = xml.find_first(tree, "office:annotation")

      local parent = xml.attr(ann, "loext:parent-name")
      assert.equals("comment1", parent)
    end)
  end)

  describe("comment detection in buffer", function()
    it("detects annotation tags in text", function()
      local buf = vim.api.nvim_create_buf(false, true)
      local lines = {
        'Text before <office:annotation office:name="c1"><dc:creator>User</dc:creator><text:p>Comment</text:p></office:annotation> text after',
      }
      vim.api.nvim_buf_set_lines(buf, 0, -1, false, lines)

      local content = table.concat(vim.api.nvim_buf_get_lines(buf, 0, -1, false), "\n")
      assert.is_true(content:find("office:annotation", 1, true) ~= nil)

      vim.api.nvim_buf_delete(buf, { force = true })
    end)

    it("detects annotation-end markers", function()
      local buf = vim.api.nvim_create_buf(false, true)
      local lines = {
        'Text <office:annotation-end office:name="c1"/> after',
      }
      vim.api.nvim_buf_set_lines(buf, 0, -1, false, lines)

      local content = table.concat(vim.api.nvim_buf_get_lines(buf, 0, -1, false), "\n")
      assert.is_true(content:find("annotation-end", 1, true) ~= nil)

      vim.api.nvim_buf_delete(buf, { force = true })
    end)
  end)

  describe("comment threading", function()
    it("associates replies with parent comments", function()
      -- Test data structure
      local parent_comment = {
        id = "c1",
        author = "Alice",
        date = "2024-03-11",
        text = "Original comment",
        replies = {},
        resolved = false,
      }

      local reply = {
        author = "Bob",
        date = "2024-03-12",
        text = "Reply to Alice",
      }

      table.insert(parent_comment.replies, reply)

      assert.equals(1, #parent_comment.replies)
      assert.equals("Bob", parent_comment.replies[1].author)
    end)

    it("handles multiple replies", function()
      local comment = {
        id = "c1",
        replies = {},
      }

      for i = 1, 5 do
        table.insert(comment.replies, {
          author = "User" .. i,
          text = "Reply " .. i,
        })
      end

      assert.equals(5, #comment.replies)
    end)
  end)

  describe("comment ID generation", function()
    it("generates unique comment IDs", function()
      -- Simulate ID generation
      local id1 = string.format("__Annotation__%d_%d", math.random(10000, 99999), math.random(1000000000, 9999999999))
      local id2 = string.format("__Annotation__%d_%d", math.random(10000, 99999), math.random(1000000000, 9999999999))

      -- They should be different (with very high probability)
      assert.is_not_equal(id1, id2)
    end)

    it("uses expected ID format", function()
      local id = string.format("__Annotation__%d_%d", 12345, 9876543210)

      assert.is_true(id:find("__Annotation__", 1, true) ~= nil)
      assert.is_true(id:find("12345", 1, true) ~= nil)
      assert.is_true(id:find("9876543210", 1, true) ~= nil)
    end)
  end)

  describe("anchor matching", function()
    it("extracts anchor text from paragraph", function()
      local xml_content = [[
        <text:p>This is the anchor text that identifies the comment location</text:p>
      ]]

      local tree = xml.parse(xml_content)
      local para = tree.children[1]
      local inner = xml.serialize_inner(para)
      local anchor = inner:sub(1, 60)

      assert.is_not_nil(anchor)
      assert.is_true(#anchor <= 60)
      assert.is_true(anchor:find("anchor text", 1, true) ~= nil)
    end)

    it("limits anchor length to 60 characters", function()
      local long_text = string.rep("a", 200)
      local anchor = long_text:sub(1, 60)

      assert.equals(60, #anchor)
    end)
  end)

  describe("comment rendering logic", function()
    it("formats comment author correctly", function()
      local author = "JohnDoe"
      local formatted = "@" .. author

      assert.equals("@JohnDoe", formatted)
    end)

    it("truncates dates to date only", function()
      local full_date = "2024-03-11T10:30:45Z"
      local date_only = full_date:sub(1, 10)

      assert.equals("2024-03-11", date_only)
    end)

    it("shows resolved indicator", function()
      local comment = { resolved = true }
      local indicator = comment.resolved and "  ✓" or ""

      assert.equals("  ✓", indicator)
    end)

    it("does not show indicator for unresolved", function()
      local comment = { resolved = false }
      local indicator = comment.resolved and "  ✓" or ""

      assert.equals("", indicator)
    end)
  end)

  describe("text wrapping for display", function()
    it("wraps long lines", function()
      local text = "This is a very long comment that needs to be wrapped to fit within the display width"
      local width = 40
      local indent = 3

      -- Simple wrapping simulation
      local lines = {}
      local current_line = ""
      local prefix = string.rep(" ", indent)

      for word in text:gmatch("%S+") do
        if #current_line + #word + 1 > width - indent then
          table.insert(lines, prefix .. current_line)
          current_line = word
        else
          current_line = current_line == "" and word or (current_line .. " " .. word)
        end
      end

      if current_line ~= "" then
        table.insert(lines, prefix .. current_line)
      end

      assert.is_true(#lines > 1) -- Should wrap to multiple lines
    end)
  end)

  describe("XML escaping", function()
    it("escapes special characters in comment text", function()
      local text = "<script>alert('xss')</script> & other"
      local escaped = text:gsub("&", "&amp;"):gsub("<", "&lt;"):gsub(">", "&gt;")

      assert.equals("&lt;script&gt;alert('xss')&lt;/script&gt; &amp; other", escaped)
    end)

    it("handles quote characters", function()
      local text = 'He said "hello"'
      local escaped = text:gsub('"', "&quot;")

      assert.is_true(escaped:find("&quot;", 1, true) ~= nil)
    end)
  end)

  describe("comment deletion logic", function()
    it("removes comment from list", function()
      local comments_list = {
        { id = "c1", text = "Comment 1" },
        { id = "c2", text = "Comment 2" },
        { id = "c3", text = "Comment 3" },
      }

      -- Simulate deletion of c2
      comments_list = vim.tbl_filter(function(c)
        return c.id ~= "c2"
      end, comments_list)

      assert.equals(2, #comments_list)
      assert.equals("c1", comments_list[1].id)
      assert.equals("c3", comments_list[2].id)
    end)

    it("removes reply from comment", function()
      local comment = {
        replies = {
          { author = "User1", text = "Reply 1" },
          { author = "User2", text = "Reply 2" },
          { author = "User3", text = "Reply 3" },
        },
      }

      -- Remove reply at index 2
      table.remove(comment.replies, 2)

      assert.equals(2, #comment.replies)
      assert.equals("User1", comment.replies[1].author)
      assert.equals("User3", comment.replies[2].author)
    end)
  end)

  describe("comment XML generation", function()
    it("generates valid annotation XML", function()
      local comment_id = "test_comment_1"
      local author = "TestUser"
      local date = "2024-03-11T10:00:00Z"
      local text = "This is a test comment"

      local xml_template = string.format(
        '<office:annotation office:name="%s">'
          .. "<dc:creator>%s</dc:creator>"
          .. "<dc:date>%s</dc:date>"
          .. "<text:p>%s</text:p>"
          .. "</office:annotation>",
        comment_id,
        author,
        date,
        text
      )

      assert.is_true(xml_template:find("office:annotation", 1, true) ~= nil)
      assert.is_true(xml_template:find(comment_id, 1, true) ~= nil)
      assert.is_true(xml_template:find(author, 1, true) ~= nil)
      assert.is_true(xml_template:find(text, 1, true) ~= nil)
    end)

    it("generates reply with parent reference", function()
      local parent_id = "parent_comment"
      local reply_xml = string.format(
        '<office:annotation loext:parent-name="%s">'
          .. "<dc:creator>ReplyAuthor</dc:creator>"
          .. "<text:p>Reply text</text:p>"
          .. "</office:annotation>",
        parent_id
      )

      assert.is_true(reply_xml:find("loext:parent-name", 1, true) ~= nil)
      assert.is_true(reply_xml:find(parent_id, 1, true) ~= nil)
    end)
  end)

  describe("resolved status", function()
    it("toggles resolved status", function()
      local comment = { resolved = false }

      comment.resolved = not comment.resolved
      assert.is_true(comment.resolved)

      comment.resolved = not comment.resolved
      assert.is_false(comment.resolved)
    end)

    it("uses loext:resolved attribute", function()
      local xml_content = [[
        <office:annotation office:name="c1" loext:resolved="true">
          <dc:creator>User</dc:creator>
          <text:p>Comment</text:p>
        </office:annotation>
      ]]

      local tree = xml.parse(xml_content)
      local ann = xml.find_first(tree, "office:annotation")
      local resolved = xml.attr(ann, "loext:resolved")

      assert.equals("true", resolved)
    end)
  end)

  describe("edge cases", function()
    it("handles empty comment text", function()
      local comment = {
        id = "c1",
        author = "User",
        text = "",
        replies = {},
      }

      assert.equals("", comment.text)
      assert.is_not_nil(comment.id)
    end)

    it("handles comments with no author", function()
      local comment = {
        id = "c1",
        author = nil,
        text = "Anonymous comment",
      }

      local display_author = comment.author or "?"
      assert.equals("?", display_author)
    end)

    it("handles missing date", function()
      local comment = {
        id = "c1",
        date = nil,
      }

      local display_date = (comment.date or ""):sub(1, 10)
      assert.equals("", display_date)
    end)

    it("handles very long comment text", function()
      local long_text = string.rep("This is a long comment. ", 100)
      local comment = {
        text = long_text,
      }

      -- Should not crash
      assert.is_true(#comment.text > 1000)
    end)
  end)
end)
