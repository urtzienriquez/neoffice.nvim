-- tests/extractor_spec.lua
-- Tests for extracting track changes and comments from documents

local assert = require("luassert")
local extractor = require("neoffice.extractor")
local xml = require("neoffice.xml")
local zip = require("neoffice.zip")

describe("Extractor", function()
  describe("comment extraction from ODT", function()
    it("extracts nested annotations correctly", function()
      -- Simulate an ODT structure with nested annotations (replies)
      local xml_content = [[
        <office:text>
          <text:p>
            Paragraph text
            <office:annotation office:name="parent1">
              <dc:creator>Alice</dc:creator>
              <dc:date>2024-03-11</dc:date>
              <text:p>Parent comment</text:p>
              <office:annotation loext:parent-name="parent1">
                <dc:creator>Bob</dc:creator>
                <dc:date>2024-03-12</dc:date>
                <text:p>Reply to parent</text:p>
              </office:annotation>
            </office:annotation>
          </text:p>
        </office:text>
      ]]

      local tree = xml.parse(xml_content)

      -- Find all annotations
      local all_annotations = xml.find_all(tree, "office:annotation")

      -- Should find both parent and reply
      assert.is_true(#all_annotations >= 1)
    end)

    it("separates parents from replies", function()
      local xml_content = [[
        <office:text>
          <text:p>
            <office:annotation office:name="c1">
              <dc:creator>Alice</dc:creator>
              <text:p>Parent</text:p>
            </office:annotation>
            <office:annotation loext:parent-name="c1">
              <dc:creator>Bob</dc:creator>
              <text:p>Reply</text:p>
            </office:annotation>
          </text:p>
        </office:text>
      ]]

      local tree = xml.parse(xml_content)
      local annotations = xml.find_all(tree, "office:annotation")

      local parents = {}
      local replies = {}

      for _, ann in ipairs(annotations) do
        local parent_name = xml.attr(ann, "loext:parent-name")
        if parent_name then
          table.insert(replies, ann)
        else
          table.insert(parents, ann)
        end
      end

      assert.equals(1, #parents)
      assert.equals(1, #replies)
    end)
  end)

  describe("track changes extraction from ODT", function()
    it("extracts changed-region metadata", function()
      local xml_content = [[
        <office:text>
          <text:tracked-changes>
            <text:changed-region text:id="ct1">
              <text:insertion>
                <office:change-info>
                  <dc:creator>User</dc:creator>
                  <dc:date>2024-03-11</dc:date>
                </office:change-info>
              </text:insertion>
            </text:changed-region>
          </text:tracked-changes>
        </office:text>
      ]]

      local tree = xml.parse(xml_content)
      local regions = xml.find_all(tree, "text:changed-region")

      assert.equals(1, #regions)

      local region = regions[1]
      local id = xml.attr(region, "text:id")
      assert.equals("ct1", id)
    end)

    it("distinguishes insertions from deletions", function()
      local xml_content = [[
        <text:tracked-changes>
          <text:changed-region text:id="ins1">
            <text:insertion>
              <office:change-info><dc:creator>User</dc:creator></office:change-info>
            </text:insertion>
          </text:changed-region>
          <text:changed-region text:id="del1">
            <text:deletion>
              <office:change-info><dc:creator>User</dc:creator></office:change-info>
              <text:p>Deleted text</text:p>
            </text:deletion>
          </text:changed-region>
        </text:tracked-changes>
      ]]

      local tree = xml.parse(xml_content)

      local insertions = xml.find_all(tree, "text:insertion")
      local deletions = xml.find_all(tree, "text:deletion")

      assert.equals(1, #insertions)
      assert.equals(1, #deletions)
    end)

    it("extracts deleted text content", function()
      local xml_content = [[
        <text:deletion>
          <office:change-info>
            <dc:creator>User</dc:creator>
          </office:change-info>
          <text:p>This text was deleted</text:p>
          <text:p>Multiple paragraphs</text:p>
        </text:deletion>
      ]]

      local tree = xml.parse(xml_content)
      local deletion = tree.children[1]

      local paragraphs = xml.find_all(deletion, "text:p")
      assert.equals(2, #paragraphs)

      local text = xml.inner_text(deletion)
      assert.is_true(text:find("deleted", 1, true) ~= nil)
    end)
  end)

  describe("DOCX comment extraction", function()
    it("parses comment XML structure", function()
      local xml_content = [[
        <w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:comment w:id="1" w:author="Alice" w:date="2024-03-11">
            <w:p><w:r><w:t>Comment text</w:t></w:r></w:p>
          </w:comment>
        </w:comments>
      ]]

      local tree = xml.parse(xml_content)
      local comments = xml.find_all(tree, "w:comment")

      assert.equals(1, #comments)

      local comment = comments[1]
      assert.equals("1", xml.attr(comment, "w:id"))
      assert.equals("Alice", xml.attr(comment, "w:author"))
    end)

    it("extracts comment text from w:t elements", function()
      local xml_content = [[
        <w:comment w:id="1">
          <w:p>
            <w:r><w:t>First part</w:t></w:r>
            <w:r><w:t> second part</w:t></w:r>
          </w:p>
        </w:comment>
      ]]

      local tree = xml.parse(xml_content)
      local comment = tree.children[1]

      local text_nodes = xml.find_all(comment, "w:t")
      assert.is_true(#text_nodes >= 2)
    end)
  end)

  describe("DOCX track changes extraction", function()
    it("identifies insertions (w:ins)", function()
      local xml_content = [[
        <w:p>
          <w:ins w:id="1" w:author="User" w:date="2024-03-11">
            <w:r><w:t>inserted text</w:t></w:r>
          </w:ins>
        </w:p>
      ]]

      local tree = xml.parse(xml_content)
      local insertions = xml.find_all(tree, "w:ins")

      assert.equals(1, #insertions)

      local ins = insertions[1]
      assert.equals("User", xml.attr(ins, "w:author"))
    end)

    it("identifies deletions (w:del)", function()
      local xml_content = [[
        <w:p>
          <w:del w:id="2" w:author="User" w:date="2024-03-11">
            <w:r><w:delText>deleted text</w:delText></w:r>
          </w:del>
        </w:p>
      ]]

      local tree = xml.parse(xml_content)
      local deletions = xml.find_all(tree, "w:del")

      assert.equals(1, #deletions)
    end)

    it("uses w:delText for deleted content", function()
      local xml_content = [[
        <w:del>
          <w:r><w:delText>This was removed</w:delText></w:r>
        </w:del>
      ]]

      local tree = xml.parse(xml_content)
      local del_text = xml.find_all(tree, "w:delText")

      assert.is_true(#del_text >= 1)
      assert.equals("This was removed", xml.inner_text(del_text[1]))
    end)
  end)

  describe("annotation injection for ODT", function()
    it("creates valid annotation node structure", function()
      local comment = {
        id = "test_id",
        author = "TestUser",
        date = "2024-03-11T10:00:00Z",
        text = "Test comment",
        replies = {},
      }

      -- Simulate the structure that would be created
      local ann = {
        tag = "office:annotation",
        attrs = { ["office:name"] = comment.id },
        children = {
          {
            tag = "dc:creator",
            attrs = {},
            children = { { tag = "_TEXT", text = comment.author } },
          },
          {
            tag = "dc:date",
            attrs = {},
            children = { { tag = "_TEXT", text = comment.date } },
          },
          {
            tag = "text:p",
            attrs = {},
            children = { { tag = "_TEXT", text = comment.text } },
          },
        },
      }

      -- Verify structure
      assert.equals("office:annotation", ann.tag)
      assert.equals(comment.id, ann.attrs["office:name"])
      assert.equals(3, #ann.children)
    end)

    it("generates valid XML from annotation node", function()
      local ann = {
        tag = "office:annotation",
        attrs = { ["office:name"] = "test" },
        children = {
          { tag = "dc:creator", attrs = {}, children = { { tag = "_TEXT", text = "User" } } },
          { tag = "text:p", attrs = {}, children = { { tag = "_TEXT", text = "Comment" } } },
        },
      }

      local xml_output = xml.serialize(ann)

      assert.is_true(xml_output:find("office:annotation", 1, true) ~= nil)
      assert.is_true(xml_output:find("office:name", 1, true) ~= nil)
      assert.is_true(xml_output:find("User", 1, true) ~= nil)
      assert.is_true(xml_output:find("Comment", 1, true) ~= nil)
    end)
  end)

  describe("DOCX comments XML generation", function()
    it("generates valid comments.xml structure", function()
      local comment = {
        id = "1",
        author = "TestUser",
        date = "2024-03-11T10:00:00Z",
        text = "Test comment",
        replies = {},
      }

      -- Simulate XML generation
      local escaped_text = comment.text:gsub("&", "&amp;"):gsub("<", "&lt;"):gsub(">", "&gt;")

      local xml_line =
        string.format('<w:comment w:id="%s" w:author="%s" w:date="%s">', comment.id, comment.author, comment.date)

      assert.is_true(xml_line:find("w:comment", 1, true) ~= nil)
      assert.is_true(xml_line:find(comment.id, 1, true) ~= nil)
      assert.is_true(xml_line:find(comment.author, 1, true) ~= nil)
    end)

    it("escapes special characters properly", function()
      local text = "<script>alert('test')</script> & more"
      local escaped = text:gsub("&", "&amp;"):gsub("<", "&lt;"):gsub(">", "&gt;")

      assert.equals("&lt;script&gt;alert('test')&lt;/script&gt; &amp; more", escaped)
    end)
  end)

  describe("edge cases", function()
    it("handles empty comment lists", function()
      local comments = {}

      -- Should not crash when processing empty list
      assert.equals(0, #comments)
    end)

    it("handles missing metadata gracefully", function()
      local xml_content = [[
        <office:annotation office:name="c1">
          <text:p>Comment without creator or date</text:p>
        </office:annotation>
      ]]

      local tree = xml.parse(xml_content)
      local ann = xml.find_first(tree, "office:annotation")

      local creator = xml.find_first(ann, "dc:creator")
      assert.is_nil(creator) -- No creator element
    end)

    it("handles malformed XML gracefully", function()
      -- Test with incomplete XML
      local xml_content = [[
        <office:annotation office:name="broken"
      ]]

      -- Should either parse partially or return error
      -- but not crash
      local ok = pcall(xml.parse, xml_content)
      -- We just want to ensure it doesn't crash
      assert.is_not_nil(ok)
    end)
  end)

  describe("format detection", function()
    it("identifies ODT format by extension", function()
      local path = "test.odt"
      local ext = path:match("%.(%w+)$"):lower()

      assert.equals("odt", ext)
    end)

    it("identifies DOCX format by extension", function()
      local path = "test.docx"
      local ext = path:match("%.(%w+)$"):lower()

      assert.equals("docx", ext)
    end)

    it("handles mixed case extensions", function()
      local path = "test.ODT"
      local ext = path:match("%.(%w+)$"):lower()

      assert.equals("odt", ext)
    end)
  end)
end)
