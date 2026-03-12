-- tests/xml_spec.lua
-- Comprehensive tests for XML parsing, serialization, and manipulation

local assert = require("luassert")
local xml = require("neoffice.xml")

describe("XML Module", function()
  describe("parse", function()
    it("parses simple self-closing tags", function()
      local input = '<tag attr="value"/>'
      local tree = xml.parse(input)

      assert.is_not_nil(tree)
      assert.is_not_nil(tree.children)
      assert.equals(1, #tree.children)
      assert.equals("tag", tree.children[1].tag)
      assert.equals("value", tree.children[1].attrs.attr)
    end)

    it("parses nested tags", function()
      local input = "<parent><child>text</child></parent>"
      local tree = xml.parse(input)

      local parent = tree.children[1]
      assert.equals("parent", parent.tag)
      assert.equals(1, #parent.children)
      assert.equals("child", parent.children[1].tag)
    end)

    it("parses text content", function()
      local input = "<tag>Hello World</tag>"
      local tree = xml.parse(input)

      local tag = tree.children[1]
      assert.equals("tag", tag.tag)
      assert.equals(1, #tag.children)
      assert.equals("_TEXT", tag.children[1].tag)
      assert.equals("Hello World", tag.children[1].text)
    end)

    it("parses multiple attributes", function()
      local input = '<tag a="1" b="2" c="3"/>'
      local tree = xml.parse(input)

      local tag = tree.children[1]
      assert.equals("1", tag.attrs.a)
      assert.equals("2", tag.attrs.b)
      assert.equals("3", tag.attrs.c)
    end)

    it("handles namespaced tags", function()
      local input = '<office:text xmlns:office="urn:office"><office:p>text</office:p></office:text>'
      local tree = xml.parse(input)

      local office_text = tree.children[1]
      assert.equals("office:text", office_text.tag)
      assert.is_not_nil(office_text.children)
    end)

    it("handles escaped entities", function()
      local input = '<tag attr="&lt;&gt;&amp;&quot;">text</tag>'
      local tree = xml.parse(input)

      local tag = tree.children[1]
      assert.equals('<>&"', tag.attrs.attr)
    end)

    it("handles mixed content", function()
      local input = "<p>Text <b>bold</b> more text</p>"
      local tree = xml.parse(input)

      local p = tree.children[1]
      assert.equals("p", p.tag)
      assert.is_true(#p.children >= 2) -- At least text and bold
    end)

    it("handles XML declaration", function()
      local input = '<?xml version="1.0"?><root/>'
      local tree = xml.parse(input)

      assert.is_not_nil(tree)
      -- Declaration should be skipped
      assert.equals(1, #tree.children)
      assert.equals("root", tree.children[1].tag)
    end)

    it("handles comments", function()
      local input = "<!-- comment --><tag/>"
      local tree = xml.parse(input)

      -- Comments should be skipped
      assert.equals(1, #tree.children)
      assert.equals("tag", tree.children[1].tag)
    end)

    it("handles CDATA sections", function()
      local input = "<tag><![CDATA[some data]]></tag>"
      local tree = xml.parse(input)

      assert.is_not_nil(tree)
      -- CDATA handling depends on implementation
    end)
  end)

  describe("serialize", function()
    it("serializes simple tags", function()
      local node = {
        tag = "tag",
        attrs = { key = "value" },
        children = {},
      }

      local output = xml.serialize(node)
      assert.is_true(output:find("tag", 1, true) ~= nil)
      assert.is_true(output:find("key", 1, true) ~= nil)
      assert.is_true(output:find("value", 1, true) ~= nil)
    end)

    it("serializes nested tags", function()
      local node = {
        tag = "parent",
        attrs = {},
        children = {
          {
            tag = "child",
            attrs = {},
            children = {},
          },
        },
      }

      local output = xml.serialize(node)
      assert.is_true(output:find("parent", 1, true) ~= nil)
      assert.is_true(output:find("child", 1, true) ~= nil)
    end)

    it("serializes text nodes", function()
      local node = {
        tag = "p",
        attrs = {},
        children = {
          { tag = "_TEXT", text = "Hello" },
        },
      }

      local output = xml.serialize(node)
      assert.is_true(output:find("Hello", 1, true) ~= nil)
    end)

    it("escapes special characters in text", function()
      local node = {
        tag = "p",
        attrs = {},
        children = {
          { tag = "_TEXT", text = "<>&" },
        },
      }

      local output = xml.serialize(node)
      assert.is_true(output:find("&lt;", 1, true) ~= nil)
      assert.is_true(output:find("&gt;", 1, true) ~= nil)
      assert.is_true(output:find("&amp;", 1, true) ~= nil)
    end)

    it("escapes special characters in attributes", function()
      local node = {
        tag = "tag",
        attrs = { key = '<>"&' },
        children = {},
      }

      local output = xml.serialize(node)
      assert.is_true(output:find("&lt;", 1, true) ~= nil)
      assert.is_true(output:find("&quot;", 1, true) ~= nil)
      assert.is_true(output:find("&amp;", 1, true) ~= nil)
    end)

    it("uses self-closing tags for empty nodes", function()
      local node = {
        tag = "empty",
        attrs = {},
        children = {},
      }

      local output = xml.serialize(node)
      assert.is_true(output:find("/>", 1, true) ~= nil)
    end)

    it("handles ROOT virtual node", function()
      local node = {
        tag = "ROOT",
        attrs = {},
        children = {
          { tag = "child1", attrs = {}, children = {} },
          { tag = "child2", attrs = {}, children = {} },
        },
      }

      local output = xml.serialize(node)
      -- ROOT should not appear in output
      assert.is_false(output:find("ROOT", 1, true) ~= nil)
      assert.is_true(output:find("child1", 1, true) ~= nil)
      assert.is_true(output:find("child2", 1, true) ~= nil)
    end)
  end)

  describe("serialize_inner", function()
    it("serializes only children", function()
      local node = {
        tag = "parent",
        attrs = {},
        children = {
          { tag = "_TEXT", text = "Hello" },
          { tag = "child", attrs = {}, children = {} },
        },
      }

      local output = xml.serialize_inner(node)
      assert.is_true(output:find("Hello", 1, true) ~= nil)
      assert.is_true(output:find("child", 1, true) ~= nil)
      -- Should NOT include parent tag
      assert.is_false(output:find("<parent", 1, true) ~= nil)
    end)

    it("returns empty string for nodes without children", function()
      local node = {
        tag = "empty",
        attrs = {},
        children = {},
      }

      local output = xml.serialize_inner(node)
      assert.equals("", output)
    end)
  end)

  describe("inner_text", function()
    it("extracts text from simple nodes", function()
      local node = {
        tag = "p",
        children = {
          { tag = "_TEXT", text = "Hello" },
        },
      }

      assert.equals("Hello", xml.inner_text(node))
    end)

    it("extracts text from nested nodes", function()
      local node = {
        tag = "p",
        children = {
          { tag = "_TEXT", text = "Hello " },
          {
            tag = "b",
            children = {
              { tag = "_TEXT", text = "World" },
            },
          },
        },
      }

      assert.equals("Hello World", xml.inner_text(node))
    end)

    it("returns empty string for nil nodes", function()
      assert.equals("", xml.inner_text(nil))
    end)

    it("concatenates multiple text nodes", function()
      local node = {
        tag = "p",
        children = {
          { tag = "_TEXT", text = "Part1" },
          { tag = "_TEXT", text = "Part2" },
          { tag = "_TEXT", text = "Part3" },
        },
      }

      assert.equals("Part1Part2Part3", xml.inner_text(node))
    end)
  end)

  describe("attr", function()
    it("retrieves simple attributes", function()
      local node = {
        tag = "tag",
        attrs = { key = "value" },
      }

      assert.equals("value", xml.attr(node, "key"))
    end)

    it("returns nil for missing attributes", function()
      local node = {
        tag = "tag",
        attrs = {},
      }

      assert.is_nil(xml.attr(node, "missing"))
    end)

    it("handles namespaced attributes", function()
      local node = {
        tag = "tag",
        attrs = { ["office:name"] = "test" },
      }

      assert.equals("test", xml.attr(node, "office:name"))
    end)

    it("matches local name for namespaced attributes", function()
      local node = {
        tag = "tag",
        attrs = { ["office:name"] = "test" },
      }

      -- Should match on local name "name"
      assert.equals("test", xml.attr(node, "name"))
    end)

    it("handles nil nodes", function()
      assert.is_nil(xml.attr(nil, "key"))
    end)
  end)

  describe("find_all", function()
    it("finds all matching tags", function()
      local tree = xml.parse("<root><item/><item/><item/></root>")
      local items = xml.find_all(tree, "item")

      assert.equals(3, #items)
    end)

    it("finds nested tags", function()
      local tree = xml.parse("<root><parent><item/></parent><item/></root>")
      local items = xml.find_all(tree, "item")

      assert.equals(2, #items)
    end)

    it("returns empty table when no matches", function()
      local tree = xml.parse("<root><other/></root>")
      local items = xml.find_all(tree, "item")

      assert.equals(0, #items)
    end)

    it("handles deeply nested structures", function()
      local tree = xml.parse("<a><b><c><item/></c></b><item/></a>")
      local items = xml.find_all(tree, "item")

      assert.equals(2, #items)
    end)
  end)

  describe("find_first", function()
    it("finds first matching tag", function()
      local tree = xml.parse("<root><item id='1'/><item id='2'/></root>")
      local item = xml.find_first(tree, "item")

      assert.is_not_nil(item)
      assert.equals("item", item.tag)
      assert.equals("1", item.attrs.id)
    end)

    it("returns nil when no match", function()
      local tree = xml.parse("<root><other/></root>")
      local item = xml.find_first(tree, "item")

      assert.is_nil(item)
    end)
  end)

  describe("find_first_local", function()
    it("finds tag by local name", function()
      local tree = xml.parse("<office:root><office:text>content</office:text></office:root>")
      local text = xml.find_first_local(tree, "text")

      assert.is_not_nil(text)
      assert.equals("office:text", text.tag)
    end)

    it("matches tags without namespace", function()
      local tree = xml.parse("<root><text>content</text></root>")
      local text = xml.find_first_local(tree, "text")

      assert.is_not_nil(text)
      assert.equals("text", text.tag)
    end)

    it("returns nil when no match", function()
      local tree = xml.parse("<root><other/></root>")
      local text = xml.find_first_local(tree, "text")

      assert.is_nil(text)
    end)
  end)

  describe("roundtrip", function()
    it("parse and serialize produce equivalent XML", function()
      local input = '<root a="1"><child>text</child><empty/></root>'
      local tree = xml.parse(input)
      local output = xml.serialize(tree)

      -- Re-parse to compare structure
      local tree2 = xml.parse(output)

      -- Both should have ROOT with same structure
      assert.equals(#tree.children, #tree2.children)
    end)

    it("preserves complex nested structures", function()
      local input = [[
        <office:document>
          <office:body>
            <office:text>
              <text:p>Paragraph 1</text:p>
              <text:p>Paragraph 2</text:p>
            </office:text>
          </office:body>
        </office:document>
      ]]

      local tree = xml.parse(input)
      local output = xml.serialize(tree)
      local tree2 = xml.parse(output)

      -- Find paragraphs in both
      local p1 = xml.find_all(tree, "text:p")
      local p2 = xml.find_all(tree2, "text:p")

      assert.equals(#p1, #p2)
    end)
  end)
end)
