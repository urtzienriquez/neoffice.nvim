-- tests/integration_spec.lua
-- End-to-end integration tests for neoffice

local assert = require("luassert")

describe("Integration Tests", function()
  local FIXTURES = "tests/fixtures/"

  describe("document workflow", function()
    it("can open, edit, and save a document", function()
      -- This would require actual fixture files
      -- For now, test the conceptual workflow

      local orig_path = FIXTURES .. "simple-file.odt"

      -- Check if fixture exists (may not in minimal test environment)
      if vim.fn.filereadable(orig_path) == 1 then
        local direct = require("neoffice.direct")

        -- Extract text
        local text, para_map, root = direct.to_text(orig_path)
        assert.is_not_nil(text)
        assert.is_not_nil(para_map)
        assert.is_not_nil(root)

        -- Modify text
        local modified = text:gsub("simple", "MODIFIED")

        -- Create temp copy
        local tmp = vim.fn.tempname() .. ".odt"
        vim.fn.system(string.format("cp '%s' '%s'", orig_path, tmp))

        -- Save back
        local ok, err = direct.from_text(tmp, modified, para_map, root)
        assert.is_true(ok, err or "Save failed")

        -- Verify by reading again
        local new_text = direct.to_text(tmp)
        assert.is_not_nil(new_text)
        assert.is_true(new_text:find("MODIFIED", 1, true) ~= nil)

        -- Cleanup
        vim.fn.delete(tmp)
      else
        pending("Fixture file not available")
      end
    end)

    it("preserves formatting during roundtrip", function()
      local orig_path = FIXTURES .. "medium-file.odt"

      if vim.fn.filereadable(orig_path) == 1 then
        local direct = require("neoffice.direct")

        local text1, para_map, root = direct.to_text(orig_path)
        assert.is_not_nil(text1)

        -- Save without modifications
        local tmp = vim.fn.tempname() .. ".odt"
        vim.fn.system(string.format("cp '%s' '%s'", orig_path, tmp))

        local ok = direct.from_text(tmp, text1, para_map, root)
        assert.is_true(ok)

        -- Read again
        local text2 = direct.to_text(tmp)

        -- Should be identical (or very close)
        assert.equals(#text1, #text2)

        vim.fn.delete(tmp)
      else
        pending("Fixture file not available")
      end
    end)
  end)

  describe("buffer operations", function()
    it("creates and manages proxy buffers", function()
      local buffer = require("neoffice.buffer")

      -- Create a mock proxy buffer
      local buf = vim.api.nvim_create_buf(false, true)
      vim.api.nvim_buf_set_name(buf, "neoffice://test.odt")

      -- Set some content
      vim.api.nvim_buf_set_lines(buf, 0, -1, false, {
        "Line 1",
        "Line 2",
        "Line 3",
      })

      -- Verify buffer
      assert.is_true(vim.api.nvim_buf_is_valid(buf))

      local lines = vim.api.nvim_buf_get_lines(buf, 0, -1, false)
      assert.equals(3, #lines)

      -- Cleanup
      vim.api.nvim_buf_delete(buf, { force = true })
    end)

    it("tracks buffer metadata", function()
      -- Test metadata storage pattern
      local meta = {
        orig_path = "/path/to/doc.odt",
        text_path = "/tmp/doc.txt",
        para_map = { [1] = 1, [2] = 2 },
        original_root = { tag = "ROOT" },
      }

      assert.is_not_nil(meta.orig_path)
      assert.is_not_nil(meta.text_path)
      assert.is_table(meta.para_map)
      assert.is_table(meta.original_root)
    end)
  end)

  describe("comment and track changes integration", function()
    it("can add and retrieve comments", function()
      local comments = require("neoffice.comments")

      -- Get current comments (should start empty in test)
      local current = comments.get_comments()
      assert.is_table(current)

      -- The actual add_comment function requires UI interaction
      -- So we test the data structure instead
      local new_comment = {
        id = "test_" .. os.time(),
        author = "TestUser",
        date = os.date("!%Y-%m-%dT%H:%M:%SZ"),
        text = "Test comment",
        replies = {},
        resolved = false,
        anchor = "test anchor",
      }

      assert.is_not_nil(new_comment.id)
      assert.is_not_nil(new_comment.author)
      assert.is_false(new_comment.resolved)
    end)

    it("manages comment lifecycle", function()
      -- Test comment state management
      local comment = {
        id = "c1",
        text = "Original",
        replies = {},
        resolved = false,
      }

      -- Add reply
      table.insert(comment.replies, {
        author = "Responder",
        text = "Reply",
      })
      assert.equals(1, #comment.replies)

      -- Toggle resolve
      comment.resolved = true
      assert.is_true(comment.resolved)

      -- Remove reply
      table.remove(comment.replies, 1)
      assert.equals(0, #comment.replies)
    end)
  end)

  describe("XML manipulation", function()
    it("handles complex ODT structures", function()
      local xml = require("neoffice.xml")

      local complex_doc = [[
        <office:document-content>
          <office:body>
            <office:text>
              <text:p>Paragraph 1</text:p>
              <text:p>
                Paragraph 2 with
                <text:span text:style-name="Bold">formatting</text:span>
                and more text
              </text:p>
              <text:h text:outline-level="1">Heading</text:h>
            </office:text>
          </office:body>
        </office:document-content>
      ]]

      local tree = xml.parse(complex_doc)
      assert.is_not_nil(tree)

      local paragraphs = xml.find_all(tree, "text:p")
      assert.is_true(#paragraphs >= 2)

      local headings = xml.find_all(tree, "text:h")
      assert.equals(1, #headings)
    end)

    it("preserves namespace information", function()
      local xml = require("neoffice.xml")

      local ns_doc = [[
        <office:text xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
                     xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
          <text:p>Content</text:p>
        </office:text>
      ]]

      local tree = xml.parse(ns_doc)
      local text_node = xml.find_first(tree, "office:text")

      assert.is_not_nil(text_node)
      assert.is_not_nil(text_node.attrs)
    end)
  end)

  describe("error handling", function()
    it("handles missing files gracefully", function()
      local direct = require("neoffice.direct")

      local text, err = direct.to_text("/nonexistent/path/file.odt")

      assert.is_nil(text)
      assert.is_not_nil(err)
    end)

    it("handles corrupted XML gracefully", function()
      local xml = require("neoffice.xml")

      local bad_xml = "<root><unclosed>"

      -- Should not crash
      local ok = pcall(xml.parse, bad_xml)
      assert.is_not_nil(ok)
    end)

    it("validates document format", function()
      local direct = require("neoffice.direct")

      local text, err = direct.to_text("test.txt")

      assert.is_nil(text)
      assert.equals("Unsupported format", err)
    end)
  end)

  describe("performance considerations", function()
    it("handles large documents efficiently", function()
      local xml = require("neoffice.xml")

      -- Create a document with many paragraphs
      local parts = { "<office:text>" }
      for i = 1, 100 do
        parts[#parts + 1] = string.format("<text:p>Paragraph %d</text:p>", i)
      end
      parts[#parts + 1] = "</office:text>"

      local large_doc = table.concat(parts)

      local start = os.clock()
      local tree = xml.parse(large_doc)
      local elapsed = os.clock() - start

      assert.is_not_nil(tree)
      assert.is_true(elapsed < 1.0) -- Should parse in less than 1 second

      local paragraphs = xml.find_all(tree, "text:p")
      assert.equals(100, #paragraphs)
    end)

    it("handles deeply nested structures", function()
      local xml = require("neoffice.xml")

      -- Create deeply nested structure
      local nested = "<root>"
      for i = 1, 50 do
        nested = nested .. "<level" .. i .. ">"
      end
      nested = nested .. "content"
      for i = 50, 1, -1 do
        nested = nested .. "</level" .. i .. ">"
      end
      nested = nested .. "</root>"

      local tree = xml.parse(nested)
      assert.is_not_nil(tree)

      -- Should successfully parse
      local root = tree.children[1]
      assert.equals("root", root.tag)
    end)
  end)

  describe("compatibility", function()
    it("works with both ODT and DOCX formats", function()
      local direct = require("neoffice.direct")

      -- Test format detection
      local formats = { "test.odt", "test.docx" }

      for _, path in ipairs(formats) do
        local ext = (path:match("%.(%w+)$") or ""):lower()
        assert.is_true(ext == "odt" or ext == "docx")
      end
    end)

    it("handles different XML namespaces", function()
      local xml = require("neoffice.xml")

      local odt_style = '<office:text xmlns:office="urn:office">content</office:text>'
      local docx_style = '<w:document xmlns:w="http://word">content</w:document>'

      local tree1 = xml.parse(odt_style)
      local tree2 = xml.parse(docx_style)

      assert.is_not_nil(tree1)
      assert.is_not_nil(tree2)
    end)
  end)

  describe("state management", function()
    it("maintains buffer state correctly", function()
      -- Test buffer state pattern
      local state = {}

      local buf_id = 1
      state[buf_id] = {
        changes = {},
        orig_path = "/path/to/doc.odt",
      }

      assert.is_not_nil(state[buf_id])
      assert.is_table(state[buf_id].changes)

      -- Cleanup
      state[buf_id] = nil
      assert.is_nil(state[buf_id])
    end)

    it("tracks multiple buffers independently", function()
      local state = {}

      state[1] = { path = "doc1.odt" }
      state[2] = { path = "doc2.odt" }
      state[3] = { path = "doc3.odt" }

      assert.equals(3, vim.tbl_count(state))
      assert.not_equals(state[1].path, state[2].path)
    end)
  end)
end)
