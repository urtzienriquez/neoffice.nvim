-- What makes typing fast must give what the slow way gives: decorations redone only where the
-- text changed, the first ones a few at a time, undo states kept as an edit after the one before,
-- the lines of the last snapshot patched, fingerprints finished while typing pauses. Each is
-- compared here with doing it all at once.
local H = dofile("tests/helpers.lua")
local neo = require("neoffice")
local changes = require("neoffice.changes")
local zip = require("neoffice.zip")
local api = vim.api

-- a long document: many paragraphs, headings, captions, a table and a box, formatted text
local function long(ext)
  local parts = {}
  for n = 1, 150 do
    local text = ("Paragraph %d with **bold**, *italic* and a [link](https://example.org/%d)."):format(n, n)
    if ext == "docx" then
      parts[#parts + 1] = ("<w:p><w:r><w:t xml:space=\"preserve\">%s</w:t></w:r></w:p>"):format(text:gsub("[*%[%]%(%)]", ""))
    else
      parts[#parts + 1] = ("<text:p>%s</text:p>"):format(text:gsub("[*%[%]%(%)]", ""))
    end
  end
  return H.copy("rich." .. ext, function(name, text)
    if ext == "docx" and name == "word/document.xml" then
      return (text:gsub("<w:body>", "<w:body>" .. table.concat(parts), 1))
    elseif ext == "odt" and name == "content.xml" then
      return (text:gsub("<office:text>", "<office:text>" .. table.concat(parts), 1))
    end
    return text
  end)
end
local function decorations()
  local ns = api.nvim_get_namespaces().neoffice_decorations
  local out = {}
  for _, m in ipairs(api.nvim_buf_get_extmarks(0, ns, 0, -1, { details = true })) do
    local d = m[4]
    out[#out + 1] = table.concat({ m[2], m[3], d.end_row or "", d.end_col or "", d.hl_group or "", d.priority or "",
      d.virt_text and vim.inspect(d.virt_text) or "" }, "|")
  end
  table.sort(out)
  return table.concat(out, "\n")
end
local function change_marks()
  local out = {}
  for _, m in ipairs(api.nvim_buf_get_extmarks(0, neo.ns.change, 0, -1, { details = true })) do
    local ch = H.doc().change[m[1]]
    out[#out + 1] = table.concat({ m[2], m[3], m[4].end_row, m[4].end_col, ch.type, ch.author or "", ch.shown or "" }, "|")
  end
  table.sort(out)
  return table.concat(out, "\n")
end
local function random(seed)
  local state = seed
  return function(n) -- (1..n, the same each run)
    state = (state * 1103515245 + 12345) % 2147483648
    return state % math.max(n, 1) + 1
  end
end
local function char_col(line, col) -- (not inside a character)
  while col > 0 and col < #line and line:byte(col + 1) >= 0x80 and line:byte(col + 1) < 0xC0 do
    col = col - 1
  end
  return col
end

for _, ext in ipairs({ "docx", "odt" }) do
  describe(ext, function()
    it("decorates where the text changed as it decorates all of it", function()
      H.open(long(ext))
      neo.decorate(0)
      local R = random(7)
      local SNIPS = { "x", " **b** ", "*i*", "\n", "\n\n", "# A heading\n\n", "Table: A caption\n\n", "[l](u)", "==h==", "\t",
        "{size = 9pt} ", "<|> ", "| a | b |\n|---|---|\n| 1 | 2 |\n", "- item\n", "x^{2}", "++u++", "::: {width = 5cm}\nboxed\n:::\n" }
      for step = 1, 60 do
        local n = api.nvim_buf_line_count(0)
        local row = R(n) - 1
        local line = api.nvim_buf_get_lines(0, row, row + 1, false)[1]
        local op = R(10)
        if op <= 6 then
          local col = char_col(line, R(#line + 1) - 1)
          api.nvim_buf_set_text(0, row, col, row, col, vim.split(SNIPS[R(#SNIPS)], "\n"))
        elseif op <= 8 then
          api.nvim_buf_set_lines(0, row, row + 1, false, {})
        else
          local col = char_col(line, math.min(R(#line + 1) - 1, #line))
          api.nvim_buf_set_text(0, row, 0, row, col, { "" })
        end
        neo.decorate_changed(api.nvim_get_current_buf())
        local now = decorations()
        neo.decorate(0)
        assert.equals(decorations(), now, "after edit " .. step)
      end
    end)

    it("decorates a long document the first time, a few paragraphs at a time, as all at once", function()
      H.open(long(ext))
      local doc = H.doc()
      assert(vim.wait(5000, function() return doc.numbers ~= nil and doc.decorating == nil end, 1))
      local gradually = decorations()
      neo.decorate(0)
      assert.equals(decorations(), gradually)
    end)

    it("puts the changes back as they were at each undo and redo, and keeps the lines of the last snapshot as they are", function()
      H.open(long(ext))
      local buf, doc = api.nvim_get_current_buf(), H.doc()
      neo.config.author = neo.config.author or "Tester"
      local R = random(11)
      local seen = {}
      for step = 1, 120 do
        vim.o.undolevels = vim.o.undolevels -- (an undo step for each)
        local op = R(10)
        local n = api.nvim_buf_line_count(0)
        local at_change = R(3) == 1 and api.nvim_buf_get_extmarks(0, neo.ns.change, 0, -1, { details = true })
        if at_change and #at_change > 0 then -- (typed right where a change starts or ends)
          local m = at_change[R(#at_change)]
          local row, col = m[2], m[3]
          if R(2) == 1 then
            row, col = m[4].end_row, m[4].end_col
          end
          if row < api.nvim_buf_line_count(0) then
            api.nvim_buf_set_text(0, row, col, row, col, { ({ "a", " ", "ab" })[R(3)] })
          end
        elseif op == 1 then
          doc.tracking = not doc.tracking
          changes.snapshot(buf)
        elseif op <= 5 then -- (one edit, or several before the text changed events)
          for _ = 1, R(2) do
            local row = R(n) - 1
            local line = api.nvim_buf_get_lines(0, row, row + 1, false)[1]
            local col = char_col(line, R(#line + 1) - 1)
            api.nvim_buf_set_text(0, row, col, row, col, vim.split(({ "w", " word ", "\n" })[R(3)], "\n"))
          end
        elseif op <= 7 then
          local row = R(n) - 1
          local line = api.nvim_buf_get_lines(0, row, row + 1, false)[1]
          local a = char_col(line, R(#line + 1) - 1)
          local b = math.min(#line, a + R(20))
          while b < #line and line:byte(b + 1) >= 0x80 and line:byte(b + 1) < 0xC0 do
            b = b + 1
          end
          api.nvim_buf_set_text(0, row, a, row, b, { "" })
        elseif op <= 9 then
          vim.cmd("silent! undo")
        else
          vim.cmd("silent! redo")
        end
        changes.after_edit(buf)
        local seq = vim.fn.undotree(buf).seq_cur
        local marks = change_marks()
        if seen[seq] and seen[seq].tracking == doc.tracking then
          assert.equals(seen[seq].marks, marks, "at state " .. seq .. ", step " .. step)
        end
        seen[seq] = { marks = marks, tracking = doc.tracking }
        assert.same(api.nvim_buf_get_lines(0, 0, -1, false), doc.shadow.lines, "the snapshot's lines, step " .. step)
      end
    end)

    it("saves an edit made right after opening the same, whether the rest is worked out while typing pauses or at once", function()
      local function edited(wait)
        local path = long(ext)
        vim.cmd("silent! %bwipeout!")
        vim.cmd("edit " .. vim.fn.fnameescape(path)) -- (no wait: before the fingerprints are all taken)
        local doc = H.doc()
        local n = api.nvim_buf_line_count(0)
        api.nvim_buf_set_text(0, n - 3, 0, n - 3, 0, { "Edited " })
        api.nvim_buf_set_lines(0, 30, 30, false, { "", "A new paragraph.", "" })
        if wait then
          assert(vim.wait(10000, function() return doc.pending == nil end, 10), "the fingerprints were never all taken")
        end
        H.save()
        local out = {}
        for name in pairs(zip.list(path)) do
          out[name] = zip.read(path, name)
        end
        return out
      end
      assert.same(edited(false), edited(true))
    end)
  end)
end

describe("underlines, superscripts, subscripts, footnotes and fields", function()
  it("are drawn as the text is decorated (the syntax only when the marks are concealed)", function()
    H.open(H.copy("empty.docx"))
    assert.equals("", vim.bo.syntax)
    api.nvim_buf_set_lines(0, 0, -1, false, { "Plain ++under++ x^{2} H_{2}O a note^[the note] here." })
    neo.decorate(0)
    local ns = api.nvim_get_namespaces().neoffice_decorations
    local groups = {}
    for _, m in ipairs(api.nvim_buf_get_extmarks(0, ns, 0, -1, { details = true })) do
      if m[4].hl_group and m[4].hl_group:match("^neoffice") then
        groups[m[4].hl_group] = (groups[m[4].hl_group] or "") .. api.nvim_buf_get_text(0, m[2], m[3], m[4].end_row, m[4].end_col, {})[1]
      end
    end
    assert.equals("under", groups.neofficeUnderline)
    assert.equals("2", groups.neofficeSuperscript)
    assert.equals("2", groups.neofficeSubscript)
    assert.equals("the note", groups.neofficeFootnote)
    assert.equals("^[]", groups.neofficeFootnoteMark)
  end)
end)
