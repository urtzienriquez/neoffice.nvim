-- What neoffice must never do: change anything in a document that you didn't change.
local H = dofile("tests/helpers.lua")
local api = vim.api

for _, name in ipairs({ "rich.docx", "rich.odt", "office.docx" }) do
  describe(name, function()
    local path, before

    before_each(function()
      path = H.copy(name)
      before = H.bytes(path)
      H.open(path)
    end)

    it("saving without editing leaves the file byte for byte as it was", function()
      H.save()
      assert.equals(before, H.bytes(path))
      assert.is_false(vim.bo.modified)
    end)

    it("saving after undoing every edit leaves the file as it was", function()
      H.replace(2, "bold", "strong")
      H.set(12, "Something else entirely.")
      vim.cmd("silent undo 0")
      H.save()
      assert.equals(before, H.bytes(path))
    end)

    it("editing a paragraph rewrites only that paragraph, and no other part", function()
      local entries, paras = H.entries(path), H.paragraph_xml(path)
      H.replace(12, "Last", "Final")
      H.save()
      assert.same({ H.MAIN[name:match("%w+$")] }, H.changed(entries, H.entries(path)))
      assert.equals(1, H.differing(paras, H.paragraph_xml(path)))
    end)

    it("adding a paragraph changes only the document's text part", function()
      local entries, paras = H.entries(path), H.paragraph_xml(path)
      api.nvim_buf_set_lines(0, H.row(12) - 1, H.row(12) - 1, false, { "A new paragraph.", "" })
      H.save()
      assert.same({ H.MAIN[name:match("%w+$")] }, H.changed(entries, H.entries(path)))
      local after = H.paragraph_xml(path)
      assert.equals(#paras + 1, #after)
      assert.same(vim.list_slice(paras, 1, 11), vim.list_slice(after, 1, 11))
    end)

    it("saving twice writes nothing the second time", function()
      H.replace(12, "Last", "Final")
      H.save()
      local first = H.bytes(path)
      H.save()
      assert.equals(first, H.bytes(path))
    end)

    it("turning tracking on and off doesn't change the buffer or the file", function()
      vim.cmd("Neoffice track on")
      vim.cmd("Neoffice track off")
      vim.cmd("Neoffice track")
      assert.is_false(vim.bo.modified)
      H.save()
      assert.equals(before, H.bytes(path))
    end)

    it("looking at comments and changes doesn't change anything", function()
      vim.cmd("Neoffice comments")
      vim.cmd("Neoffice comments")
      vim.cmd("Neoffice changes")
      vim.cmd("lclose")
      vim.cmd("Neoffice next")
      vim.cmd("Neoffice prev")
      assert.is_false(vim.bo.modified)
      H.save()
      assert.equals(before, H.bytes(path))
    end)

    it("a failed save (an image that doesn't exist) leaves the file untouched", function()
      H.set(12, "![](no/such/image.png)")
      H.save()
      assert.equals(before, H.bytes(path))
      assert.is_true(vim.bo.modified)
      H.set(12, "Last paragraph, fixed.") -- and the next save works
      H.save()
      assert.equals("Last paragraph, fixed.", H.reopen() and H.paras()[12])
    end)

    it("writing to another file leaves the original untouched", function()
      H.replace(12, "Last", "Final")
      local other = vim.fn.tempname() .. "." .. name:match("%w+$")
      vim.cmd("silent write " .. other)
      assert.equals(before, H.bytes(path))
      assert.is_true(vim.bo.modified)
      H.open(other)
      assert.equals("Final paragraph.", H.paras()[12])
    end)

    it("a copy written without edits is identical", function()
      local other = vim.fn.tempname() .. "." .. name:match("%w+$")
      vim.cmd("silent write " .. other)
      assert.equals(before, H.bytes(other))
    end)

    it("shows the same text after saving and reopening", function()
      H.replace(2, "bold", "strong")
      api.nvim_buf_set_lines(0, H.row(3) - 1, H.row(3), false, { "Plain **bold** and *more*." })
      local shown = H.lines()
      H.reopen()
      assert.same(shown, H.lines())
    end)

    if H.has_libreoffice then
      it("LibreOffice opens the edited file", function()
        H.replace(12, "Last", "Final")
        api.nvim_buf_set_lines(0, 1, 1, false, { "", "## New *section*" })
        H.save()
        local ok, text = H.libreoffice_reads(path, "Final paragraph.")
        assert.is_true(ok, text)
        assert.truthy(text:find("New section", 1, true))
      end)
    end
  end)
end

describe("a document that can't be read", function()
  it("opens read-only and empty, and can't be written over", function()
    local path = vim.fn.tempname() .. ".docx"
    local f = assert(io.open(path, "wb"))
    f:write("this is not a zip file")
    f:close()
    H.open(path)
    assert.is_false(vim.bo.modifiable)
    assert.same({ "" }, H.lines())
    pcall(vim.cmd, "silent write!")
    assert.equals("this is not a zip file", H.bytes(path))
  end)
end)

describe("a document without paragraphs", function()
  for _, name in ipairs({ "empty.docx", "empty.odt" }) do
    it(name .. " stays as it was unless you type something", function()
      local path = H.copy(name)
      local before = H.bytes(path)
      H.open(path)
      H.save()
      assert.equals(before, H.bytes(path))
      api.nvim_buf_set_lines(0, 0, -1, false, { "# Title", "", "First words." })
      H.reopen()
      assert.same({ "# Title", "First words." }, H.paras())
    end)
  end
end)

describe("a new document", function()
  for _, ext in ipairs({ "docx", "odt", "doc" }) do
    it("(." .. ext .. ") opens empty and is created when written", function()
      if ext == "doc" and not H.has_libreoffice then
        return pending("needs LibreOffice")
      end
      local path = vim.fn.tempname() .. "/new." .. ext
      vim.fn.mkdir(vim.fs.dirname(path), "p")
      H.open(path)
      assert.same({ "" }, H.lines())
      assert.is_true(vim.bo.modifiable)
      assert.is_nil(vim.uv.fs_stat(path)) -- (not until written)
      api.nvim_buf_set_lines(0, 0, -1, false, { "# A title", "", "Some **bold** text.", "", "![](" .. H.FIXTURES .. "dot.png)" })
      H.save()
      assert.truthy(vim.uv.fs_stat(path))
      assert.is_false(vim.bo.modified)
      H.replace(2, "Some", "More") -- (and saving again works on the file itself)
      H.reopen()
      assert.same({ "# A title", "More **bold** text." }, vim.list_slice(H.paras(), 1, 2))
      assert.truthy(H.paras()[3]:match("^!%[%]%("))
    end)
  end

  it("written right away is a valid empty document", function()
    local path = vim.fn.tempname() .. ".docx"
    H.open(path)
    H.save()
    H.open(path)
    assert.same({ "" }, H.lines())
    if H.has_libreoffice then
      assert.is_true((H.libreoffice_reads(path, "")))
    end
  end)
end)

describe("the file on disk", function()
  for _, ext in ipairs({ "docx", "odt" }) do
    it("(." .. ext .. ") is written again if it was deleted (or moved) since the last save", function()
      local path = H.copy("rich." .. ext)
      H.open(path)
      H.replace(12, "Last", "Final")
      H.save()
      os.remove(path)
      H.replace(1, "Introduction", "Preface")
      H.save()
      assert.truthy(vim.uv.fs_stat(path))
      H.open(path)
      assert.equals("# Preface", H.paras()[1])
      assert.equals("Final paragraph.", H.paras()[12])
    end)

    it("(." .. ext .. ") isn't written over when another program changed it, unless with :w!", function()
      local path = H.copy("rich." .. ext)
      H.open(path)
      H.replace(12, "Last", "Final")
      vim.uv.fs_copyfile(H.FIXTURES .. "empty." .. ext, path) -- (another program saves it)
      local theirs = H.bytes(path)
      H.save()
      assert.equals(theirs, H.bytes(path))
      assert.is_true(vim.bo.modified)
      vim.cmd("silent write!")
      assert.is_false(vim.bo.modified)
      H.open(path)
      assert.equals("Final paragraph.", H.paras()[12])
    end)

    it("(." .. ext .. ") gets the edits even after they were written to another file first", function()
      local path = H.copy("rich." .. ext)
      H.open(path)
      H.replace(12, "Last", "Final")
      vim.cmd("silent write " .. vim.fn.tempname() .. "." .. ext)
      H.save()
      H.open(path)
      assert.equals("Final paragraph.", H.paras()[12])
    end)
  end
end)

describe("a part that would come out broken", function()
  it("isn't written: the save says so and the file stays as it was", function()
    local path = H.copy("rich.docx")
    H.open(path)
    local before = H.bytes(path)
    local doc = H.doc()
    local save = doc.fmt.save
    doc.fmt.save = function(...) -- (as a bug would: two roots)
      local entries = save(...)
      entries["word/document.xml"] = "<w:document/><w:body/>"
      return entries
    end
    H.set(1, "Edited.")
    local errors = {}
    local notify = vim.notify
    vim.notify = function(msg) errors[#errors + 1] = msg end
    H.save()
    vim.notify, doc.fmt.save = notify, save
    assert.equals(before, H.bytes(path))
    assert.truthy(table.concat(errors, "\n"):find("broken word/document.xml", 1, true))
  end)
end)

describe("a change right after opening, before all is worked out", function()
  local zip = require("neoffice.zip")
  local function open_now(path) -- (as H.open, but without letting what waits for the text to show run)
    vim.cmd("silent! %bwipeout!")
    vim.cmd("edit " .. vim.fn.fnameescape(path))
  end
  -- a long document: what its paragraphs are as opened is worked out a few at a time, after the text
  -- shows (see M.finish_sigs); a change before that is done must still be saved, and only it
  local function long(ext)
    return H.copy("rich." .. ext, function(name, text)
      local many = {}
      for n = 1, 120 do
        many[#many + 1] = ext == "docx" and ("<w:p><w:r><w:t>Paragraph %d.</w:t></w:r></w:p>"):format(n)
          or ("<text:p>Paragraph %d.</text:p>"):format(n)
      end
      if ext == "docx" and name == "word/document.xml" then
        return (text:gsub("<w:body>", "<w:body>" .. table.concat(many), 1))
      elseif ext == "odt" and name == "content.xml" then
        return (text:gsub("<office:text>", "<office:text>" .. table.concat(many), 1))
      end
      return text
    end)
  end
  for _, ext in ipairs({ "docx", "odt" }) do
    it("saves an edit to a paragraph not worked out yet, and rewrites no other (" .. ext .. ")", function()
      local path = long(ext)
      open_now(path)
      assert.truthy(H.doc().pending, "(the document is long enough for this test)")
      local k = vim.fn.index(H.paras(), "Paragraph 100.") + 1
      H.replace(k, "100", "one hundred")
      local xml = require("neoffice.xml")
      local before = xml.serialize(xml.parse(zip.read(path, H.MAIN[ext]))) -- (as written: "<a/>" for "<a />"...)
      H.save()
      local after = zip.read(path, H.MAIN[ext])
      assert.truthy(after:find("Paragraph one hundred.", 1, true))
      -- (only that paragraph changed: the paragraphs before and after it as they were)
      local function between(x, a, b)
        return x:sub((x:find(a, 1, true)), (x:find(b, 1, true)))
      end
      assert.equals(between(before, "Paragraph 1.", "Paragraph 99."), between(after, "Paragraph 1.", "Paragraph 99."))
      local rest = before:find("Paragraph 101.", 1, true)
      assert.equals(before:sub(rest), after:sub((after:find("Paragraph 101.", 1, true))))
    end)

    it("saves a deletion tracked right after opening (" .. ext .. ")", function()
      local path = long(ext)
      open_now(path)
      vim.cmd("Neoffice track on")
      local k = vim.fn.index(H.paras(), "Paragraph 90.") + 1
      H.replace(k, "90", "")
      H.reopen()
      assert.equals(1, #H.changes(require("neoffice").config.author))
    end)

    it("saves a comment added right after opening (" .. ext .. ")", function()
      local path = long(ext)
      open_now(path)
      local k = vim.fn.index(H.paras(), "Paragraph 80.") + 1
      H.cursor(k, "80")
      H.input = "Said early."
      require("neoffice.comments").add()
      H.save()
      assert.truthy((zip.read(path, ext == "docx" and "word/comments.xml" or "content.xml") or ""):find("Said early.", 1, true))
    end)

    it("takes a paragraph typed again the same as changed: its marks may not be (" .. ext .. ")", function()
      local path = long(ext)
      open_now(path)
      local k = vim.fn.index(H.paras(), "Paragraph 110.") + 1
      H.replace(k, "110", "")
      H.replace(k, "Paragraph .", "Paragraph 110.")
      require("neoffice").finish_sigs(vim.api.nvim_get_current_buf())
      local doc = H.doc()
      assert.is_false(doc.sigs[vim.fn.index(doc.texts, "Paragraph 110.") + 1]) -- (written at the next save)
      H.save()
      assert.truthy(zip.read(path, H.MAIN[ext]):find("Paragraph 110.", 1, true))
    end)
  end
end)

describe("saving again and again in one session", function()
  local function markup_para_line(line)
    return line:sub(1, 1) == "{" and require("neoffice.markup").para_line(line) or nil
  end
  -- (at a save, a paragraph as it was since the last one keeps its fingerprint: what changed is
  -- still found, whatever changed it)
  local zip = require("neoffice.zip")
  for _, ext in ipairs({ "docx", "odt" }) do
    it("writes each change: text, a format's line, a comment (" .. ext .. ")", function()
      local path = H.copy("rich." .. ext)
      H.open(path)
      local function main() return zip.read(path, H.MAIN[ext]) end
      H.replace(1, "Introduction", "Edited title")
      H.save()
      assert.truthy(main():find("Edited title", 1, true))
      local r = H.row(3)
      vim.api.nvim_buf_set_lines(0, r - 1, r - 1, false, { "{align = center}" }) -- (a format: rows not of the paragraphs)
      H.save()
      -- (also for a paragraph further down, whose rows the new line didn't touch)
      H.reopen()
      local k = vim.fn.index(H.paras(), "Text with inserted words and deleted words inside.") + 1
      assert.truthy(k > 0, table.concat(H.lines(), "\n"))
      local above
      for row = H.row(k) - 1, 1, -1 do
        above = above or markup_para_line(H.lines()[row])
      end
      assert.equals("center", above and above.align, table.concat(H.lines(), "\n"))
      H.cursor(4, "inserted")
      H.input = "Late note."
      require("neoffice.comments").add()
      H.save()
      assert.truthy((zip.read(path, ext == "docx" and "word/comments.xml" or "content.xml") or ""):find("Late note.", 1, true))
      H.replace(4, "words", "WORDS")
      H.save()
      H.reopen()
      assert.truthy(vim.tbl_contains(H.paras(), "Edited title") or table.concat(H.paras(), "\n"):find("Edited title", 1, true))
      assert.truthy(table.concat(H.paras(), "\n"):find("WORDS", 1, true))
      assert.equals("Late note.", H.comments()[#H.comments()].text)
    end)
  end
end)

describe("a document changed by another program while open", function()
  local zip = require("neoffice.zip")
  for _, ext in ipairs({ "docx", "odt" }) do
    it("is read again when Neovim notices (and again), showing the change (" .. ext .. ")", function()
      local path = H.copy("rich." .. ext)
      H.open(path)
      for round = 1, 2 do
        local main = zip.read(path, H.MAIN[ext])
        local word = round == 1 and "Introduction" or "Overview"
        zip.write(path, { [H.MAIN[ext]] = (main:gsub(word, round == 1 and "Overview" or "Summary", 1)) })
        local errors = {}
        local notify = vim.notify
        vim.notify = function(msg, level) errors[#errors + 1] = level == vim.log.levels.ERROR and msg or nil end
        local stamp = vim.uv.fs_stat(path).mtime.sec + round * 10 -- (as another program would: later)
        vim.uv.fs_utime(path, stamp, stamp)
        vim.o.autoread = true
        vim.cmd("checktime") -- (Neovim reads it again, keeping the buffer: as when LibreOffice saved it)
        vim.wait(0)
        vim.notify = notify
        assert.same({}, errors)
        assert.equals(round == 1 and "# Overview" or "# Summary", H.paras()[1])
        local was, now = round == 1 and "This" or "Now this", round == 1 and "Now this" or "Then this"
        H.replace(2, was, now) -- (and editing, then saving, works on what was read)
        assert.truthy(#H.doc().touched > 0 and not H.doc().untracked, "what changes is still followed after reading it again")
        H.save()
        assert.truthy(zip.read(path, H.MAIN[ext]):find(now, 1, true))
      end
    end)
  end
end)
