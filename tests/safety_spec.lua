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
      require("neoffice.comments").add(nil, "Said early.")
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
      require("neoffice.comments").add(nil, "Late note.")
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

describe("writing the file", function()
  local zip = require("neoffice.zip")
  local function notes(fn)
    local said, notify = {}, vim.notify
    vim.notify = function(msg) said[#said + 1] = msg end
    local ok, err = pcall(fn)
    vim.notify = notify
    assert(ok, err)
    return table.concat(said, "\n")
  end

  for _, ext in ipairs({ "docx", "odt" }) do
    it("leaves the file as it was when what is written comes out short, as on a full disk (" .. ext .. ")", function()
      local path = H.copy("rich." .. ext)
      local before = H.bytes(path)
      H.open(path)
      H.set(12, "Something else entirely.")
      local real_open = io.open
      io.open = function(p, mode) -- (each file neoffice writes beside another gets half of what it should)
        local f, err = real_open(p, mode)
        if f and mode == "wb" and p:match("%.neoffice~$") then
          return setmetatable({}, { __index = {
            write = function(_, data) return f:write(data:sub(1, math.floor(#data / 2))) end,
            close = function() return f:close() end,
          } })
        end
        return f, err
      end
      local said = notes(function() H.save() end)
      io.open = real_open
      assert.truthy(said:find("Save failed", 1, true), said)
      assert.equals(before, H.bytes(path))
      H.save() -- (and once the disk has room, it is written)
      assert.truthy(zip.read(path, H.MAIN[ext]):find("Something else entirely.", 1, true))
    end)
  end

  it("writes through a symbolic link to the file it points to, and keeps the file's permissions", function()
    local path = H.copy("rich.docx")
    vim.uv.fs_chmod(path, tonumber("640", 8))
    local link = vim.fn.tempname() .. ".docx"
    assert(vim.uv.fs_symlink(path, link))
    H.open(link)
    H.set(12, "Written through a link.")
    H.save()
    assert.equals("link", vim.uv.fs_lstat(link).type)
    assert.truthy(zip.read(path, "word/document.xml"):find("Written through a link.", 1, true))
    assert.equals(tonumber("640", 8), bit.band(vim.uv.fs_stat(path).mode, 4095))
  end)

  it("says when the file seems open in Word or LibreOffice", function()
    for _, lock in ipairs({ "~$ch.docx", ".~lock.rich.docx#" }) do
      local path = H.copy("rich.docx")
      local f = assert(io.open(vim.fs.joinpath(vim.fs.dirname(path), lock), "w"))
      f:close()
      local said = notes(function() H.open(path) end)
      assert.truthy(said:find("seems open in", 1, true), said)
    end
  end)

  it("says when a document is signed, or protected in Word", function()
    local path = H.copy("rich.docx")
    zip.write(path, { ["_xmlsignatures/sig1.xml"] = "<Signature/>" })
    local said = notes(function() H.open(path) end)
    assert.truthy(said:find("digitally signed", 1, true), said)
    path = H.copy("rich.docx")
    local settings = zip.read(path, "word/settings.xml") or '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"></w:settings>'
    settings = settings:gsub("(<w:settings[^>]*>)", '%1<w:documentProtection w:edit="forms" w:enforcement="1"/>', 1)
    zip.write(path, { ["word/settings.xml"] = settings })
    said = notes(function() H.open(path) end)
    assert.truthy(said:find("protected in Word", 1, true), said)
  end)
end)

describe("what a save may never write", function()
  local zip = require("neoffice.zip")
  local function count(text, pattern)
    return select(2, text:gsub(pattern, ""))
  end
  local function said_by(fn)
    local said, notify = {}, vim.notify
    vim.notify = function(msg) said[#said + 1] = msg end
    local ok, err = pcall(fn)
    vim.notify = notify
    assert(ok, err)
    return table.concat(said, "\n")
  end

  it("a bookmark twice: an empty paragraph holding one, with what was around it deleted", function()
    local function p(text, style)
      return ("<w:p>%s<w:r><w:t>%s</w:t></w:r></w:p>"):format(style and ('<w:pPr><w:pStyle w:val="%s"/></w:pPr>'):format(style) or "", text)
    end
    local path = H.copy("rich.docx", function(name, text)
      if name ~= "word/document.xml" then
        return text
      end
      local body = p("First paragraph.") .. p("Second paragraph.")
        .. '<w:p><w:bookmarkStart w:id="932" w:name="appendix"/><w:bookmarkEnd w:id="932"/></w:p>'
        .. '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:p><w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple></w:p></w:tc></w:tr></w:tbl>'
        .. p("A caption under the table.") .. p("Last heading", "Heading1")
      return (text:gsub("<w:body>.-(<w:sectPr)", "<w:body>" .. body .. "%1", 1))
    end)
    H.open(path)
    local lines = H.lines()
    local from
    for r, line in ipairs(lines) do
      if line:match("^|") and not from then
        from = r
      end
    end
    assert.truthy(from, table.concat(lines, "\n"))
    if lines[from - 1]:match("^{") then -- (its options' line too)
      from = from - 1
    end
    api.nvim_buf_set_lines(0, from - 1, -1, false, {}) -- (the table and all after it)
    H.save()
    local main = zip.read(path, "word/document.xml")
    assert.equals(1, count(main, 'w:bookmarkStart w:id="932"'))
    assert.equals(1, count(main, 'w:bookmarkEnd w:id="932"'))
    assert.is_nil(main:find("A caption under the table.", 1, true))
  end)

  it("a bookmark twice: an empty paragraph holding one, its mark gone into another paragraph", function()
    local function p(text)
      return ("<w:p><w:r><w:t>%s</w:t></w:r></w:p>"):format(text)
    end
    local path = H.copy("rich.docx", function(name, text)
      if name ~= "word/document.xml" then
        return text
      end
      local body = p("First paragraph.") .. p("Para A, before the bookmark.")
        .. '<w:p><w:bookmarkStart w:id="933" w:name="marked"/><w:bookmarkEnd w:id="933"/></w:p>'
        .. p("Para B, after it.") .. "<w:p/>" .. p("Para C.") .. "<w:p/>" .. p("Last paragraph.")
      return (text:gsub("<w:body>.-(<w:sectPr)", "<w:body>" .. body .. "%1", 1))
    end)
    H.open(path)
    -- (its marks moved into another paragraph, as when lines are joined, that one edited; the empty
    -- one still there: kept as it was, it would hold the bookmark too)
    local neo = require("neoffice")
    local c = vim.iter(ipairs(H.lines())):find(function(_, l) return l:find("^Para C") end)
    for _, m in ipairs(api.nvim_buf_get_extmarks(0, neo.ns.object, 0, -1, { details = true })) do
      local o = H.doc().object[m[1]]
      if o and o.node and o.node.tag:match("^w:bookmark") then
        api.nvim_buf_set_extmark(0, neo.ns.object, c - 1, 0, { id = m[1], end_row = c - 1, end_col = 0, right_gravity = false })
      end
    end
    api.nvim_buf_set_text(0, c - 1, #"Para C", c - 1, #"Para C", { ", edited" })
    local said = said_by(function() H.save() end)
    assert.is_nil(said:find("Save failed", 1, true), said)
    local main = zip.read(path, "word/document.xml")
    assert.equals(1, count(main, 'w:bookmarkStart w:id="933"'))
    assert.equals(1, count(main, 'w:bookmarkEnd w:id="933"'))
  end)

  it("a drawing's id twice: a new image's id is above those in headers and footers too", function()
    local path = H.copy("rich.docx")
    local main = zip.read(path, "word/document.xml")
    local top = 0
    for id in main:gmatch(':docPr[^>]- id="(%d+)"') do
      top = math.max(top, tonumber(id))
    end
    zip.write(path, { ["word/footer9.xml"] = '<?xml version="1.0" encoding="UTF-8"?><w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><w:p><w:r><w:drawing><wp:inline><wp:docPr id="' .. (top + 1) .. '" name="Logo"/></wp:inline></w:drawing></w:r></w:p></w:ftr>' })
    H.open(path)
    H.set(12, "An image: ![a dot](" .. H.FIXTURES .. "dot.png)")
    H.save()
    local ids = {}
    for _, part in ipairs({ "word/document.xml", "word/footer9.xml" }) do
      for id in zip.read(path, part):gmatch(':docPr[^>]- id="(%d+)"') do
        assert.is_nil(ids[id], "drawing id " .. id .. " twice")
        ids[id] = true
      end
    end
  end)

  for _, alt in ipairs({ "a dot", "another alt text" }) do
  it("a drawing's id twice: a picture's line copied right above it, where its mark starts (" .. alt .. ")", function()
    local path = H.copy("empty.docx")
    H.open(path)
    api.nvim_buf_set_lines(0, 0, -1, false, { "First.", "", "![a dot](" .. H.FIXTURES .. "dot.png){width = 2cm, float = right}", "Beside it.", "", "Last." })
    H.save()
    H.reopen()
    for r, line in ipairs(H.lines()) do
      if line:find("^!%[") then
        local copy = line:gsub("!%[[^%]]*%]", "![" .. alt .. "]", 1)
        api.nvim_buf_set_lines(0, r - 1, r - 1, false, { copy, "Copied.", "" })
        break
      end
    end
    local said = said_by(function() H.save() end)
    assert.is_nil(said:find("Save failed", 1, true), said)
    local main = zip.read(path, "word/document.xml")
    assert.truthy(main:find("Copied.", 1, true))
    local ids, n = {}, 0
    for id in main:gmatch(':docPr[^>]- id="(%d+)"') do
      assert.is_nil(ids[id], "drawing id " .. id .. " twice")
      ids[id], n = true, n + 1
    end
    assert.equals(2, n)
  end)
  end

  for _, ext in ipairs({ "docx", "odt" }) do
    it("a note's id twice: a paragraph starting with a footnote copied right above it (" .. ext .. ")", function()
      local path = H.copy("empty." .. ext)
      H.open(path)
      local line = "^[A note.] starts this paragraph."
      api.nvim_buf_set_lines(0, 0, -1, false, { "First.", "", line, "", "Last." })
      H.save()
      H.reopen()
      for r, l in ipairs(H.lines()) do
        if l == line then
          api.nvim_buf_set_lines(0, r - 1, r - 1, false, { l, "" })
          break
        end
      end
      local said = said_by(function() H.save() end)
      assert.is_nil(said:find("Save failed", 1, true), said)
      local main = zip.read(path, H.MAIN[ext])
      local ids, n = {}, 0
      for id in main:gmatch(ext == "docx" and 'footnoteReference w:id="([^"]+)"' or '<text:note [^>]-text:id="([^"]+)"') do
        assert.is_nil(ids[id], "note id " .. id .. " twice")
        ids[id], n = true, n + 1
      end
      assert.equals(2, n)
      H.reopen()
      assert.equals(2, #vim.tbl_filter(function(l) return l == line end, H.lines()))
    end)
  end

  for _, ext in ipairs({ "docx", "odt" }) do
    it("writes a paragraph whose text changed even when its fingerprint says it didn't (" .. ext .. ")", function()
      local path = H.copy("empty." .. ext)
      H.open(path)
      api.nvim_buf_set_lines(0, 0, -1, false, { "One paragraph here.", "", "Another, different one.", "", "Last." })
      H.save()
      H.reopen()
      local neo, buf = require("neoffice"), api.nvim_get_current_buf()
      neo.finish_sigs(buf)
      local doc = H.doc()
      -- (a fingerprint taken from the wrong paragraph, as when the lines matched paired it with
      -- another: the first's is the second's; then the first is typed as the second)
      doc.sigs[1] = doc.sigs[2]
      api.nvim_buf_set_lines(0, 0, 1, false, { "Another, different one." })
      local said = said_by(function() H.save() end)
      assert.is_nil(said:find("Save failed", 1, true), said)
      local main = zip.read(path, H.MAIN[ext])
      assert.is_nil(main:find("One paragraph here.", 1, true), "the edit is lost")
      assert.equals(2, count(main, "Another, different one."))
    end)
  end

  it("a comment never shown (its mark in deleted text) stays when another is resolved, replied to or deleted", function()
    for _, action in ipairs({ "resolve", "reply", "delete" }) do
      local path = H.copy("rich.docx", function(name, text)
        if name ~= "word/document.xml" then
          return text
        end
        -- (comment 7: its reference alone in a tracked deletion, as Word may leave it)
        return (text:gsub("(<w:del w:id=\"1\" w:author=\"Bob\" w:date=\"2024%-01%-02T10:00:00Z\">)", '<w:del w:id="97" w:author="Bob" w:date="2024-01-02T10:00:00Z"><w:r><w:commentReference w:id="7"/></w:r></w:del>%1', 1))
      end)
      local comments = zip.read(path, "word/comments.xml")
      zip.write(path, { ["word/comments.xml"] = comments:gsub("</w:comments>", '<w:comment w:id="7" w:author="Carol" w:date="2024-01-03T10:00:00Z"><w:p><w:r><w:t>Never shown.</w:t></w:r></w:p></w:comment></w:comments>') })
      assert.truthy(zip.read(path, "word/document.xml"):find('commentReference w:id="7"', 1, true))
      H.open(path)
      local m = api.nvim_buf_get_extmarks(0, require("neoffice").ns.comment, 0, -1, {})[1]
      api.nvim_win_set_cursor(0, { m[2] + 1, m[3] })
      local c = require("neoffice.comments")
      if action == "resolve" then
        c.resolve()
      elseif action == "reply" then
        c.reply("A reply.")
      else
        c.delete()
      end
      local said = said_by(function() H.save() end)
      assert.is_nil(said:find("Save failed", 1, true), said)
      local after = zip.read(path, "word/comments.xml")
      assert.truthy(after:find("Never shown.", 1, true), action .. ": the comment never shown is gone")
      assert.truthy(zip.read(path, "word/document.xml"):find('commentReference w:id="7"', 1, true))
    end
  end)

  it("a deleted comment leaves no mark behind, not even one in a deletion written back as it was", function()
    local path = H.copy("rich.docx", function(name, text)
      if name ~= "word/document.xml" then
        return text
      end
      -- (comment 0's reference in a tracked deletion of its own, as Word may leave it)
      return (text:gsub('<w:r><w:rPr><w:rStyle w:val="CommentReference" /></w:rPr><w:commentReference w:id="0" /></w:r>',
        '<w:del w:id="98" w:author="Bob" w:date="2024-01-02T10:00:00Z">%0</w:del>', 1))
    end)
    assert.truthy(zip.read(path, "word/document.xml"):find('<w:del w:id="98"', 1, true))
    H.open(path)
    local m = api.nvim_buf_get_extmarks(0, require("neoffice").ns.comment, 0, -1, {})[1]
    api.nvim_win_set_cursor(0, { m[2] + 1, m[3] })
    require("neoffice.comments").delete()
    local said = said_by(function() H.save() end)
    assert.is_nil(said:find("Save failed", 1, true), said)
    local main = zip.read(path, "word/document.xml")
    assert.is_nil(main:find('comment%a+ w:id="0"'), "a mark of the deleted comment is left")
    assert.is_nil((zip.read(path, "word/comments.xml") or ""):find('w:comment w:id="0"', 1, true))
  end)

  it("isn't written at all when it would have more of what Word repairs than the document had", function()
    local path = H.copy("rich.docx")
    local before = H.bytes(path)
    H.open(path)
    H.set(12, "Edited.")
    local xml = require("neoffice.xml")
    local body = xml.find(H.doc().ctx.root, "w:body")
    for _ = 1, 2 do -- (as a bug would: a bookmark written twice)
      xml.append(body, xml.elem("w:bookmarkStart", { { "w:id", "777" }, { "w:name", "twice" } }), 1)
    end
    local said = said_by(function() H.save() end)
    assert.truthy(said:find("Save failed", 1, true) and said:find("bookmark", 1, true), said)
    assert.equals(before, H.bytes(path))
  end)

  it("a comment started twice: one whose start Word put inside a deletion", function()
    local path = H.copy("rich.docx", function(name, text)
      if name ~= "word/document.xml" then
        return text
      end
      return (text:gsub('<w:commentRangeStart w:id="0" />', '<w:del w:id="91" w:author="Bob" w:date="2024-01-02T10:00:00Z"><w:commentRangeStart w:id="0" /><w:bookmarkStart w:id="95" w:name="inside"/><w:bookmarkEnd w:id="95"/><w:r><w:delText xml:space="preserve">gone </w:delText></w:r></w:del>', 1))
    end)
    H.open(path)
    local n
    for k, text in ipairs(H.paras()) do
      n = n or text:find("and more.", 1, true) and k
    end
    H.replace(n, "and more.", "and others.")
    H.save()
    local main = zip.read(path, "word/document.xml")
    assert.truthy(main:find("and others.", 1, true))
    assert.equals(1, count(main, 'w:commentRangeStart w:id="0"'))
    assert.equals(1, count(main, 'w:bookmarkStart w:id="95"'))
    assert.truthy(main:find("gone ", 1, true), "the deleted text is kept")
  end)

  it("a picture twice: one Word put in the same run as a text box", function()
    local path = H.copy("empty.docx")
    H.open(path)
    api.nvim_buf_set_lines(0, 0, -1, false, { "First.", "", "::: {width = 4cm, float = left}", "In a box.", ":::", "",
      "![a dot](" .. H.FIXTURES .. "dot.png){width = 2cm, float = right}", "Text beside the picture.", "", "Second.", "", "Third." })
    H.save()
    local main = zip.read(path, "word/document.xml")
    -- (the box's run and the picture's, one run: as Word may write them)
    local merged, n = main:gsub("(</mc:AlternateContent>)</w:r><w:r>(<w:drawing>)", "%1%2", 1)
    assert.equals(1, n, "the box's run right before the picture's")
    zip.write(path, { ["word/document.xml"] = merged })
    H.open(path)
    local lines = H.lines()
    for r, line in ipairs(lines) do -- (the picture's paragraph edited: written again)
      local at = line:find("beside the picture", 1, true)
      if at then
        api.nvim_buf_set_text(0, r - 1, at - 1, r - 1, at - 1 + #"beside", { "next to" })
        break
      end
    end
    local said = said_by(function() H.save() end)
    assert.is_nil(said:find("Save failed", 1, true), said)
    main = zip.read(path, "word/document.xml")
    local ids = {}
    for id in main:gmatch(':docPr[^>]- id="(%d+)"') do
      assert.is_nil(ids[id], "drawing id " .. id .. " twice")
      ids[id] = true
    end
    assert.truthy(main:find("next to the picture", 1, true))
    H.reopen()
    local text = table.concat(H.lines(), "\n")
    assert.truthy(text:find("In a box.", 1, true) and text:find("a dot", 1, true), text)
  end)

  it("a table cell ending with a table: its paragraphs deleted, the table in it kept", function()
    local path = H.copy("rich.docx", function(name, text)
      if name ~= "word/document.xml" then
        return text
      end
      local inner = '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="1000"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>inner</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'
      local outer = '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid><w:tr>'
        .. '<w:tc><w:p><w:r><w:t>before inner</w:t></w:r></w:p>' .. inner .. '<w:p><w:r><w:t>after inner</w:t></w:r></w:p></w:tc>'
        .. '<w:tc><w:p><w:r><w:t>other cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'
      return (text:gsub("<w:body>", "<w:body><w:p><w:r><w:t>Before.</w:t></w:r></w:p>" .. outer, 1))
    end)
    H.open(path)
    api.nvim_buf_set_lines(0, 0, -1, false, { "Only this." })
    local said = said_by(function() H.save() end)
    assert.is_nil(said:find("Save failed", 1, true), said)
    local main = zip.read(path, "word/document.xml")
    assert.truthy(main:find("Only this.", 1, true))
    assert.truthy(main:find("inner", 1, true), "the table inside, not shown, is kept")
    assert.is_nil(main:find("</w:tbl></w:tc>", 1, true), "a cell ends with a table")
  end)

  for _, ext in ipairs({ "docx", "odt" }) do
    it("loses no row of a table broken in two by a line that isn't a row (" .. ext .. ")", function()
      local path = H.copy("empty." .. ext)
      H.open(path)
      local lines = { "Before.", "", "| A | B | C |", "|---|---|---|" }
      for i = 1, 6 do
        lines[#lines + 1] = ("| r%d | x%d | y%d |"):format(i, i, i)
      end
      vim.list_extend(lines, { "", "After." })
      api.nvim_buf_set_lines(0, 0, -1, false, lines)
      H.save()
      H.reopen()
      for r, line in ipairs(H.lines()) do -- (row 3 split after its second cell: " y3 |" isn't a row)
        if line:find("| r3", 1, true) then
          local col = line:find("|", line:find("x3", 1, true), true)
          api.nvim_buf_set_text(0, r - 1, col, r - 1, col, { "", "" })
          break
        end
      end
      local said = said_by(function() H.save() end)
      assert.is_nil(said:find("Save failed", 1, true), said)
      local main = zip.read(path, H.MAIN[ext])
      for _, text in ipairs({ "y3", "r4", "x5", "y6", "After." }) do
        assert.truthy(main:find(text, 1, true), text .. " is lost")
      end
    end)

    it("loses no paragraph after a new table whose empty cells go (" .. ext .. ")", function()
      local neo = require("neoffice")
      neo.config.styles = true -- (the styles at the top: the table goes above them)
      local path = H.copy("empty." .. ext)
      H.open(path)
      api.nvim_buf_set_lines(0, 0, -1, false, { "First paragraph.", "", "Last paragraph." })
      H.save()
      H.reopen()
      -- (an empty table typed above it all: its empty cells beyond the first go, as padding)
      api.nvim_buf_set_lines(0, 0, 0, false, { "|     |     |     |", "|-----|-----|-----|", "Right under the table." })
      local ok, said = pcall(said_by, function() H.save() end)
      neo.config.styles = false
      assert(ok, said)
      assert.is_nil(said:find("Save failed", 1, true), said)
      local main = zip.read(path, H.MAIN[ext])
      for _, text in ipairs({ "Right under the table.", "First paragraph.", "Last paragraph." }) do
        assert.truthy(main:find(text, 1, true), text .. " is lost")
      end
    end)

    it("loses no paragraph when a row's last cells are deleted and one is added under the table (" .. ext .. ")", function()
      local path = H.copy("empty." .. ext)
      H.open(path)
      api.nvim_buf_set_lines(0, 0, -1, false, { "Before.", "", "| a | b | c | d |", "|---|---|---|---|", "| e | f | g | h |",
        "| i | j | k | l |", "", "After one.", "", "After two." })
      H.save()
      H.reopen()
      for r, l in ipairs(H.lines()) do
        if l:find("| i", 1, true) then
          api.nvim_buf_set_text(0, r - 1, l:find("k", 1, true) - 2, r - 1, #l, { "|" }) -- (its last two cells deleted)
          api.nvim_buf_set_lines(0, r + 1, r + 1, false, { "New.", "" })
          break
        end
      end
      local said = said_by(function() H.save() end)
      assert.is_nil(said:find("Save failed", 1, true), said)
      local main = zip.read(path, H.MAIN[ext])
      for _, text in ipairs({ "New.", "After one.", "After two.", ">j<" }) do
        assert.truthy(main:find(text, 1, true), text .. " is lost")
      end
      assert.is_nil(main:find(">k<", 1, true))
    end)

    it("loses no text of a box at the start, its lines joined into one (" .. ext .. ")", function()
      local path = H.copy("empty." .. ext)
      H.open(path)
      api.nvim_buf_set_lines(0, 0, -1, false, { "::: {width = 5cm, float = right}", "In the box.", ":::", "", "After the box." })
      H.save()
      H.reopen()
      local lines = H.lines()
      local r = vim.iter(ipairs(lines)):find(function(_, l) return l:find("^:::") end)
      api.nvim_buf_set_lines(0, r - 1, r + 2, false, { table.concat(vim.list_slice(lines, r, r + 2), " ") })
      local said = said_by(function() H.save() end)
      assert.is_nil(said:find("Save failed", 1, true), said)
      local main = zip.read(path, H.MAIN[ext])
      for _, text in ipairs({ "In the box.", "After the box." }) do
        assert.truthy(main:find(text, 1, true), text .. " is lost")
      end
    end)

    it("isn't written at all when a paragraph would be lost (" .. ext .. ")", function()
      local path = H.copy("rich." .. ext)
      local before = H.bytes(path)
      H.open(path)
      H.set(12, "Edited, and then lost.")
      local fmt = H.doc().fmt
      local fill = fmt.fill
      fmt.fill = function(ctx, node, ...) -- (as a bug would: the paragraph taken out once written)
        fill(ctx, node, ...)
        if require("neoffice.xml").text(node):find("then lost", 1, true) then
          require("neoffice.xml").remove(node)
        end
      end
      local ok, said = pcall(said_by, function() H.save() end)
      fmt.fill = fill
      assert(ok, said)
      assert.truthy(said:find("would have lost the text", 1, true), said)
      assert.equals(before, H.bytes(path))
    end)

    it("saves paragraphs put into a new box (" .. ext .. ")", function()
      local path = H.copy("rich." .. ext)
      H.open(path)
      local r = H.row(12)
      local line = H.lines()[r]
      api.nvim_buf_set_lines(0, r - 1, r, false, { "::: {width = 5cm, float = right}", line, ":::" })
      local said = said_by(function() H.save() end)
      assert.is_nil(said:find("Save failed", 1, true), said)
      H.reopen()
      assert.truthy(vim.tbl_contains(H.lines(), line), table.concat(H.lines(), "\n"))
    end)
  end
end)

describe("a heading, title or quote in a style the document hasn't", function()
  local zip = require("neoffice.zip")
  local lines = { "% A new title", "", "### A new heading", "", "> A new quote" }

  it("adds the style to the document's (odt)", function()
    local path = H.copy("empty.odt")
    local text = zip.read(path, "styles.xml") -- (H.copy's edits don't reach styles.xml)
    for _, style in ipairs({ "Heading_20_3", "Title", "Block_20_Text" }) do
      text = text:gsub('<style:style style:name="' .. style .. '".-</style:style>', "")
      assert.is_nil(text:find('style:name="' .. style .. '"', 1, true), style .. " still there")
    end
    zip.write(path, { ["styles.xml"] = (text:gsub('style:parent%-style%-name="Title"', 'style:parent-style-name="Standard"')) })
    H.open(path)
    api.nvim_buf_set_lines(0, 0, -1, false, lines)
    H.save()
    local styles, content = zip.read(path, "styles.xml"), zip.read(path, "content.xml")
    for _, style in ipairs({ "Heading_20_3", "Title", "Quotations" }) do
      local defined = select(2, styles:gsub('style:name="' .. style .. '"', ""))
      assert.equals(1, defined, style .. " defined")
      assert.truthy(content:find('text:style-name="' .. style .. '"', 1, true), style .. " used")
    end
    assert.truthy(styles:find('style:name="Heading_20_3"[^>]-style:default%-outline%-level="3"'))
    H.reopen()
    for _, line in ipairs(lines) do
      assert.truthy(vim.tbl_contains(H.lines(), line), line)
    end
  end)

  it("gives the style an id no style of the document has (docx)", function()
    local path = H.copy("empty.docx", function(name, text)
      if name ~= "word/styles.xml" then
        return text
      end
      text = text:gsub('<w:name w:val="Heading 3" />', '<w:name w:val="Plain three" />')
      text = text:gsub('<w:outlineLvl w:val="2" />', "")
      text = text:gsub('<w:name w:val="Title" />', '<w:name w:val="Plain title" />')
      return (text:gsub('<w:name w:val="Block Text" />', '<w:name w:val="Plain block" />'))
    end)
    H.open(path)
    api.nvim_buf_set_lines(0, 0, -1, false, lines)
    H.save()
    local seen = {}
    for id in zip.read(path, "word/styles.xml"):gmatch('w:styleId="([^"]*)"') do
      assert.is_nil(seen[id], "two styles " .. id)
      seen[id] = true
    end
    H.reopen()
    for _, line in ipairs(lines) do
      assert.truthy(vim.tbl_contains(H.lines(), line), line)
    end
  end)
end)

describe("a comment on text that starts inside a footnote", function()
  local zip = require("neoffice.zip")
  local neo = require("neoffice")
  for _, ext in ipairs({ "docx", "odt" }) do
    it("is written on the footnote, started before it ends (" .. ext .. ")", function()
      local path = H.copy("empty." .. ext)
      H.open(path)
      api.nvim_buf_set_lines(0, 0, -1, false, { "Words with a note.^[the note]", "", "Next." })
      H.save()
      H.reopen()
      local line = H.lines()[1]
      api.nvim_win_set_cursor(0, { 1, #line - 4 }) -- (in the note's text, to the end of the line)
      vim.cmd("normal! v$\27")
      require("neoffice.comments").add(true, "On the note")
      H.save()
      local main = zip.read(path, H.MAIN[ext])
      if ext == "docx" then
        local s, e, r = main:find("commentRangeStart", 1, true), main:find("commentRangeEnd", 1, true), main:find("commentReference", 1, true)
        assert(s and e and r, main)
        assert.is_true(s < e and e < r, "started before it ends")
      else
        local s, e = main:find("<office:annotation ", 1, true) or main:find("<office:annotation>", 1, true), main:find("<office:annotation-end", 1, true)
        assert(s, main)
        assert.is_true(not e or s < e, "started before it ends")
      end
      H.reopen()
      assert.equals(1, #H.comments())
      assert.equals("On the note", H.comments()[1].text)
      assert.equals(1, #api.nvim_buf_get_extmarks(0, neo.ns.comment, 0, -1, {}))
    end)
  end
end)

describe("a comment in a table", function()
  local neo = require("neoffice")
  local function said_by(fn)
    local said, notify = {}, vim.notify
    vim.notify = function(msg) said[#said + 1] = msg end
    local ok, err = pcall(fn)
    vim.notify = notify
    assert(ok, err)
    return table.concat(said, "\n")
  end
  local function covered()
    local m = api.nvim_buf_get_extmarks(0, neo.ns.comment, 0, -1, { details = true })
    assert.equals(1, #m)
    return table.concat(api.nvim_buf_get_text(0, m[1][2], m[1][3], m[1][4].end_row, m[1][4].end_col, {}), "\n")
  end
  -- (where it starts on the row "| one   | two  |", how far it goes: what it then covers)
  local cases = {
    { "inside a cell", 3, 2, "ne" },
    { "across the row, from its first |", 0, 30, "one   | two" },
    { "from the spaces and | before a cell", 7, 3, "t" },
    { "at a point", 4, 0, "" },
  }
  for _, ext in ipairs({ "docx", "odt" }) do
    for _, case in ipairs(cases) do
      it("is written whole, " .. case[1] .. " (" .. ext .. ")", function()
        local path = H.copy("empty." .. ext)
        H.open(path)
        api.nvim_buf_set_lines(0, 0, -1, false, { "Before.", "", "| alpha | beta |", "|-------|------|", "| one   | two  |", "| three | four |", "", "After." })
        H.save()
        H.reopen()
        local row
        for k, l in ipairs(H.lines()) do
          if l:find("one", 1, true) then
            row = k
          end
        end
        api.nvim_win_set_cursor(0, { row, case[2] })
        if case[3] > 0 then
          vim.cmd("normal! v" .. case[3] .. "l\27")
        end
        require("neoffice.comments").add(case[3] > 0, "In the table")
        local said = said_by(function() H.save() end)
        assert.is_nil(said:find("Save failed", 1, true), said)
        H.reopen()
        assert.equals(1, #H.comments())
        assert.equals("In the table", H.comments()[1].text)
        assert.equals(case[4], covered())
      end)
    end
  end

  for _, ext in ipairs({ "docx", "odt" }) do
    it("is written on the next paragraph when it is on a line that is no paragraph's (" .. ext .. ")", function()
      local path = H.copy("empty." .. ext)
      H.open(path)
      api.nvim_buf_set_lines(0, 0, -1, false, { "Before.", "", "{rules = header}", "| a | b |", "|---|---|", "| 1 | 2 |" })
      H.save()
      H.reopen()
      local row
      for k, l in ipairs(H.lines()) do
        if l:find("^{") then
          row = k
        end
      end
      assert(row, table.concat(H.lines(), "\n"))
      api.nvim_win_set_cursor(0, { row, 2 })
      vim.cmd("normal! v4l\27")
      require("neoffice.comments").add(true, "On the options' line")
      local said = said_by(function() H.save() end)
      assert.is_nil(said:find("Save failed", 1, true), said)
      H.reopen()
      assert.equals(1, #H.comments())
      assert.equals("On the options' line", H.comments()[1].text)
    end)
  end

  it("is written whole when it starts on the table's line of dashes (docx)", function()
    local path = H.copy("empty.docx")
    H.open(path)
    api.nvim_buf_set_lines(0, 0, -1, false, { "Before.", "", "| alpha | beta |", "|-------|------|", "| one   | two  |", "", "After." })
    H.save()
    H.reopen()
    local dashes
    for k, l in ipairs(H.lines()) do
      if l:find("^|%-") then
        dashes = k
      end
    end
    api.nvim_win_set_cursor(0, { dashes, 3 })
    vim.cmd("normal! vj2l\27")
    require("neoffice.comments").add(true, "From the dashes")
    local said = said_by(function() H.save() end)
    assert.is_nil(said:find("Save failed", 1, true), said)
    H.reopen()
    assert.equals(1, #H.comments())
    assert.truthy(covered():find("^one") or covered():find("^o"), covered())
  end)
end)

describe("a comment at a point on a table's line of dashes", function()
  for _, ext in ipairs({ "docx", "odt" }) do
    it("is written on the cell after it (" .. ext .. ")", function()
      local path = H.copy("empty." .. ext)
      H.open(path)
      api.nvim_buf_set_lines(0, 0, -1, false, { "Before.", "", "| alpha | beta |", "|-------|------|", "| one   | two  |", "", "After." })
      H.save()
      H.reopen()
      local dashes
      for k, l in ipairs(H.lines()) do
        if l:find("^|%-") then
          dashes = k
        end
      end
      api.nvim_win_set_cursor(0, { dashes, 3 })
      require("neoffice.comments").add(false, "On the dashes")
      local said, notify = {}, vim.notify
      vim.notify = function(msg) said[#said + 1] = msg end
      local ok, err = pcall(H.save)
      vim.notify = notify
      assert(ok, err)
      assert.is_nil(table.concat(said, "\n"):find("Save failed", 1, true), table.concat(said, "\n"))
      H.reopen()
      assert.equals(1, #H.comments())
      assert.equals("On the dashes", H.comments()[1].text)
    end)
  end
end)

describe("a tracked change's id", function()
  local zip = require("neoffice.zip")
  local neo = require("neoffice")
  local function ids(path)
    local seen, twice = {}, {}
    local main = zip.read(path, "word/document.xml")
    for tag, id in main:gmatch('<w:(%a+) w:id="(%d+)"') do
      if tag == "ins" or tag == "del" or tag == "rPrChange" or tag == "pPrChange" then
        local key = tag .. " " .. id
        if seen[key] then
          twice[#twice + 1] = key
        end
        seen[key] = true
      end
    end
    return twice, seen
  end

  it("is given again to a copy of a paragraph that has it (written as it was, its id and all)", function()
    local D = 'w:author="A" w:date="2024-01-01T00:00:00Z"'
    local path = H.copy("empty.docx", function(name, text)
      if name ~= "word/document.xml" then
        return text
      end
      local p = '<w:p><w:pPr><w:rPr><w:ins w:id="41" ' .. D .. '/></w:rPr></w:pPr><w:ins w:id="42" ' .. D .. '><w:r><w:t>Inserted.</w:t></w:r></w:ins></w:p>'
      return (text:gsub("<w:body>", "<w:body>" .. p .. "<w:p><w:r><w:t>Other.</w:t></w:r></w:p>", 1))
    end)
    H.open(path)
    local doc = H.doc()
    local xml = require("neoffice.xml")
    local body = xml.find(doc.ctx.root, "w:body")
    local first = xml.find(body, "w:p")
    xml.append(body, xml.copy(first), 2) -- (as a copy written as the paragraph was would be)
    doc.fmt.unique_drawings(doc.ctx)
    local seen = {}
    for _, el in ipairs(xml.collect(body, { ["w:ins"] = true })) do
      local id = xml.attr(el, "w:id")
      assert.is_nil(seen[id], "id " .. id .. " twice")
      seen[id] = true
    end
    assert.truthy(seen["41"] and seen["42"], "the paragraph's own ids kept")
  end)

  it("isn't written twice when a paragraph with a tracked format change is copied", function()
    local path = H.copy("empty.docx", function(name, text)
      if name ~= "word/document.xml" then
        return text
      end
      local p = '<w:p><w:r><w:rPr><w:b/><w:rPrChange w:id="7" w:author="A" w:date="2024-01-01T00:00:00Z"><w:rPr/></w:rPrChange></w:rPr><w:t>Made bold.</w:t></w:r></w:p>'
      return (text:gsub("<w:body>", "<w:body>" .. p .. "<w:p><w:r><w:t>Other.</w:t></w:r></w:p>", 1))
    end)
    H.open(path)
    local lines = H.lines()
    for _, l in ipairs(lines) do
      if l:find("Made bold", 1, true) then
        api.nvim_buf_set_lines(0, #lines, #lines, false, { "", l })
        break
      end
    end
    H.save()
    local twice, seen = ids(path)
    assert.truthy(seen["rPrChange 7"], "the change is kept")
    assert.same({}, twice)
  end)
end)

describe("a table's line of dashes", function()
  local neo = require("neoffice")
  local zip = require("neoffice.zip")
  local function open_table(ext, rows)
    local path = H.copy("empty." .. ext)
    H.open(path)
    local lines = { "First.", "" }
    vim.list_extend(lines, rows)
    vim.list_extend(lines, { "", "Last." })
    api.nvim_buf_set_lines(0, 0, -1, false, lines)
    H.save()
    H.reopen()
    return path
  end
  local function row_of(pattern)
    for k, l in ipairs(H.lines()) do
      if l:find(pattern) then
        return k
      end
    end
  end
  local function no_dashes_written(path, ext)
    local main = zip.read(path, H.MAIN[ext])
    assert.is_nil(main:find("%-%-%-"), "dashes written as text")
  end
  for _, ext in ipairs({ "docx", "odt" }) do
    it("isn't written as a row when the first row is deleted (" .. ext .. ")", function()
      local path = open_table(ext, { "| a | b |", "|---|---|", "| 1 | 2 |", "| 3 | 4 |" })
      local r = row_of("^| +a ")
      api.nvim_buf_set_lines(0, r - 1, r, false, {})
      H.save()
      no_dashes_written(path, ext)
      H.reopen()
      assert.equals(1, #vim.tbl_filter(function(l) return l:find("^|%-") ~= nil end, H.lines()))
      assert.truthy(H.lines()[row_of("^|%-") - 1]:find("1", 1, true))
    end)

    it("isn't written as a row when a row is put between it and the first row (" .. ext .. ")", function()
      local path = open_table(ext, { "| a | b |", "|---|---|", "| 1 | 2 |" })
      local r = row_of("^| +a ")
      api.nvim_buf_set_lines(0, r, r, false, { "| x | y |" })
      H.save()
      no_dashes_written(path, ext)
      H.reopen()
      assert.truthy(row_of("^| +x "), table.concat(H.lines(), "\n"))
    end)

    it("leaves a row of dashes for values missing a row, the first row deleted (" .. ext .. ")", function()
      local path = open_table(ext, { "| a | b |", "|---|---|", "| - | - |", "| 1 | 2 |" })
      local r = row_of("^| +a ")
      api.nvim_buf_set_lines(0, r - 1, r, false, {})
      H.save()
      no_dashes_written(path, ext)
      H.reopen()
      assert.truthy(row_of("^| +%- +| +%- +|"), table.concat(H.lines(), "\n"))
      assert.truthy(row_of("^| +1 "), table.concat(H.lines(), "\n"))
    end)

    it("isn't written as a row when the first row is retyped with changes tracked (" .. ext .. ")", function()
      local path = open_table(ext, { "| a | b |", "|---|---|", "| 1 | 2 |" })
      neo.config.author = "Reviewer"
      H.doc().tracking = true
      require("neoffice.changes").snapshot(api.nvim_get_current_buf())
      local r = row_of("^| +a ")
      api.nvim_buf_set_lines(0, r - 1, r, false, { H.lines()[r] .. " extra |" })
      require("neoffice.changes").after_edit(api.nvim_get_current_buf())
      H.save()
      no_dashes_written(path, ext)
    end)
  end
end)

describe("a caption added with changes tracked", function()
  local neo = require("neoffice")
  for _, ext in ipairs({ "docx", "odt" }) do
    it("rejected, leaves nothing of it, its word and number neither (" .. ext .. ")", function()
      local path = H.copy("empty." .. ext)
      H.open(path)
      api.nvim_buf_set_lines(0, 0, -1, false, { "First.", "", "Last." })
      H.save()
      H.reopen()
      neo.config.author = "Reviewer"
      H.doc().tracking = true
      require("neoffice.changes").snapshot(api.nvim_get_current_buf())
      api.nvim_buf_set_lines(0, 2, 2, false, { "Table: A caption", "" })
      require("neoffice.changes").after_edit(api.nvim_get_current_buf())
      H.save()
      local main = require("neoffice.zip").read(path, H.MAIN[ext])
      if ext == "docx" then -- (its number in the insertion too: a field in its runs, a w:fldSimple can't be)
        assert.is_nil(main:find("<w:fldSimple", 1, true), main)
        assert.truthy(main:find("<w:ins [^>]*>%s*<w:r>%s*<w:fldChar"), main)
      end
      H.open(path)
      assert.truthy(vim.tbl_contains(H.lines(), "Table: A caption"), table.concat(H.lines(), "\n"))
      require("neoffice.changes").resolve("reject", true)
      H.save()
      H.open(path)
      local text = table.concat(H.lines(), "\n")
      assert.is_nil(text:find("Table", 1, true), text)
      assert.is_nil(text:find("⟦", 1, true), text)
      local written = require("neoffice.zip").read(path, H.MAIN[ext])
      assert.is_nil(written:find(ext == "docx" and "SEQ Table" or "<text:sequence ", 1, true), "its number still there")
    end)

    it("is a caption again, typed where a paragraph was deleted (" .. ext .. ")", function()
      local path = H.copy("empty." .. ext)
      H.open(path)
      api.nvim_buf_set_lines(0, 0, -1, false, { "First.", "", "Old paragraph.", "", "Last." })
      H.save()
      H.reopen()
      neo.config.author = "Reviewer"
      H.doc().tracking = true
      require("neoffice.changes").snapshot(api.nvim_get_current_buf())
      api.nvim_buf_set_lines(0, 2, 3, false, { "Table: A caption" })
      require("neoffice.changes").after_edit(api.nvim_get_current_buf())
      H.save()
      H.open(path)
      local text = table.concat(H.lines(), "\n")
      -- (a caption: its word, not its number written; in an .odt, the paragraph deleted is in it, as
      -- LibreOffice writes it, shown deleted)
      assert.truthy(text:find("\nTable: ", 1, true), text)
      assert.truthy(text:find("A caption", 1, true), text)
      assert.is_nil(text:find("⟦", 1, true), text)
    end)

    it("is a caption again when reopened, its text inserted (" .. ext .. ")", function()
      local path = H.copy("empty." .. ext)
      H.open(path)
      api.nvim_buf_set_lines(0, 0, -1, false, { "Some words.", "", "Last." })
      H.save()
      H.reopen()
      neo.config.author = "Reviewer"
      H.doc().tracking = true
      require("neoffice.changes").snapshot(api.nvim_get_current_buf())
      api.nvim_buf_set_lines(0, 0, 0, false, { "Table: A caption", "" })
      require("neoffice.changes").after_edit(api.nvim_get_current_buf())
      H.save()
      H.reopen()
      assert.equals("Table: A caption", H.lines()[1])
      local inserted = {}
      for _, m in ipairs(api.nvim_buf_get_extmarks(0, neo.ns.change, 0, -1, { details = true })) do
        if H.doc().change[m[1]].type == "insert" then
          inserted[#inserted + 1] = table.concat(api.nvim_buf_get_text(0, m[2], m[3], m[4].end_row, m[4].end_col, {}), "\n")
        end
      end
      assert.truthy(vim.iter(inserted):any(function(t) return t:find("^A caption") ~= nil end), vim.inspect(inserted))
    end)
  end
end)

describe("a box", function()
  local function said_by(fn)
    local said, notify = {}, vim.notify
    vim.notify = function(msg) said[#said + 1] = msg end
    local ok, err = pcall(fn)
    vim.notify = notify
    assert(ok, err)
    return table.concat(said, "\n")
  end
  local doc = { "Before.", "", "::: {width = 5cm}", "In the box.", ":::", "", "# A heading", "", "After the box.", "", "| a | b |", "|---|---|", "| 1 | 2 |", "", "The end." }
  for _, ext in ipairs({ "docx", "odt" }) do
    it("never closed (its line of colons typed in) isn't saved, and says where it starts (" .. ext .. ")", function()
      local path = H.copy("empty." .. ext)
      H.open(path)
      api.nvim_buf_set_lines(0, 0, -1, false, doc)
      H.save()
      H.reopen()
      local before = H.bytes(path)
      for k, l in ipairs(H.lines()) do
        if l == ":::" then
          api.nvim_buf_set_text(0, k - 1, 2, k - 1, 2, { "> quote" })
        end
      end
      local said = said_by(function() H.save() end)
      assert.truthy(said:find("never closed", 1, true), said)
      assert.equals(before, H.bytes(path))
    end)

    it("never closed right after opening (before all is read) isn't saved either, without an error (" .. ext .. ")", function()
      local path = H.copy("empty." .. ext)
      H.open(path)
      local long = vim.list_extend({}, doc)
      for n = 1, 300 do
        vim.list_extend(long, { "", "Paragraph " .. n .. "." })
      end
      api.nvim_buf_set_lines(0, 0, -1, false, long)
      H.save()
      local before = H.bytes(path)
      vim.cmd("silent! %bwipeout!")
      vim.cmd("edit " .. vim.fn.fnameescape(path)) -- (no waiting: its paragraphs' fingerprints still to take)
      for k, l in ipairs(api.nvim_buf_get_lines(0, 0, -1, false)) do
        if l == ":::" then
          api.nvim_buf_set_text(0, k - 1, 2, k - 1, 2, { "> quote" })
          break
        end
      end
      local said = said_by(function() H.save() end)
      assert.truthy(said:find("never closed", 1, true), said)
      assert.equals(before, H.bytes(path))
    end)

    it("with a table in it isn't saved (the table would be taken apart) (" .. ext .. ")", function()
      local path = H.copy("empty." .. ext)
      H.open(path)
      api.nvim_buf_set_lines(0, 0, -1, false, { "Before.", "", "| a | b |", "|---|---|", "| 1 | 2 |", "", "After." })
      H.save()
      H.reopen()
      local before = H.bytes(path)
      local first, last
      for k, l in ipairs(H.lines()) do
        if l:find("^|") then
          first, last = first or k, k
        end
      end
      api.nvim_buf_set_lines(0, last, last, false, { ":::" })
      api.nvim_buf_set_lines(0, first - 1, first - 1, false, { "::: {width = 5cm}" })
      local said = said_by(function() H.save() end)
      assert.truthy(said:find("a table can't go in a box", 1, true), said)
      assert.equals(before, H.bytes(path))
    end)
  end
end)

describe("table rows deleted with changes tracked", function()
  local neo = require("neoffice")
  for _, ext in ipairs({ "docx", "odt" }) do
    it("keep each cell's own text, deleted (" .. ext .. ")", function()
      local path = H.copy("empty." .. ext)
      H.open(path)
      api.nvim_buf_set_lines(0, 0, -1, false, { "Before.", "", "| h1 | h2 | h3 |", "|---|---|---|",
        "| alpha | beta | gamma |", "| delta | epsilon | zeta |", "| eta | theta | iota |", "", "After." })
      H.save()
      H.reopen()
      neo.config.author = "Reviewer"
      H.doc().tracking = true
      require("neoffice.changes").snapshot(api.nvim_get_current_buf())
      local r
      for k, l in ipairs(H.lines()) do
        if l:find("alpha", 1, true) then
          r = k
        end
      end
      api.nvim_buf_set_lines(0, r - 1, r + 1, false, {}) -- (the rows of alpha and delta)
      require("neoffice.changes").after_edit(api.nvim_get_current_buf())
      H.save()
      H.reopen()
      local text = table.concat(H.lines(), "\n")
      for _, word in ipairs({ "alpha", "beta", "gamma", "delta", "epsilon", "zeta", "eta", "theta", "iota" }) do
        assert.truthy(text:find("| " .. word .. " ", 1, true) or text:find(" " .. word .. " |", 1, true), word .. " lost:\n" .. text)
      end
      local deleted = {}
      for _, m in ipairs(api.nvim_buf_get_extmarks(0, neo.ns.change, 0, -1, { details = true })) do
        if H.doc().change[m[1]].type == "delete" then
          deleted[#deleted + 1] = table.concat(api.nvim_buf_get_text(0, m[2], m[3], m[4].end_row, m[4].end_col, {}), "\n")
        end
      end
      local all = table.concat(deleted, " ")
      for _, word in ipairs({ "alpha", "beta", "gamma", "delta", "epsilon", "zeta" }) do
        assert.truthy(all:find(word, 1, true), word .. " not shown deleted: " .. all)
      end
    end)
  end
end)

describe("tracked changes and comments made in one save", function()
  local neo = require("neoffice")
  it("each have an id of their own, whatever chance gives (odt)", function()
    local path = H.copy("empty.odt")
    H.open(path)
    api.nvim_buf_set_lines(0, 0, -1, false, { "First paragraph with alpha in it.", "", "Second paragraph with omega in it.", "", "Third." })
    H.save()
    H.reopen()
    neo.config.author = "Reviewer"
    H.doc().tracking = true
    require("neoffice.changes").snapshot(api.nvim_get_current_buf())
    H.replace(1, "alpha ", "")
    H.replace(2, "omega ", "")
    require("neoffice.changes").after_edit(api.nvim_get_current_buf())
    api.nvim_win_set_cursor(0, { 1, 0 })
    require("neoffice.comments").add(false, "One")
    api.nvim_win_set_cursor(0, { H.row(3), 0 })
    require("neoffice.comments").add(false, "Two")
    local random = math.random
    math.random = function() return 7 end -- (the same number every time)
    local ok, err = pcall(H.save)
    math.random = random
    assert(ok, err)
    H.open(path) -- (as written by that save: not saved again)
    local deleted = {}
    for _, m in ipairs(api.nvim_buf_get_extmarks(0, neo.ns.change, 0, -1, { details = true })) do
      if H.doc().change[m[1]].type == "delete" then
        deleted[#deleted + 1] = table.concat(api.nvim_buf_get_text(0, m[2], m[3], m[4].end_row, m[4].end_col, {}), "\n")
      end
    end
    table.sort(deleted)
    assert.same({ "alpha ", "omega " }, deleted)
    local texts = vim.tbl_map(function(c) return c.text end, H.comments())
    table.sort(texts)
    assert.same({ "One", "Two" }, texts)
  end)
end)

describe("a deletion of whole paragraphs, changes tracked", function()
  local neo = require("neoffice")
  local zip = require("neoffice.zip")
  local changes = require("neoffice.changes")
  -- (as LibreOffice writes it: the paragraphs gone from the text, in the deletion; the next one
  -- holds where they were)
  local function libreoffice_deletion()
    local info = "<office:change-info><dc:creator>A</dc:creator><dc:date>2024-01-01T00:00:00</dc:date></office:change-info>"
    local region = '<text:tracked-changes><text:changed-region text:id="ct1"><text:deletion>' .. info
      .. '<text:p text:style-name="Standard">Beta two.</text:p><text:p text:style-name="Standard"/><text:p text:style-name="Standard">Gamma three.</text:p>'
      .. '<text:p text:style-name="Standard"/></text:deletion></text:changed-region></text:tracked-changes>'
    return H.copy("empty.odt", function(name, text)
      if name ~= "content.xml" then
        return text
      end
      return (text:gsub("<office:text>(.-)</office:text>", function()
        return "<office:text>" .. region .. '<text:p text:style-name="Standard">Alpha one.</text:p>'
          .. '<text:p text:style-name="Standard"><text:change text:change-id="ct1"/>Epsilon five.</text:p></office:text>'
      end, 1))
    end)
  end

  it("by LibreOffice keeps what it deleted when the paragraph is edited (odt)", function()
    local path = libreoffice_deletion()
    H.open(path)
    local r
    for k, l in ipairs(H.lines()) do
      if l:find("Epsilon five.", 1, true) then
        r = k
      end
    end
    local line = H.lines()[r]
    assert.is_nil(line:find("\\¶", 1, true), line) -- (a break, not the character)
    api.nvim_buf_set_text(0, r - 1, #line - 1, r - 1, #line, { "!" })
    H.save()
    local main = zip.read(path, "content.xml")
    assert.is_nil(main:find("¶", 1, true), "the breaks written as characters")
    assert.truthy(main:find("<text:p[^>]*>Beta two.</text:p>"), main)
    assert.truthy(main:find("<text:p[^>]*>Gamma three.</text:p>"), main)
  end)

  it("by LibreOffice, rejected, brings its paragraphs back as paragraphs (odt)", function()
    local path = libreoffice_deletion()
    H.open(path)
    changes.resolve("reject", true)
    H.save()
    H.open(path)
    -- (Beta, an empty paragraph, Gamma: deleted; then the break that joined them to Epsilon)
    assert.same({ "Alpha one.", "", "Beta two.", "", "", "Gamma three.", "", "Epsilon five." }, H.lines())
  end)

  for _, ext in ipairs({ "docx", "odt" }) do
    it("leaves no empty paragraphs once accepted (" .. ext .. ")", function()
      local path = H.copy("empty." .. ext)
      H.open(path)
      api.nvim_buf_set_lines(0, 0, -1, false, { "Alpha one.", "", "Beta two.", "", "", "", "Gamma three.", "", "Delta four.", "", "Epsilon five." })
      H.save()
      H.reopen()
      neo.config.author = "Reviewer"
      H.doc().tracking = true
      changes.snapshot(api.nvim_get_current_buf())
      api.nvim_buf_set_lines(0, 2, 9, false, {}) -- (Beta, the empty paragraph, Gamma, Delta)
      changes.after_edit(api.nvim_get_current_buf())
      H.save()
      if ext == "odt" then -- (as LibreOffice writes it: the paragraphs in the deletion, not left empty)
        local main = zip.read(path, "content.xml")
        assert.is_nil(main:find("<text:p[^>]*><text:change [^>]*/></text:p>"), main)
      end
      H.open(path)
      changes.resolve("accept", true)
      assert.same({ "Alpha one.", "", "Epsilon five." }, H.lines())
    end)

    it("rejected, brings back the paragraphs as they were (" .. ext .. ")", function()
      local path = H.copy("empty." .. ext)
      H.open(path)
      local was = { "Alpha one.", "", "Beta two.", "", "", "", "Gamma three.", "", "Delta four.", "", "Epsilon five." }
      api.nvim_buf_set_lines(0, 0, -1, false, was)
      H.save()
      H.reopen()
      neo.config.author = "Reviewer"
      H.doc().tracking = true
      changes.snapshot(api.nvim_get_current_buf())
      api.nvim_buf_set_lines(0, 2, 9, false, {})
      changes.after_edit(api.nvim_get_current_buf())
      H.save()
      H.open(path)
      changes.resolve("reject", true)
      assert.same(was, H.lines())
    end)

    it("rejected, doesn't join a paragraph to the next when two were typed again as one (" .. ext .. ")", function()
      local path = H.copy("empty." .. ext)
      H.open(path)
      api.nvim_buf_set_lines(0, 0, -1, false, { "First part.", "", "Second part.", "", "Last." })
      H.save()
      H.reopen()
      neo.config.author = "Reviewer"
      H.doc().tracking = true
      changes.snapshot(api.nvim_get_current_buf())
      api.nvim_buf_set_lines(0, 0, 3, false, { "First part. Second part." })
      changes.after_edit(api.nvim_get_current_buf())
      H.save()
      H.open(path)
      changes.resolve("reject", true)
      local text = vim.tbl_map(function(l) return vim.trim(l) end, H.lines())
      assert.same({ "First part.", "", "Second part.", "", "Last." }, text)
    end)
  end

  it("of the last line of a paragraph of two keeps the paragraph's break (docx)", function()
    local path = H.copy("empty.docx")
    H.open(path)
    api.nvim_buf_set_lines(0, 0, -1, false, { "One line,", "and another.", "", "Next paragraph." })
    H.save() -- (not opened again: there, a paragraph is on a line of its own)
    neo.config.author = "Reviewer"
    H.doc().tracking = true
    changes.snapshot(api.nvim_get_current_buf())
    local r
    for k, l in ipairs(H.lines()) do
      if l:find("and another", 1, true) then
        r = k
      end
    end
    api.nvim_buf_set_lines(0, r - 1, r, false, {})
    changes.after_edit(api.nvim_get_current_buf())
    H.save()
    H.open(path)
    changes.resolve("accept", true)
    local text = vim.tbl_map(function(l) return vim.trim(l) end, H.lines())
    assert.same({ "One line,", "", "Next paragraph." }, text)
  end)
end)

describe("a comment whose reference is in a deletion Word made", function()
  local zip = require("neoffice.zip")
  local neo = require("neoffice")
  it("is referenced once when the deletion is in the paragraph after its text (docx)", function()
    local D = 'w:author="A" w:date="2024-01-01T00:00:00Z"'
    local path = H.copy("rich.docx", function(name, text)
      if name ~= "word/document.xml" then
        return text
      end
      local ref = '<w:r><w:rPr><w:rStyle w:val="CommentReference" /></w:rPr><w:commentReference w:id="0" /></w:r>'
      text = text:gsub(vim.pesc(ref), "", 1)
      -- (its reference in a deletion at the start of the next paragraph)
      return (text:gsub('(<w:commentRangeEnd w:id="0" />.-</w:p><w:p>.-</w:pPr>)', '%1<w:del w:id="78" ' .. D .. '><w:r><w:delText>away</w:delText></w:r>' .. ref .. '</w:del>', 1))
    end)
    H.open(path)
    neo.config.author = "Reviewer"
    H.doc().tracking = true
    require("neoffice.changes").snapshot(api.nvim_get_current_buf())
    H.replace(3, "and more.", "and more, edited.")
    H.replace(4, "Text with", "Text, edited, with")
    require("neoffice.changes").after_edit(api.nvim_get_current_buf())
    local said, notify = {}, vim.notify
    vim.notify = function(msg) said[#said + 1] = msg end
    local ok, err = pcall(H.save)
    vim.notify = notify
    assert(ok, err)
    assert.is_nil(table.concat(said, "\n"):find("Save failed", 1, true), table.concat(said, "\n"))
    local main = zip.read(path, "word/document.xml")
    assert.equals(1, select(2, main:gsub('<w:commentReference w:id="0"', "")), "references")
    H.open(path)
    assert.equals(1, #H.comments())
  end)

  it("is referenced once when its paragraph is written again (docx)", function()
    local D = 'w:author="A" w:date="2024-01-01T00:00:00Z"'
    local path = H.copy("rich.docx", function(name, text)
      if name ~= "word/document.xml" then
        return text
      end
      -- (its range on text that stays, its reference in a deletion after it)
      local ref = '<w:r><w:rPr><w:rStyle w:val="CommentReference" /></w:rPr><w:commentReference w:id="0" /></w:r>'
      text = text:gsub(vim.pesc(ref), "", 1)
      return (text:gsub('(<w:commentRangeEnd w:id="0" />)', '%1<w:del w:id="77" ' .. D .. '><w:r><w:delText>gone</w:delText></w:r>' .. ref .. '</w:del>', 1))
    end)
    H.open(path)
    neo.config.author = "Reviewer"
    H.doc().tracking = true
    require("neoffice.changes").snapshot(api.nvim_get_current_buf())
    H.replace(3, "and more.", "and more, edited.")
    require("neoffice.changes").after_edit(api.nvim_get_current_buf())
    H.save()
    local main = zip.read(path, "word/document.xml")
    assert.truthy(main:find("edited", 1, true), "not saved")
    assert.equals(1, select(2, main:gsub('<w:commentReference w:id="0"', "")), "references")
    assert.equals(1, select(2, main:gsub('<w:commentRangeStart w:id="0"', "")), "starts")
    H.open(path)
    assert.equals(1, #H.comments())
  end)
end)

describe("table rows with changes tracked", function()
  local neo = require("neoffice")
  local zip = require("neoffice.zip")
  local changes = require("neoffice.changes")
  local base = { "Before.", "", "| h1 | h2 |", "|---|---|", "| alpha | beta |", "| gamma | delta |", "| eta | theta |", "", "After." }
  local function tracked(ext, edit)
    local path = H.copy("empty." .. ext)
    H.open(path)
    api.nvim_buf_set_lines(0, 0, -1, false, base)
    H.save()
    H.reopen()
    neo.config.author = "Reviewer"
    H.doc().tracking = true
    changes.snapshot(api.nvim_get_current_buf())
    edit()
    changes.after_edit(api.nvim_get_current_buf())
    H.save()
    return path
  end
  local function row(pattern)
    for k, l in ipairs(H.lines()) do
      if l:find(pattern) then
        return k
      end
    end
  end
  local function row_marks(path, tag)
    local n = 0
    for trpr in zip.read(path, "word/document.xml"):gmatch("<w:trPr>(.-)</w:trPr>") do
      n = n + (trpr:find("<w:" .. tag .. " ") and 1 or 0)
    end
    return n
  end

  it("deleted are marked deleted, as Word marks them: accepting takes them away (docx)", function()
    local path = tracked("docx", function()
      local r = row("alpha")
      api.nvim_buf_set_lines(0, r - 1, r + 1, false, {})
    end)
    assert.equals(2, row_marks(path, "del"))
  end)

  it("all deleted: the table is all marked deleted (docx)", function()
    local path = tracked("docx", function()
      api.nvim_buf_set_lines(0, row("^| +h1") - 1, row("theta"), false, {})
    end)
    assert.equals(4, row_marks(path, "del"))
  end)

  it("inserted is marked inserted: rejecting takes it away (docx)", function()
    local path = tracked("docx", function()
      local r = row("gamma")
      api.nvim_buf_set_lines(0, r, r, false, { "| new1 | new2 |" })
    end)
    assert.equals(1, row_marks(path, "ins"))
  end)

  it("inserted is a tracked insertion, each cell's of its own (odt)", function()
    local path = tracked("odt", function()
      local r = row("gamma")
      api.nvim_buf_set_lines(0, r, r, false, { "| new1 | new2 |" })
    end)
    local main = zip.read(path, "content.xml")
    local starts, ids = 0, {}
    for id in main:gmatch('<text:change%-start text:change%-id="([^"]+)"') do
      starts = starts + 1
      ids[id] = true
    end
    assert.equals(2, starts, "a change in each cell")
    assert.equals(2, vim.tbl_count(ids), "each its own")
    assert.truthy(main:find("<text:insertion>", 1, true))
  end)

  it("whose cells' text only was deleted aren't marked deleted (docx)", function()
    local path = tracked("docx", function()
      local r = row("alpha")
      local line = H.lines()[r]
      local a, b = line:find("alpha", 1, true)
      api.nvim_buf_set_text(0, r - 1, a - 1, r - 1, b, { "" })
      local c, d = H.lines()[r]:find("beta", 1, true)
      api.nvim_buf_set_text(0, r - 1, c - 1, r - 1, d, { "" })
    end)
    assert.equals(0, row_marks(path, "del"))
  end)

  for _, ext in ipairs({ "docx", "odt" }) do
    it("deleted, then accepted here, leave the table as a table (" .. ext .. ")", function()
      local path = tracked(ext, function()
        api.nvim_buf_set_lines(0, row("^| +h1") - 1, row("theta"), false, {})
      end)
      H.open(path)
      changes.resolve("accept", true)
      local lines = H.lines()
      assert.equals("After.", lines[#lines], table.concat(lines, "\n"))
      for _, l in ipairs(lines) do
        if l:match("^|") then
          assert.truthy(l:match("|%s*$"), l) -- (each row its own line, ending with its "|")
        end
      end
      H.save()
      assert.truthy(zip.read(path, H.MAIN[ext]):find("After.", 1, true))
    end)

    it("deleted and accepted at once, before saving, leave the table as a table (" .. ext .. ")", function()
      local path = H.copy("empty." .. ext)
      H.open(path)
      api.nvim_buf_set_lines(0, 0, -1, false, base)
      H.save()
      H.reopen()
      neo.config.author = "Reviewer"
      H.doc().tracking = true
      changes.snapshot(api.nvim_get_current_buf())
      local r = row("alpha")
      api.nvim_win_set_cursor(0, { r, 3 })
      vim.cmd("normal! d$") -- (the rest of the row from inside its first cell: not its line's end)
      changes.after_edit(api.nvim_get_current_buf())
      changes.resolve("accept", true)
      for _, l in ipairs(H.lines()) do
        if l:match("^|") then
          assert.truthy(l:match("|%s*$"), table.concat(H.lines(), "\n"))
        end
      end
      assert.truthy(row("^| +gamma +| +delta +|$"), table.concat(H.lines(), "\n"))
      assert.equals("After.", H.lines()[#H.lines()])
    end)

    it("inserted, then rejected here, leave the rows after it as they were (" .. ext .. ")", function()
      local path = tracked(ext, function()
        local r = row("gamma")
        api.nvim_buf_set_lines(0, r, r, false, { "| new1 | new2 |" })
      end)
      H.open(path)
      changes.resolve("reject", true)
      assert.truthy(row("^| +eta +| +theta +|$"), table.concat(H.lines(), "\n"))
      assert.is_nil(table.concat(H.lines(), "\n"):find("new1", 1, true))
    end)
  end
end)

describe("a caption deleted with changes tracked", function()
  local neo = require("neoffice")
  local zip = require("neoffice.zip")
  local changes = require("neoffice.changes")
  for _, ext in ipairs({ "docx", "odt" }) do
    local function deleted()
      local path = H.copy("empty." .. ext)
      H.open(path)
      api.nvim_buf_set_lines(0, 0, -1, false, { "Before.", "", "Table: A caption here", "", "| a | b |", "|---|---|", "| 1 | 2 |", "", "After." })
      H.save()
      H.reopen()
      neo.config.author = "Reviewer"
      H.doc().tracking = true
      changes.snapshot(api.nvim_get_current_buf())
      local first, last
      for k, l in ipairs(H.lines()) do
        if l:find("^Table") then
          first = k
        end
        if l:find("^|") then
          last = k
        end
      end
      api.nvim_buf_set_lines(0, first - 1, last + 1, false, {})
      changes.after_edit(api.nvim_get_current_buf())
      H.save()
      return path
    end

    it("has its word and number deleted too: accepting leaves nothing of it (" .. ext .. ")", function()
      local path = deleted()
      H.open(path)
      changes.resolve("accept", true)
      local text = table.concat(H.lines(), "\n")
      assert.is_nil(text:find("Table", 1, true), text)
      assert.is_nil(text:find("caption", 1, true), text)
    end)

    it("rejected, is there again with its number (" .. ext .. ")", function()
      local path = deleted()
      H.open(path)
      changes.resolve("reject", true)
      local text = table.concat(H.lines(), "\n")
      assert.truthy(text:find("A caption here", 1, true), text)
      assert.truthy(text:find("Table", 1, true), text)
      H.save()
      assert.truthy(zip.read(path, H.MAIN[ext]):find("A caption here", 1, true))
    end)
  end
end)

describe("a comment whose text was all deleted, at the document's end", function()
  local neo = require("neoffice")
  for _, ext in ipairs({ "docx", "odt" }) do
    it("is kept, on the paragraph before, which didn't change (" .. ext .. ")", function()
      local path = H.copy("empty." .. ext)
      H.open(path)
      api.nvim_buf_set_lines(0, 0, -1, false, { "Para one.", "", "| a | b |", "|---|---|", "| 1 | 2 |", "", "Last para here." })
      H.save()
      H.reopen()
      local n = #H.lines()
      api.nvim_win_set_cursor(0, { n, 0 })
      vim.cmd("normal! v3l\27")
      require("neoffice.comments").add(true, "On the last")
      H.save()
      H.reopen()
      n = #H.lines()
      api.nvim_buf_set_lines(0, n - 2, n, false, {}) -- (the last paragraph, and the blank line before it)
      local said, notify = {}, vim.notify
      vim.notify = function(msg) said[#said + 1] = msg end
      local ok, err = pcall(H.save)
      vim.notify = notify
      assert(ok, err)
      assert.is_nil(table.concat(said, "\n"):find("Save failed", 1, true), table.concat(said, "\n"))
      H.open(path)
      assert.equals(1, #H.comments())
      assert.equals("On the last", H.comments()[1].text)
      assert.is_nil(table.concat(H.lines(), "\n"):find("Last para", 1, true))
    end)
  end
end)

describe("a table cell whose text ends with spaces, then a line break", function()
  local zip = require("neoffice.zip")
  it("opens, and is written as it was when another paragraph is edited (docx)", function()
    local cell = '<w:r><w:t xml:space="preserve">abc     </w:t></w:r><w:r><w:br/></w:r>'
    local tbl = '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>h</w:t></w:r></w:p></w:tc></w:tr>'
      .. "<w:tr><w:tc><w:p>" .. cell .. "</w:p></w:tc></w:tr></w:tbl>"
    local path = H.copy("empty.docx", function(name, text)
      if name ~= "word/document.xml" then
        return text
      end
      return (text:gsub("<w:body>", "<w:body><w:p><w:r><w:t>Intro.</w:t></w:r></w:p>" .. tbl .. "<w:p><w:r><w:t>End.</w:t></w:r></w:p>", 1))
    end)
    H.open(path)
    assert.truthy(vim.tbl_contains(H.lines(), "| abc |"), table.concat(H.lines(), "\n"))
    H.set(1, "Intro, edited.")
    H.save()
    local main = zip.read(path, "word/document.xml")
    assert.truthy(main:find("Intro, edited.", 1, true))
    assert.truthy(main:find(cell, 1, true), "the cell as it was")
  end)
end)

describe("a box right before a table", function()
  local zip = require("neoffice.zip")
  for _, ext in ipairs({ "docx", "odt" }) do
    it("stays before it, in a paragraph of its own, not one more each save (" .. ext .. ")", function()
      local path = H.copy("empty." .. ext)
      H.open(path)
      api.nvim_buf_set_lines(0, 0, -1, false, { "Before.", "", "| a | b |", "|---|---|", "| 1 | 2 |", "", "After." })
      H.save()
      H.reopen()
      local r
      for k, l in ipairs(H.lines()) do
        if l:find("^| +a ") then
          r = k
        end
      end
      api.nvim_buf_set_lines(0, r - 1, r - 1, false, { "::: {width = 5cm}", "In the box.", ":::", "" })
      H.save()
      local paras = function()
        return select(2, zip.read(path, H.MAIN[ext]):gsub(ext == "docx" and "<w:p[ >]" or "<text:p[ >]", ""))
      end
      local n = paras()
      H.set(1, "Before, edited.") -- (saved again, twice, in the same session)
      H.save()
      H.set(1, "Before, edited again.")
      H.save()
      assert.equals(n, paras(), "a paragraph more each save")
      H.open(path)
      local box, tbl
      for k, l in ipairs(H.lines()) do
        box = box or (l == "In the box." and k)
        tbl = tbl or (l:find("^| +a ") and k)
      end
      assert.truthy(box and tbl and box < tbl, table.concat(H.lines(), "\n"))
    end)
  end
end)

describe("a heading added with changes tracked", function()
  local neo = require("neoffice")
  for _, ext in ipairs({ "docx", "odt" }) do
    it("rejected, leaves nothing of it, nor makes the next paragraph a heading (" .. ext .. ")", function()
      local path = H.copy("empty." .. ext)
      H.open(path)
      api.nvim_buf_set_lines(0, 0, -1, false, { "First.", "", "Last." })
      H.save()
      H.reopen()
      neo.config.author = "Reviewer"
      H.doc().tracking = true
      require("neoffice.changes").snapshot(api.nvim_get_current_buf())
      api.nvim_buf_set_lines(0, 2, 2, false, { "## A new heading", "" })
      require("neoffice.changes").after_edit(api.nvim_get_current_buf())
      H.save()
      H.open(path)
      require("neoffice.changes").resolve("reject", true)
      assert.same({ "First.", "", "Last." }, H.lines())
    end)
  end
end)

describe("a comment the text doesn't show (on deleted text)", function()
  local neo = require("neoffice")
  local zip = require("neoffice.zip")
  it("is listed in the comments panel, and kept as it is (docx)", function()
    local D = 'w:author="A" w:date="2024-01-01T00:00:00Z"'
    local path = H.copy("rich.docx", function(name, text)
      if name ~= "word/document.xml" then
        return text
      end
      -- (its range gone, its reference alone in a deletion)
      local ref = '<w:r><w:rPr><w:rStyle w:val="CommentReference" /></w:rPr><w:commentReference w:id="0" /></w:r>'
      text = text:gsub('<w:commentRangeStart w:id="0" />', "", 1):gsub('<w:commentRangeEnd w:id="0" />', "", 1)
      return (text:gsub(vim.pesc(ref), '<w:del w:id="79" ' .. D .. ">" .. ref .. "</w:del>", 1)) -- (as Word leaves it: nothing else deleted)
    end)
    H.open(path)
    assert.equals(0, #api.nvim_buf_get_extmarks(0, neo.ns.comment, 0, -1, {}))
    vim.cmd("Neoffice comments")
    local panel = table.concat(api.nvim_buf_get_lines(0, 0, -1, false), "\n")
    vim.cmd("wincmd p")
    assert.truthy(panel:find("Not in the text here (1)", 1, true), panel)
    assert.truthy(panel:find("A comment", 1, true), panel)
    H.set(1, "Edited elsewhere.")
    H.save()
    local main = zip.read(path, "word/document.xml")
    assert.truthy(main:find('<w:commentReference w:id="0"', 1, true), "kept")
  end)
end)

describe("a box whose caption's number is a field in its runs, as Word writes it", function()
  local zip = require("neoffice.zip")
  it("shows as a box, with its caption, and so does the text after it (docx)", function()
    local path = H.copy("empty.docx")
    H.open(path)
    api.nvim_buf_set_lines(0, 0, -1, false, { "Before.", "", "::: {width = 6cm}", "In the box.", "", "Figure: A caption in it", ":::", "", "After the box." })
    H.save()
    local x = zip.read(path, "word/document.xml")
    local runs = '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> SEQ Figure \\* ARABIC </w:instrText></w:r>'
      .. '<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>%1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>'
    local n
    x, n = x:gsub('<w:fldSimple w:instr=" SEQ Figure \\%* ARABIC "><w:r><w:rPr><w:noProof/></w:rPr><w:t>(%d+)</w:t></w:r></w:fldSimple>', runs)
    assert.equals(1, n)
    zip.write(path, { ["word/document.xml"] = x })
    H.open(path)
    assert.same({ "Before.", "", "::: {width = 6cm}", "In the box.", "", "Figure: A caption in it", ":::", "", "After the box." }, H.lines())
  end)
end)

describe("a document whose paragraphs after its last table are all deleted", function()
  local zip = require("neoffice.zip")
  it("still ends with a paragraph, as Word and LibreOffice need (docx)", function()
    local path = H.copy("empty.docx")
    H.open(path)
    api.nvim_buf_set_lines(0, 0, -1, false, { "Before.", "", "| a | b |", "|---|---|", "| 1 | 2 |", "", "After." })
    H.save()
    H.reopen()
    local n = #H.lines()
    api.nvim_buf_set_lines(0, n - 2, n, false, {}) -- ("After.", and the blank line before it)
    H.save()
    local body = zip.read(path, "word/document.xml"):match("<w:body>(.*)</w:body>"):gsub("<w:sectPr.*$", "")
    assert.is_nil(body:find("</w:tbl>%s*$"), "ends with a table")
    H.open(path)
    assert.truthy(vim.iter(H.lines()):any(function(l) return l:find("^| +1 ") ~= nil end))
  end)
end)
