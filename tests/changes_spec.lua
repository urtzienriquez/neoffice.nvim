-- Tracked changes: reviewing them (at the cursor, in a selection, all), undoing that, and
-- recording your own edits.
local H = dofile("tests/helpers.lua")
local neo = require("neoffice")
local changes = require("neoffice.changes")
local zip = require("neoffice.zip")
local api = vim.api

local function mine()
  return H.changes(neo.config.author)
end

local function kinds(list) -- "type:shown text" of changes, sorted
  local out = vim.tbl_map(function(ch) return ch.type .. ":" .. (ch.shown or "") end, list or H.changes())
  table.sort(out)
  return out
end

local function select(n, from, to) -- visually select paragraph n from text `from` to text `to` (included)
  local r, line = H.row(n), H.lines()[H.row(n)]
  local s = assert(line:find(from, 1, true))
  local e = assert(line:find(to, s, true)) + #to - 1
  vim.fn.setpos("'<", { 0, r, s, 0 })
  vim.fn.setpos("'>", { 0, r, e, 0 })
  vim.cmd("normal! gvv")
end

local function undo(redo)
  vim.cmd(redo and "silent redo" or "silent undo")
  changes.after_edit(api.nvim_get_current_buf()) -- (what TextChanged does)
end

for _, ext in ipairs({ "docx", "odt" }) do
  describe(ext .. " review", function()
    before_each(function()
      H.open(H.copy("rich." .. ext))
    end)

    it("shows insertions, and deleted text struck through in the text", function()
      assert.equals("Text with inserted words and deleted words inside.", H.paras()[4])
      assert.same({ "delete:deleted words", "insert:inserted words" }, kinds())
    end)

    it("accepting keeps what you see, rejecting goes back", function()
      H.cursor(4, "inserted")
      changes.resolve("accept") -- (the inserted text stays)
      H.cursor(4, "deleted")
      changes.resolve("accept") -- (the deleted text goes)
      assert.equals("Text with inserted words and  inside.", H.paras()[4])
      H.open(H.copy("rich." .. ext))
      H.cursor(4, "inserted")
      changes.resolve("reject") -- (the inserted text goes)
      H.cursor(4, "deleted")
      changes.resolve("reject") -- (the deleted text stays)
      assert.equals("Text with  and deleted words inside.", H.paras()[4])
      assert.equals(0, #H.changes())
      H.reopen()
      assert.equals("Text with  and deleted words inside.", H.paras()[4])
      assert.equals(0, #H.changes())
    end)

    it("acts on every change a visual selection touches", function()
      select(4, "with", "inside")
      changes.resolve("reject", false, true) -- (the last selection, as :'<,'>Neoffice reject)
      assert.equals("Text with  and deleted words inside.", H.paras()[4])
      assert.equals(0, #H.changes())
    end)

    it("acts on the selected part only, and the rest stays a change", function()
      select(4, "words inside", "words") -- (the "words" of "deleted words")
      changes.resolve("accept", false, true)
      select(4, "inserted", "inserted")
      changes.resolve("reject", false, true)
      assert.equals("Text with  words and deleted  inside.", H.paras()[4])
      assert.same({ "delete:deleted ", "insert: words" }, kinds())
      H.reopen()
      assert.equals("Text with  words and deleted  inside.", H.paras()[4])
      assert.same({ "delete:deleted ", "insert: words" }, kinds())
    end)

    it("is undone and redone with u and <C-r>, changes and all", function()
      H.cursor(4, "deleted")
      changes.resolve("accept")
      H.cursor(4, "inserted")
      changes.resolve("accept") -- (a change of marks only: still a step to undo)
      assert.equals(0, #H.changes())
      undo()
      assert.same({ "insert:inserted words" }, kinds())
      undo()
      assert.equals("Text with inserted words and deleted words inside.", H.paras()[4])
      assert.same({ "delete:deleted words", "insert:inserted words" }, kinds())
      undo(true)
      assert.equals("Text with inserted words and  inside.", H.paras()[4])
      assert.same({ "insert:inserted words" }, kinds())
      H.reopen()
      assert.same({ "insert:inserted words" }, kinds())
    end)

    it("accepts and rejects them all", function()
      vim.cmd("Neoffice accept all")
      H.reopen()
      assert.equals("Text with inserted words and  inside.", H.paras()[4])
      assert.equals(0, #H.changes())
      H.open(H.copy("rich." .. ext))
      vim.cmd("Neoffice reject all")
      H.reopen()
      assert.equals("Text with  and deleted words inside.", H.paras()[4])
    end)

    it("leaves the other paragraphs alone when reviewing", function()
      local before = H.paragraph_xml(H.doc().path)
      vim.cmd("Neoffice accept all")
      H.save()
      assert.equals(1, H.differing(before, H.paragraph_xml(H.doc().path)))
    end)

    it("keeps text typed right before a deletion out of it", function()
      local r = H.row(4)
      local col = H.lines()[r]:find("deleted", 1, true) - 1
      api.nvim_buf_set_text(0, r - 1, col, r - 1, col, { "new " }) -- (typed right before it)
      changes.after_edit(api.nvim_get_current_buf())
      assert.same({ "delete:deleted words", "insert:inserted words" }, kinds())
      H.reopen()
      assert.equals("Text with inserted words and new deleted words inside.", H.paras()[4])
      vim.cmd("Neoffice accept all")
      assert.equals("Text with inserted words and new  inside.", H.paras()[4])
    end)

    it("jumps between changes and lists them", function()
      api.nvim_win_set_cursor(0, { 1, 0 })
      changes.jump(1)
      assert.equals(H.row(4), api.nvim_win_get_cursor(0)[1])
      changes.list()
      assert.equals(2, #vim.fn.getloclist(0))
      vim.cmd("lclose")
    end)

    it("says so when there is no change at the cursor, changing nothing", function()
      api.nvim_win_set_cursor(0, { 1, 0 })
      changes.resolve("accept")
      assert.is_false(vim.bo.modified)
    end)
  end)

  describe(ext .. " tracking your edits", function()
    before_each(function()
      H.open(H.copy("rich." .. ext))
      vim.cmd("Neoffice track on")
    end)

    it("keeps deleted text, struck through, and inserted text highlighted; and saves them", function()
      H.replace(12, "Last", "Final")
      assert.equals("LastFinal paragraph.", H.paras()[12])
      assert.same({ "delete:Last", "insert:Final" }, kinds(mine()))
      H.reopen()
      assert.equals("LastFinal paragraph.", H.paras()[12])
      assert.same({ "delete:Last", "insert:Final" }, kinds(mine()))
      changes.resolve("reject", true)
      assert.equals("Last paragraph.", H.paras()[12])
    end)

    it("grows one deletion while deleting letter by letter, and forgets text you typed and deleted", function()
      local r = H.row(12)
      local function edit(c1, c2, text)
        api.nvim_buf_set_text(0, r - 1, c1, r - 1, c2, { text })
        changes.after_edit(api.nvim_get_current_buf())
      end
      edit(3, 4, "") -- "t"
      edit(2, 3, "") -- "s"
      edit(0, 0, "xy")
      edit(1, 2, "") -- "y", typed just now
      assert.same({ "delete:st", "insert:x" }, kinds(mine()))
      assert.equals("xLast paragraph.", H.paras()[12])
    end)

    it("extends your insertion while you type", function()
      local r = H.row(12)
      for _, text in ipairs({ " and", " more", " words" }) do
        local col = #H.lines()[r] - 1 -- (before the final ".")
        api.nvim_buf_set_text(0, r - 1, col, r - 1, col, { text })
        changes.after_edit(api.nvim_get_current_buf())
      end
      assert.equals("Last paragraph and more words.", H.paras()[12])
      assert.equals(1, #mine())
    end)

    it("doesn't let deleted text be deleted (it stays deleted by whoever deleted it)", function()
      H.replace(4, "deleted words", "")
      assert.equals("Text with inserted words and deleted words inside.", H.paras()[4])
      assert.same({ "delete:deleted words", "insert:inserted words" }, kinds())
      assert.equals(0, #mine())
    end)

    it("is undone with u, the change and all", function()
      H.replace(12, "Last ", "")
      assert.equals(1, #mine())
      undo()
      assert.equals("Last paragraph.", H.paras()[12])
      assert.equals(0, #mine())
    end)

    it("makes text typed inside someone else's insertion yours, splitting theirs", function()
      local function type_at(at, text)
        H.cursor(4, at)
        local r, c = unpack(api.nvim_win_get_cursor(0))
        api.nvim_buf_set_text(0, r - 1, c, r - 1, c, { text })
        changes.after_edit(api.nvim_get_current_buf())
      end
      local function owners()
        local out = {}
        for _, m in ipairs(api.nvim_buf_get_extmarks(0, require("neoffice").ns.change, 0, -1, { details = true })) do
          local ch = H.doc().change[m[1]]
          if ch.type == "insert" then
            out[#out + 1] = ch.author .. ":" .. H.lines()[m[2] + 1]:sub(m[3] + 1, m[4].end_col)
          end
        end
        return out
      end
      local me = require("neoffice").config.author
      type_at("words and", "my ")
      type_at("inserted", "Oh ") -- (at the very start of Bob's)
      assert.same({ me .. ":Oh ", "Bob:inserted ", me .. ":my ", "Bob:words" }, owners())
      H.reopen()
      assert.same({ me .. ":Oh ", "Bob:inserted ", me .. ":my ", "Bob:words" }, owners())
    end)

    -- the owner of each change mark: "author type:text"
    local function owners(rows)
      local out = {}
      for _, m in ipairs(api.nvim_buf_get_extmarks(0, neo.ns.change, 0, -1, { details = true })) do
        local ch = H.doc().change[m[1]]
        if not rows or (m[2] + 1 >= rows[1] and m[2] + 1 <= rows[2]) then
          local text = table.concat(api.nvim_buf_get_text(0, m[2], m[3], m[4].end_row, m[4].end_col, {}), "\n")
          out[#out + 1] = ch.author .. " " .. ch.type .. ":" .. text
        end
      end
      return out
    end
    local function type_at(row, col, text) -- (as typing does: the text, then TextChanged)
      api.nvim_buf_set_text(0, row - 1, col, row - 1, col, vim.split(text, "\n"))
      changes.after_edit(api.nvim_get_current_buf())
    end

    it("keeps text typed at the start of a deletion (also on a new line, as O does) out of it", function()
      local me = neo.config.author
      local r = H.row(12)
      api.nvim_buf_set_text(0, r - 1, 0, r - 1, #"Last paragraph.", {}) -- (deleted: stays, struck through)
      changes.after_edit(api.nvim_get_current_buf())
      assert.same({ me .. " delete:Last paragraph." }, owners({ r, r }))
      type_at(r, 0, "\n") -- O then <Esc>: an empty line, not deleted
      assert.same({ me .. " delete:Last paragraph." }, owners({ r, r + 1 }))
      type_at(r + 1, 0, "New\n") -- (at the start of the deletion's line: a line of mine above it)
      assert.same({ me .. " insert:New\n", me .. " delete:Last paragraph." }, owners({ r, r + 2 }))
    end)

    it("splits a deletion around text typed inside it", function()
      local r = H.row(4)
      local c = H.lines()[r]:find("deleted words", 1, true) + #"deleted" - 1
      type_at(r, c, "XY")
      assert.same({ "Bob delete:deleted", neo.config.author .. " insert:XY", "Bob delete: words" }, vim.tbl_filter(function(o)
        return not o:find("inserted words", 1, true)
      end, owners({ r, r })))
      H.reopen()
      assert.equals("Text with inserted words and deletedXY words inside.", H.paras()[4])
    end)

    it("records what you typed where you typed it, even with the same letters around it", function()
      local r = H.row(4)
      type_at(r, H.lines()[r]:find("words and", 1, true) - 1, "ed in ") -- ("inserted ed in words": also "inserted ed in" + " words"...)
      assert.same({ "Bob insert:inserted ", neo.config.author .. " insert:ed in ", "Bob insert:words" }, vim.tbl_filter(function(o)
        return not o:find("delete")
      end, owners({ r, r })))
      H.reopen()
      assert.same({ "Bob insert:inserted ", neo.config.author .. " insert:ed in ", "Bob insert:words" }, vim.tbl_filter(function(o)
        return not o:find("delete")
      end, owners({ r, r })))
    end)

    it("records a replacement (cw, a paste over a selection) as what was there and what is there now", function()
      local r = H.row(12)
      api.nvim_buf_set_text(0, r - 1, 5, r - 1, 14, { "paragraphs" }) -- ("paragraph" -> "paragraphs", in one edit)
      changes.after_edit(api.nvim_get_current_buf())
      local me = neo.config.author
      assert.same({ me .. " delete:paragraph", me .. " insert:paragraphs" }, owners({ r, r }))
    end)

    it("records separate edits made at once (:s///g) each where it is", function()
      local r = H.row(12)
      api.nvim_buf_set_text(0, r - 1, 5, r - 1, 14, { "line" })
      api.nvim_buf_set_text(0, r - 1, 0, r - 1, 4, { "First" }) -- ("First line.", before one TextChanged)
      changes.after_edit(api.nvim_get_current_buf())
      local me = neo.config.author
      assert.same({ me .. " delete:Last", me .. " insert:First", me .. " delete:paragraph", me .. " insert:line" }, owners({ r, r }))
      assert.equals("LastFirst paragraphline.", H.paras()[12])
    end)

    it("edits an image's alt text and width without a tracked change (nor a second image)", function()
      local r = H.row(5)
      local c = H.lines()[r]:find("![]", 1, true) + 1
      type_at(r, c, "A dot")
      local w = H.lines()[r]:find("1.4cm", 1, true) - 1
      api.nvim_buf_set_text(0, r - 1, w, r - 1, w + #"1.4cm", { "3cm" })
      changes.after_edit(api.nvim_get_current_buf())
      assert.same({}, owners({ r, r }))
      H.reopen()
      assert.truthy(H.paras()[5]:match("^An image !%[A dot%]%([^)]*%){width = 3cm} here%.$"), H.paras()[5])
    end)

    it("tells, while you type, that what you just typed is yours", function()
      local r = H.row(4)
      local c = H.lines()[r]:find("words and", 1, true) - 1 -- (inside Bob's insertion)
      type_at(r, c, "my ")
      api.nvim_win_set_cursor(0, { r, c + 3 }) -- (in insert mode the cursor is after what you typed)
      local get_mode = api.nvim_get_mode
      api.nvim_get_mode = function() return { mode = "i", blocking = false } end
      changes.show_info(api.nvim_get_current_buf())
      api.nvim_get_mode = get_mode
      local info = api.nvim_buf_get_extmarks(0, -1, { r - 1, 0 }, { r - 1, -1 }, { details = true, type = "virt_text" })
      local texts = vim.tbl_map(function(m) return m[4].virt_text[1][1] end, info)
      assert.truthy(vim.iter(texts):find(function(t) return t:find(neo.config.author .. " ·", 1, true) end), vim.inspect(texts))
    end)

    it("with tracking off, makes what you type plain text, splitting the changes around it", function()
      vim.cmd("Neoffice track off")
      local r = H.row(4)
      local line = H.lines()[r]
      type_at(r, line:find("words and", 1, true) - 1, "my ") -- (inside Bob's insertion)
      line = H.lines()[r]
      type_at(r, line:find("deleted words", 1, true) + #"deleted" - 1, "XY") -- (inside Bob's deletion)
      local expected = { "Bob insert:inserted ", "Bob insert:words", "Bob delete:deleted", "Bob delete: words" }
      assert.same(expected, owners({ r, r }))
      H.reopen()
      assert.equals("Text with inserted my words and deletedXY words inside.", H.paras()[4])
      assert.same(expected, owners({ r, r }))
    end)

    it("doesn't record deleting blank lines", function()
      local r = H.row(12)
      api.nvim_buf_set_lines(0, r - 1, r - 1, false, { "", "" }) -- (two empty paragraphs)
      changes.after_edit(api.nvim_get_current_buf())
      api.nvim_buf_set_lines(0, r - 1, r + 1, false, {}) -- (and gone again)
      changes.after_edit(api.nvim_get_current_buf())
      assert.equals(0, #mine())
      assert.equals("Last paragraph.", H.paras()[12])
      assert.falsy(table.concat(H.lines()):find("¶", 1, true))
    end)

    it("keeps a deleted paragraph where it was, struck through; accepting removes it, rejecting keeps it", function()
      api.nvim_win_set_cursor(0, { H.row(5), 0 })
      vim.cmd("normal! dd") -- ("An image ... here.")
      changes.after_edit(api.nvim_get_current_buf())
      assert.equals(12, #H.paras())
      assert.equals(2, #mine()) -- (the text before and after the image, which stays: images aren't tracked)
      H.reopen()
      assert.equals(12, #H.paras())
      vim.cmd("Neoffice track off")
      -- the whole line (V), with its paragraph break: the image in it stays (images aren't tracked)
      vim.fn.setpos("'<", { 0, H.row(5), 1, 0 })
      vim.fn.setpos("'>", { 0, H.row(5), 1, 0 })
      vim.cmd("normal! gvVV")
      changes.resolve("accept", false, true)
      local p = H.paras()
      assert.equals("Text with inserted words and deleted words inside.", p[4])
      assert.truthy(p[5]:match("^!%[%]%(.-%){width = [^}]*}$"))
      assert.equals("- first item", p[6])
      H.reopen()
      assert.truthy(H.paras()[5]:match("^!%[%]%(.-%){width = [^}]*}$"))
    end)

    it("records a new paragraph, and rejecting it removes it", function()
      api.nvim_buf_set_lines(0, H.row(12), H.row(12), false, { "", "Brand new." })
      changes.after_edit(api.nvim_get_current_buf())
      H.reopen()
      assert.equals("Brand new.", H.paras()[13])
      vim.cmd("Neoffice track off")
      for _, ch in ipairs({ "Bob" }) do -- (only your own: Bob's stay)
        assert.truthy(vim.iter(H.changes()):any(function(c) return c.author == ch end))
      end
      local marks = api.nvim_buf_get_extmarks(0, neo.ns.change, { H.row(12) - 1, 0 }, -1, {})
      for _ = 1, #marks do -- (one at a time, while it is there: its text and its break may go at once)
        if #H.paras() < 13 then
          break
        end
        api.nvim_win_set_cursor(0, { H.row(13), 0 })
        changes.resolve("reject")
      end
      assert.equals(12, #H.paras())
      assert.equals("Last paragraph.", H.paras()[12])
    end)

    it("doesn't record changes in formatting marks, or your review", function()
      H.replace(12, "Last", "**Last**")
      assert.equals(0, #mine())
      vim.cmd("Neoffice accept all")
      assert.equals(0, #H.changes())
    end)

    it("stops when turned off", function()
      vim.cmd("Neoffice track off")
      H.replace(12, "Last", "Final")
      assert.equals(0, #mine())
      assert.equals("Final paragraph.", H.paras()[12])
    end)

    if H.has_libreoffice then
      it("writes changes LibreOffice sees as tracked changes", function()
        H.replace(12, "Last", "Final")
        H.save()
        local dir, path = vim.fn.tempname(), H.doc().path
        local to = ext == "docx" and "odt" or "docx"
        vim.system({ "soffice", "-env:UserInstallation=file://" .. dir .. "/p", "--headless", "--convert-to", to, "--outdir", dir, path }):wait()
        local main = zip.read(dir .. "/rich." .. to, to == "odt" and "content.xml" or "word/document.xml")
        local author = vim.pesc(neo.config.author)
        if to == "odt" then
          assert.truthy(main:find("<text:insertion><office:change%-info><dc:creator>" .. author))
          assert.truthy(main:find("<text:deletion><office:change%-info><dc:creator>" .. author))
        else
          assert.truthy(main:find('<w:ins [^>]*w:author="' .. author))
          assert.truthy(main:find('<w:del [^>]*w:author="' .. author))
        end
      end)
    end
  end)
end

describe("tracking", function()
  it("starts on when the document has it on, and the document's setting is never changed", function()
    local path = H.copy("rich.docx")
    zip.write(path, { ["word/settings.xml"] = '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:trackRevisions/></w:settings>' })
    H.open(path)
    assert.is_true(H.doc().tracking)
    vim.cmd("Neoffice track off")
    H.replace(12, "Last", "Final")
    H.save()
    assert.truthy(zip.read(path, "word/settings.xml"):find("trackRevisions"))
  end)

  it("follows the track_changes option", function()
    neo.config.track_changes = true
    H.open(H.copy("rich.odt"))
    neo.config.track_changes = nil
    assert.is_true(H.doc().tracking)
  end)
end)

describe("docx formatting changes", function()
  local change = '<w:rPrChange w:id="50" w:author="Bob" w:date="2024-01-01T00:00:00Z"><w:rPr/></w:rPrChange>'
  for _, action in ipairs({ "accept", "reject" }) do
    it("can be " .. action .. "ed", function()
      H.open(H.copy("rich.docx", function(_, s) return (s:gsub("<w:b /><w:bCs /></w:rPr>", "<w:b /><w:bCs />" .. change .. "</w:rPr>", 1)) end))
      assert.equals(3, #H.changes())
      H.cursor(2, "bold")
      changes.resolve(action)
      H.reopen()
      assert.equals(action == "accept", H.paras()[2]:find("**bold**", 1, true) ~= nil)
      assert.falsy(zip.read(H.doc().path, "word/document.xml"):find("rPrChange"))
    end)
  end
end)

describe("several reviewers", function()
  local function setup()
    H.open(H.copy("rich.docx"))
    vim.cmd("Neoffice track on")
    H.replace(12, "Last", "Final") -- (yours, next to Bob's)
    vim.cmd("Neoffice track off")
  end

  it("get a colour each", function()
    setup()
    local groups = {}
    for _, m in ipairs(api.nvim_buf_get_extmarks(0, neo.ns.change, 0, -1, { details = true })) do
      groups[neo.docs[api.nvim_get_current_buf()].change[m[1]].author .. ":" .. m[4].hl_group:gsub("%a+(%d)$", "%1")] = true
    end
    local bob = vim.tbl_filter(function(k) return k:match("^Bob:") end, vim.tbl_keys(groups))
    local me = vim.tbl_filter(function(k) return k:match("^" .. vim.pesc(neo.config.author) .. ":") end, vim.tbl_keys(groups))
    assert.equals(1, #bob)
    assert.equals(1, #me)
    assert.are_not.equals(bob[1]:match("%d$"), me[1]:match("%d$"))
  end)

  it("are named at the cursor", function()
    setup()
    H.cursor(4, "inserted")
    changes.show_info(api.nvim_get_current_buf())
    local info = api.nvim_buf_get_extmarks(0, api.nvim_get_namespaces().neoffice_change_info, 0, -1, { details = true })
    assert.equals(1, #info)
    assert.truthy(info[1][4].virt_text[1][1]:find("Bob · 2024%-01%-02 · inserted"))
  end)

  it("can be accepted, rejected and listed one at a time", function()
    setup()
    assert.same({ "Bob", neo.config.author }, changes.authors())
    vim.cmd("Neoffice accept all Bob")
    assert.same({ neo.config.author }, changes.authors())
    assert.equals("Text with inserted words and  inside.", H.paras()[4])
    vim.cmd("Neoffice changes " .. neo.config.author)
    assert.equals(2, #vim.fn.getloclist(0))
    vim.cmd("lclose")
    vim.cmd("Neoffice reject all " .. neo.config.author)
    assert.equals(0, #H.changes())
    assert.equals("Last paragraph.", H.paras()[12])
  end)

  it("are completed by name", function()
    setup()
    assert.same({ "Bob" }, vim.fn.getcompletion("Neoffice accept all B", "cmdline"))
    assert.is_true(vim.tbl_contains(vim.fn.getcompletion("Neoffice changes ", "cmdline"), "Bob"))
  end)
end)

-- John inserted an empty paragraph (Word marks its paragraph break as inserted), then a line of
-- someone else's follows: typing in John's empty line, with tracking on or off.
describe("typing in someone else's inserted empty paragraph", function()
  local D = 'w:date="2024-05-01T10:00:00Z"'
  local BODY = '<w:p><w:r><w:t>Some text</w:t></w:r></w:p>'
    .. '<w:p><w:pPr><w:rPr><w:ins w:id="1" w:author="John" ' .. D .. '/></w:rPr></w:pPr></w:p>'
    .. '<w:p><w:pPr><w:rPr><w:ins w:id="2" w:author="Ann" ' .. D .. '/></w:rPr></w:pPr><w:ins w:id="3" w:author="Ann" ' .. D .. '><w:r><w:t>Next line by Ann.</w:t></w:r></w:ins></w:p>'
    .. '<w:p><w:r><w:t>Last.</w:t></w:r></w:p>'
  local function open()
    H.open(H.copy("empty.docx", function(entry, text)
      return entry == "word/document.xml" and text:gsub("<w:body>(.-)(<w:sectPr)", "<w:body>" .. BODY .. "%2") or text
    end))
    assert.same({ "Some text", "", "", "Next line by Ann.", "", "Last." }, H.lines())
  end
  local function owner_of(text) -- the change covering `text` (nil: plain text)
    local r = vim.fn.index(H.lines(), text) + 1
    local found
    for _, m in ipairs(api.nvim_buf_get_extmarks(0, neo.ns.change, { r - 1, 0 }, { r - 1, -1 }, { details = true, overlap = true })) do
      if m[2] <= r - 1 and m[3] == 0 and (m[4].end_row > r - 1 or m[4].end_col >= #text) then
        found = H.doc().change[m[1]].author .. " " .. H.doc().change[m[1]].type
      end
    end
    return found
  end
  local function type_in_line_3(text)
    api.nvim_buf_set_text(0, 2, 0, 2, 0, { text })
    changes.after_edit(api.nvim_get_current_buf())
  end

  it("with tracking on, what you type is your insertion; on its own paragraph", function()
    open()
    vim.cmd("Neoffice track on")
    type_in_line_3("Mine.")
    assert.equals(neo.config.author .. " insert", owner_of("Mine."))
    H.reopen()
    assert.same({ "Some text", "Mine.", "Next line by Ann.", "Last." }, H.paras())
    assert.equals(neo.config.author .. " insert", owner_of("Mine."))
    assert.equals("Ann insert", owner_of("Next line by Ann."))
  end)

  it("with tracking off, what you type is plain text; on its own paragraph", function()
    open()
    vim.cmd("Neoffice track off")
    type_in_line_3("Plain.")
    assert.is_nil(owner_of("Plain."))
    H.reopen()
    assert.same({ "Some text", "Plain.", "Next line by Ann.", "Last." }, H.paras())
    assert.is_nil(owner_of("Plain."))
    assert.equals("Ann insert", owner_of("Next line by Ann."))
    assert.truthy(zip.read(H.doc().path, "word/document.xml"):find('w:author="John"', 1, true)) -- (his paragraph break)
  end)
end)
