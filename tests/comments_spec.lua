-- Comments: adding, replying, resolving and deleting, without touching the other comments.
local H = dofile("tests/helpers.lua")
local neo = require("neoffice")
local comments = require("neoffice.comments")
local zip = require("neoffice.zip")
local xml = require("neoffice.xml")
local api = vim.api

local function add(n, text, comment) -- comment `text` in paragraph n (as :'<,'>Neoffice comment does)
  local r = H.row(n)
  local s = assert(H.lines()[r]:find(text, 1, true))
  vim.fn.setpos("'<", { 0, r, s, 0 })
  vim.fn.setpos("'>", { 0, r, s + #text - 1, 0 })
  vim.fn.histadd("cmd", "") -- (visualmode() needs a visual selection to have been made)
  vim.cmd("normal! gvv")
  H.input = comment
  comments.add(true)
end

for _, name in ipairs({ "rich.docx", "office.docx", "rich.odt" }) do
  describe(name, function()
    before_each(function()
      H.open(H.copy(name))
    end)

    it("shows the comment and its range", function()
      local list = H.comments()
      assert.equals(1, #list)
      assert.equals("A comment", list[1].text)
      assert.equals("Alice", list[1].author)
      local m = api.nvim_buf_get_extmarks(0, neo.ns.comment, 0, -1, { details = true })[1]
      local line = H.lines()[m[2] + 1]
      assert.equals("Commented text", line:sub(m[3] + 1, m[4].end_col))
    end)

    it("adds a comment on a selection", function()
      add(12, "paragraph", "Why last?")
      assert.is_true(vim.bo.modified)
      H.reopen()
      local list = H.comments()
      assert.equals(2, #list)
      assert.equals("Why last?", list[2].text)
      assert.equals(neo.config.author, list[2].author)
      local m = vim.iter(api.nvim_buf_get_extmarks(0, neo.ns.comment, 0, -1, { details = true })):find(function(m)
        return H.doc().comment[m[1]] == list[2]
      end)
      assert.equals("paragraph", H.lines()[m[2] + 1]:sub(m[3] + 1, m[4].end_col))
    end)

    it("adds a comment at the cursor (no selection)", function()
      H.cursor(12, "paragraph")
      H.input = "Here."
      comments.add()
      H.reopen()
      assert.equals("Here.", H.comments()[2].text)
    end)

    it("replies, resolves and reopens from the document", function()
      H.cursor(3, "Commented")
      H.input = "A reply from Neovim"
      comments.reply()
      comments.resolve()
      H.reopen()
      local c = H.comments()[1]
      assert.is_true(c.resolved)
      assert.equals("A reply from Neovim", c.replies[#c.replies].text)
      H.cursor(3, "Commented")
      comments.resolve()
      H.reopen()
      assert.is_falsy(H.comments()[1].resolved)
    end)

    it("deletes a comment, leaving no trace of it", function()
      H.cursor(3, "Commented")
      comments.delete()
      H.reopen()
      assert.equals(0, #H.comments())
      local main = zip.read(H.doc().path, H.MAIN[name:match("%w+$")])
      assert.falsy(main:find("commentRange") or main:find("commentReference") or main:find("office:annotation"))
      assert.equals("Commented text and more.", H.paras()[3])
    end)

    it("deletes a reply only", function()
      H.cursor(3, "Commented")
      H.input = "First reply"
      comments.reply()
      H.input = "Second reply"
      comments.reply()
      H.reopen()
      vim.cmd("Neoffice comments") -- in the panel, on the last reply's header line
      local lines, row = api.nvim_buf_get_lines(0, 0, -1, false), nil
      for k = #lines, 1, -1 do
        if lines[k]:find("↳", 1, true) then
          row = k
          break
        end
      end
      api.nvim_win_set_cursor(0, { row, 0 })
      comments.delete()
      vim.cmd("wincmd p")
      H.reopen()
      local replies = vim.tbl_map(function(r) return r.text end, H.comments()[1].replies)
      assert.is_true(vim.tbl_contains(replies, "First reply"))
      assert.is_false(vim.tbl_contains(replies, "Second reply"))
    end)

    it("keeps the comment when the text it is on is deleted", function()
      H.replace(3, "Commented text and more.", "Other words.")
      H.reopen()
      assert.equals(1, #H.comments())
      assert.equals("Other words.", H.paras()[3])
    end)

    it("leaves the existing comments' XML alone when adding another", function()
      local ext = name:match("%w+$")
      local function existing()
        local path = H.doc().path
        if ext == "odt" then
          return zip.read(path, "content.xml"):match("<office:annotation .-</office:annotation>")
        end
        return xml.serialize(xml.collect(xml.parse(zip.read(path, "word/comments.xml")), { ["w:comment"] = true })[1])
      end
      local before = existing()
      add(12, "paragraph", "Another")
      H.reopen()
      assert.equals(before, existing())
    end)
  end)
end

describe("docx comment parts", function()
  it("are created (and registered) in a document that had no comments", function()
    local path = H.copy("empty.docx")
    H.open(path)
    api.nvim_buf_set_lines(0, 0, -1, false, { "Some text to comment." })
    H.save()
    add(1, "text", "First comment ever")
    H.reopen()
    assert.equals("First comment ever", H.comments()[1].text)
    assert.truthy(zip.read(path, "[Content_Types].xml"):find("/word/comments.xml", 1, true))
    assert.truthy(zip.read(path, "word/_rels/document.xml.rels"):find('Target="comments.xml"', 1, true))
  end)

  it("lose a deleted comment everywhere, including Word's commentsIds and commentsExtensible", function()
    local path = H.copy("office.docx")
    local para = zip.read(path, "word/comments.xml"):match('w14:paraId="(%x+)"')
    zip.write(path, {
      ["word/commentsIds.xml"] = ('<w16cid:commentsIds xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid"><w16cid:commentId w16cid:paraId="%s" w16cid:durableId="1A2B3C4D"/></w16cid:commentsIds>'):format(para),
      ["word/commentsExtensible.xml"] = '<w16cex:commentsExtensible xmlns:w16cex="http://schemas.microsoft.com/office/word/2018/wordml/cex"><w16cex:commentExtensible w16cex:durableId="1A2B3C4D" w16cex:dateUtc="2024-01-01T10:00:00Z"/></w16cex:commentsExtensible>',
    })
    H.open(path)
    H.cursor(3, "Commented")
    comments.delete()
    H.save()
    assert.falsy(zip.read(path, "word/commentsIds.xml"):find(para, 1, true))
    assert.falsy(zip.read(path, "word/commentsExtensible.xml"):find("1A2B3C4D", 1, true))
    assert.falsy(zip.read(path, "word/comments.xml"):find(para, 1, true))
  end)

  it("thread a reply to a comment that had no paragraph id", function()
    local path = H.copy("rich.docx") -- (pandoc writes comments without w14:paraId)
    H.open(path)
    H.cursor(3, "Commented")
    H.input = "Threaded?"
    comments.reply()
    H.save()
    local extended = zip.read(path, "word/commentsExtended.xml")
    assert.truthy(extended:find("w15:paraIdParent", 1, true))
    H.open(path)
    assert.equals("Threaded?", H.comments()[1].replies[1].text)
  end)

  if H.has_libreoffice then
    it("LibreOffice reads new comments and replies", function()
      local path = H.copy("office.docx")
      H.open(path)
      H.cursor(3, "Commented")
      H.input = "Reply for LibreOffice"
      comments.reply()
      add(12, "paragraph", "Comment for LibreOffice")
      H.save()
      local dir = vim.fn.tempname()
      vim.system({ "soffice", "-env:UserInstallation=file://" .. dir .. "/p", "--headless", "--convert-to", "odt", "--outdir", dir, path }):wait()
      local content = zip.read(dir .. "/office.odt", "content.xml")
      assert.truthy(content:find("Reply for LibreOffice", 1, true))
      assert.truthy(content:find("Comment for LibreOffice", 1, true))
      assert.truthy(content:find("loext:parent%-name"))
    end)
  end
end)

describe("the comments panel", function()
  it("moves between comments with ]] and [[", function()
    H.open(H.copy("rich.docx"))
    H.cursor(12, "paragraph")
    H.input = "Second"
    comments.add()
    vim.cmd("Neoffice comments")
    local function line() return api.nvim_get_current_line() end
    api.nvim_win_set_cursor(0, { 1, 0 })
    assert.truthy(line():find("Alice"))
    vim.cmd("normal ]]")
    assert.truthy(line():find(neo.config.author, 1, true))
    vim.cmd("normal ]]") -- (no further: stays)
    assert.truthy(line():find(neo.config.author, 1, true))
    vim.cmd("normal [[")
    assert.truthy(line():find("Alice"))
  end)
end)
