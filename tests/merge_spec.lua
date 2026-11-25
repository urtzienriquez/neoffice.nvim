-- :Neoffice merge: a revised document's edits made in the LaTeX it was exported from, and only
-- those; what can't be merged safely is left as it is and listed.
local H = dofile("tests/helpers.lua")
local neo = require("neoffice")
local merge = require("neoffice.merge")
merge.memory = vim.fn.tempname() .. ".json" -- (the sources merged into: not the user's)
local api = vim.api

local BIB = [[
@article{smith2020,
  author = {Smith, A. and Lee, B. and Kim, C.},
  year = {2020},
}
@article{jones2019,
  author = {Jones, D.},
  year = {2019},
}
@comment{exported by Zotero, @article{fake2000, author = {Fake, A.}, year = {2000}}}
@article{wild2025, title={Lizards}, author={Wild, Kristoffer H. and Huey, Raymond B. and Pianka, Eric R.}, year={2025} }
@article{brown2018,
  author = {Brown, E. and Green, F.},
  year = {2018},
}
@article{martinez2021,
  author = {Mart{\'\i}nez-Fre{\'\i}ria, F. and Santos, X. and Brito, J.},
  date = {2021-03},
}
]]

local SOURCE = [[
\section{Introduction}

We studied \emph{Vipera aspis} in the north of Iberia
\parencite{smith2020,jones2019}. % (wrapped as I like it)
Warming raised its water loss by \Sexpr{rise()} percent.

A second paragraph, with \textbf{bold words} in it.
]]

-- The source, its .bib, and the document it was exported to (as the export writes it), opened
-- with tracking on (off with `untracked`, and a copy as sent kept beside it, as the export
-- does): returns the source's path.
local function setup(document, source, untracked)
  local dir = vim.fn.tempname()
  vim.fn.mkdir(dir, "p")
  for name, text in pairs({ ["main.tex"] = source or SOURCE, ["refs.bib"] = BIB }) do
    local f = assert(io.open(dir .. "/" .. name, "w"))
    f:write(text)
    f:close()
  end
  H.open(dir .. "/main.docx")
  api.nvim_buf_set_lines(0, 0, -1, false, document or {
    "# Introduction",
    "",
    "We studied *Vipera aspis* in the north of Iberia (Smith et al., 2020; Jones, 2019). Warming raised its water loss by 12 percent.",
    "",
    "A second paragraph, with **bold words** in it.",
  })
  H.save()
  if untracked then
    assert(vim.uv.fs_copyfile(dir .. "/main.docx", dir .. "/main.sent.docx"))
  end
  H.open(dir .. "/main.docx")
  vim.cmd(untracked and "Neoffice track off" or "Neoffice track on")
  return dir .. "/main.tex"
end

-- The proposal, the notes, and the edits (each also listed in the quickfix list).
local function run(path)
  local out, issues = merge.compute(api.nvim_get_current_buf(), { path })
  local notes, edits = {}, {}
  for _, i in ipairs(issues) do
    table.insert(i.user_data == "edit" and edits or notes, i.text)
  end
  return out[path], notes, edits
end

local function has(list, text)
  return vim.iter(list):any(function(t) return t:find(text, 1, true) ~= nil end)
end

describe(":Neoffice merge", function()
  it("changes nothing when nothing changed", function()
    local path = setup()
    local out, issues = run(path)
    assert.equals(SOURCE, out)
    assert.same({}, issues)
  end)

  it("changes only the words edited, keeping the wrapping, comments and commands", function()
    local path = setup()
    H.replace(2, "north", "south")
    H.replace(3, "second", "last")
    local out, _, edits = run(path)
    assert.equals(SOURCE:gsub("north", "south"):gsub("second", "last"), out)
    assert.same({ "change: north → south", "change: second → last" }, edits) -- (in the quickfix list too)
  end)

  it("writes typed formatting as LaTeX, and keeps the spacing around typed words", function()
    local path = setup()
    H.replace(2, "We studied", "We *carefully* studied")
    H.replace(2, "the north", "the **far** north")
    local out = run(path)
    assert.truthy(out:find("We \\emph{carefully} studied \\emph{Vipera aspis} in the \\textbf{far} north of Iberia\n", 1, true))
  end)

  it("rebuilds a citation from which a reference was removed", function()
    local path = setup()
    H.replace(2, "; Jones, 2019", "")
    local out, issues = run(path)
    assert.equals(SOURCE:gsub("smith2020,jones2019", "smith2020"), out)
    assert.same({}, issues)
  end)

  it("looks up typed citations in the .bib, and lists those it can't find", function()
    local path = setup()
    H.replace(3, "in it.", "in it (Brown & Green, 2018; Martínez-Freiría et al., 2021).")
    local out = run(path)
    assert.truthy(out:find("in it \\parencite{brown2018,martinez2021}.", 1, true))
    setup()
    H.replace(3, "in it.", "in it (Nobody, 1999).")
    local _, issues = run(path)
    assert.truthy(has(issues, "citation not found in the .bib: Nobody 1999"))
  end)

  it("reads the .bib as it is written: entries on one line or many", function()
    local path = setup()
    H.replace(3, "in it.", "in it (Wild et al., 2025; Brown & Green, 2018).")
    local out = run(path)
    assert.truthy(out:find("in it \\parencite{wild2025,brown2018}.", 1, true))
    setup()
    H.replace(3, "in it.", "in it (Fake, 2000).") -- (in a comment of the .bib: not an entry)
    local _, notes = run(path)
    assert.truthy(has(notes, "citation not found in the .bib: Fake 2000"))
  end)

  it("keeps a computed value that was edited, and says so", function()
    local path = setup()
    H.replace(2, "12 percent", "15 percent")
    local out, issues = run(path)
    assert.equals(SOURCE, out)
    assert.truthy(has(issues, "a computed value changed in the document: the source keeps it (\\Sexpr{rise()})"))
  end)

  it("deletes a computed value with the sentence it was in", function()
    local path = setup()
    H.replace(2, " Warming raised its water loss by 12 percent.", "")
    local out, issues = run(path)
    assert.equals(SOURCE:gsub("\nWarming raised its water loss by \\Sexpr{rise%(%)} percent%.", ""), out)
    assert.same({}, issues)
  end)

  it("leaves what the source changed since the document was sent, if the document changed it too", function()
    local path = setup(nil, SOURCE:gsub("north of Iberia", "north of Spain"))
    H.replace(2, "of Iberia", "of the Peninsula")
    H.replace(3, "second", "last")
    local out, issues = run(path)
    assert.truthy(out:find("north of Spain", 1, true)) -- (not overwritten)
    assert.truthy(out:find("A last paragraph", 1, true)) -- (the rest is merged)
    assert.truthy(has(issues, "conflict"))
  end)

  it("lists the comments at their source line", function()
    local path = setup()
    local r = H.row(3)
    vim.fn.setpos("'<", { 0, r, 3, 0 })
    vim.fn.setpos("'>", { 0, r, 8, 0 })
    vim.fn.histadd("cmd", "")
    vim.cmd("normal! gvv")
    H.input = "Which one?"
    require("neoffice.comments").add(true)
    local out, issues = merge.compute(api.nvim_get_current_buf(), { path })
    assert.equals(SOURCE, out[path])
    assert.equals(1, #issues)
    assert.truthy(issues[1].text:find("Which one?", 1, true))
    assert.equals(7, issues[1].lnum)
    api.nvim_win_set_cursor(0, { r, 3 }) -- (resolved: dealt with already, not listed)
    require("neoffice.comments").resolve()
    _, issues = merge.compute(api.nvim_get_current_buf(), { path })
    assert.same({}, issues)
  end)

  it("adds a new paragraph after the one before it, and says so", function()
    local path = setup()
    local r = H.row(3)
    api.nvim_buf_set_lines(0, r, r, false, { "", "A *new* one." })
    require("neoffice.changes").after_edit(api.nvim_get_current_buf())
    local out, issues = run(path)
    assert.truthy(out:find("bold words} in it.\n\nA \\emph{new} one.", 1, true))
    assert.truthy(has(issues, "new paragraph"))
  end)

  describe("with the copy as sent", function()
    it("finds edits made without tracking", function()
      local path = setup(nil, nil, true)
      H.replace(2, "north", "south")
      H.replace(2, "; Jones, 2019", "")
      assert.equals(0, #api.nvim_buf_get_extmarks(0, neo.ns.change, 0, -1, {})) -- (nothing tracked)
      local out, issues, copy = merge.compute(api.nvim_get_current_buf(), { path })
      assert.equals(vim.fs.dirname(path) .. "/main.sent.docx", copy)
      assert.equals((SOURCE:gsub("north", "south"):gsub("smith2020,jones2019", "smith2020")), out[path])
      assert.same({ "edit", "edit" }, vim.tbl_map(function(i) return i.user_data end, issues))
    end)

    it("and without it, only tracked ones", function()
      local path = setup(nil, nil, true)
      H.replace(2, "north", "south")
      local out = merge.compute(api.nvim_get_current_buf(), { path }, { sent = false })
      assert.equals(SOURCE, out[path])
    end)

    it("deletes a paragraph deleted without tracking, and says so", function()
      local path = setup(nil, nil, true)
      local r = H.row(3)
      api.nvim_buf_set_lines(0, r - 2, r, false, {})
      local out, issues = run(path)
      assert.equals((SOURCE:gsub("\nA second paragraph[^\n]*\n", "")), out)
      assert.truthy(has(issues, "paragraph deleted in the document"))
    end)

    it("adds a paragraph added without tracking", function()
      local path = setup(nil, nil, true)
      api.nvim_buf_set_lines(0, -1, -1, false, { "", "The end, *really*." })
      local out = run(path)
      assert.truthy(out:find("in it.\n\nThe end, \\emph{really}.", 1, true))
    end)

    it("is the document given with the sources, and the sources those of the folder when none", function()
      local path = setup(nil, nil, true)
      local dir = vim.fs.dirname(path)
      vim.uv.fs_rename(dir .. "/main.sent.docx", dir .. "/old.docx")
      H.replace(2, "north", "south")
      local cwd = vim.fn.getcwd()
      vim.cmd.cd(dir)
      merge.run({ "old.docx" })
      vim.cmd.cd(cwd)
      assert.equals(2, #api.nvim_list_tabpages())
      local texts = vim.tbl_map(function(w) return table.concat(api.nvim_buf_get_lines(api.nvim_win_get_buf(w), 0, -1, false), "\n") end, api.nvim_tabpage_list_wins(0))
      assert.truthy(vim.iter(texts):any(function(t) return t:find("the south of Iberia", 1, true) end))
      vim.cmd("tabonly!")
    end)
  end)

  it("points a citation it couldn't look up at its text in the proposal", function()
    local path = setup()
    H.replace(3, "in it.", "in it (Nobody et al., 1999).")
    merge.run({ path })
    local item = vim.iter(vim.fn.getqflist()):find(function(i) return i.text:find("Nobody", 1, true) end)
    assert.truthy(item)
    local line = api.nvim_buf_get_lines(item.bufnr, item.lnum - 1, item.lnum, false)[1]
    assert.truthy(api.nvim_buf_get_name(item.bufnr):find("with the document's edits", 1, true))
    assert.equals("Nobody et al., 1999", line:sub(item.col, item.col + #"Nobody et al., 1999" - 1))
    vim.cmd("tabonly!")
  end)

  it("pairs a short paragraph, like a heading, in the file of the paragraphs around it", function()
    local path = setup()
    local other = vim.fs.dirname(path) .. "/other.tex"
    local text = "\\section{Introduction}\n\nA chapter about something else entirely, with no words in common at all.\n"
    vim.fn.writefile(vim.split(text, "\n"), other)
    H.replace(1, "Introduction", "Background")
    local out = merge.compute(api.nvim_get_current_buf(), { other, path })
    assert.equals(table.concat(vim.split(text, "\n"), "\n") .. "\n", out[other])
    assert.equals((SOURCE:gsub("{Introduction}", "{Background}")), out[path])
  end)

  it("done again, after taking its edits (all, or some), proposes only the rest", function()
    local path = setup()
    H.replace(2, "We studied", "We carefully studied") -- (inserted)
    H.replace(2, "north", "south") -- (replaced)
    H.replace(2, " of Iberia", "") -- (deleted)
    H.replace(2, "; Jones, 2019", "") -- (a citation edited)
    api.nvim_buf_set_lines(0, -1, -1, false, { "", "A new paragraph at the end." }) -- (added)
    require("neoffice.changes").after_edit(api.nvim_get_current_buf())
    local merged = run(path)
    local function write(text)
      local f = assert(io.open(path, "w"))
      f:write(text)
      f:close()
    end
    write(merged) -- (all taken)
    local again, issues, edits = run(path)
    assert.equals(merged, again)
    assert.same({}, edits) -- (nothing left to list)
    assert.falsy(has(issues, "conflict") or has(issues, "new paragraph"))
    local some = SOURCE:gsub("north", "south") -- (one taken)
    write(some)
    again, issues = run(path)
    assert.equals(merged, again)
    assert.falsy(has(issues, "conflict"))
  end)

  it("done again with changes taken and not written, refreshes its tab and keeps them", function()
    local path = setup()
    local doc = api.nvim_get_current_buf()
    H.replace(2, "north", "south")
    H.replace(3, "second", "last")
    merge.run({ path })
    assert.truthy(vim.fn.getqflist({ title = true }).title:find("2 edit(s)", 1, true))
    local source = vim.fn.bufnr(path)
    api.nvim_set_current_win(vim.fn.win_findbuf(source)[1])
    vim.cmd("1/north/diffget") -- (one change taken, as do does; not written)
    assert.is_true(vim.bo[source].modified)
    api.nvim_set_current_win(vim.fn.win_findbuf(doc)[1])
    merge.run({ path }) -- (again: no "save first", the same tab)
    assert.equals(2, #api.nvim_list_tabpages())
    assert.is_true(vim.bo[source].modified)
    assert.truthy(table.concat(api.nvim_buf_get_lines(source, 0, -1, false), "\n"):find("the south of Iberia", 1, true))
    local items = vim.fn.getqflist()
    assert.same({ "change: second → last" }, vim.tbl_map(function(i) return i.text end, items))
    local wins = vim.tbl_filter(function(w) return vim.wo[w].diff end, api.nvim_tabpage_list_wins(0))
    assert.equals(2, #wins)
    vim.cmd("tabonly!")
    vim.cmd("silent! %bwipeout!")
  end)

  it("follows what you take: the quickfix list keeps only what is left", function()
    local path = setup()
    H.replace(2, "north", "south")
    H.replace(3, "second", "last")
    merge.run({ path })
    local source = vim.fn.bufnr(path)
    local proposal = vim.iter(api.nvim_tabpage_list_wins(0)):map(api.nvim_win_get_buf):find(function(b) return b ~= source end)
    local function col(b) return api.nvim_win_get_position(vim.fn.win_findbuf(b)[1])[2] end
    assert.is_true(col(proposal) < col(source)) -- (the proposal on the left, yours on the right)
    local function edits()
      return vim.tbl_map(function(i) return i.text end, vim.tbl_filter(function(i) return i.user_data == "edit" end, vim.fn.getqflist()))
    end
    assert.same({ "change: north → south", "change: second → last" }, edits())
    api.nvim_set_current_win(vim.fn.win_findbuf(source)[1])
    vim.cmd("1/north/diffget") -- (as do does)
    vim.api.nvim_exec_autocmds("TextChanged", { buffer = source })
    assert.is_true(vim.wait(3000, function() return #edits() == 1 end))
    assert.same({ "change: second → last" }, edits())
    assert.truthy(vim.fn.getqflist({ title = true }).title:find("1 edit(s)", 1, true))
    vim.cmd("tabonly!")
    vim.cmd("silent! %bwipeout!")
  end)

  it("warns when nothing tells what changed (no tracked changes, no export)", function()
    local path = setup()
    local messages = {}
    local notify = vim.notify
    vim.notify = function(m, level) messages[#messages + 1] = { m, level } end
    merge.run({ path })
    vim.wait(100)
    vim.notify = notify
    assert.equals(vim.log.levels.WARN, messages[#messages][2])
    assert.truthy(messages[#messages][1]:find("export it was made from", 1, true))
  end)

  it("keeps its diff when going to the notes: each in the window that shows its file", function()
    local path = setup({
      "# Introduction",
      "",
      "We studied *Vipera aspis* in the north of Iberia (Smith et al., 2020; Jones, 2019). Warming raised its water loss by 12 percent.",
      "",
      "A second paragraph, with **bold words** in it.",
      "",
      "Written by the export alone, and nowhere in the source at all.",
    })
    local doc = api.nvim_get_current_buf()
    H.replace(2, "north", "south")
    H.replace(3, "in it.", "in it (Nobody et al., 1999).")
    H.replace(4, "alone", "only")
    local switchbuf = vim.o.switchbuf
    merge.run({ path })
    local tab, source = api.nvim_get_current_tabpage(), vim.fn.bufnr(path)
    local items = vim.fn.getqflist()
    local function find(text)
      for n, i in ipairs(items) do
        if i.text:find(text, 1, true) then
          return n
        end
      end
    end
    local function go(n) -- (<CR> on item n of the list)
      vim.cmd("copen")
      api.nvim_win_set_cursor(0, { n, 0 })
      vim.cmd("normal \r")
    end
    go(find("change: north"))
    assert.equals(source, api.nvim_get_current_buf())
    go(find("Nobody")) -- (in the proposal)
    assert.truthy(api.nvim_buf_get_name(0):find("with the document's edits", 1, true))
    assert.equals(tab, api.nvim_get_current_tabpage())
    vim.cmd(find("change: north") .. "cc") -- (:cc and :cnext too)
    assert.equals(source, api.nvim_get_current_buf())
    assert.equals(#items, find("not found in the source")) -- (those only in the document: last)
    go(#items) -- (in the document: its tab)
    assert.equals(doc, api.nvim_get_current_buf())
    vim.cmd(find("change: north") .. "cc") -- (and back to the diff, from there)
    assert.equals(tab, api.nvim_get_current_tabpage())
    assert.equals(source, api.nvim_get_current_buf())
    local wins = vim.tbl_filter(function(w) return vim.wo[w].diff end, api.nvim_tabpage_list_wins(tab))
    assert.equals(2, #wins) -- (the diff as it was)
    assert.same({ source }, vim.tbl_filter(function(b) return b == source end, vim.tbl_map(api.nvim_win_get_buf, wins)))
    vim.cmd("tabclose")
    vim.wait(100)
    assert.equals(switchbuf, vim.o.switchbuf) -- (yours again, with no merge tab left)
    vim.cmd("silent! %bwipeout!")
  end)

  it("needs a document buffer and a source", function()
    vim.cmd("enew")
    merge.run({ "x.tex" })
    assert.equals(1, #api.nvim_list_tabpages())
    setup()
    merge.memory = vim.fn.tempname() .. ".json" -- (none remembered)
    H.input = "" -- (and asked, no answer)
    merge.run({})
    assert.equals(1, #api.nvim_list_tabpages())
    assert.truthy(neo.docs[api.nvim_get_current_buf()])
  end)

  describe("the export it was made from", function()
    it("is found in the sources' folder or below, rather than an edited copy kept there", function()
      local path = setup(nil, nil, true)
      local dir = vim.fs.dirname(path)
      vim.fn.mkdir(dir .. "/toword", "p")
      vim.uv.fs_rename(dir .. "/main.sent.docx", dir .. "/toword/main.docx") -- (as an export leaves it)
      H.replace(2, "north", "south") -- (edits without tracking, of several rounds: accepted already)
      H.replace(3, "second", "last")
      H.save()
      assert(vim.uv.fs_copyfile(dir .. "/main.docx", dir .. "/collaborator.docx")) -- (a copy kept, edited)
      H.replace(2, "Iberia", "Spain")
      local out, issues, copy = merge.compute(api.nvim_get_current_buf(), { path })
      assert.equals(dir .. "/toword/main.docx", copy)
      assert.equals((SOURCE:gsub("north of Iberia", "south of Spain"):gsub("second", "last")), out[path])
      assert.equals(3, #issues)
    end)
  end)

  describe("its sources", function()
    it("are those it merged into before, for a document opened anywhere", function()
      local path = setup()
      merge.memory = vim.fn.tempname() .. ".json"
      merge.compute(api.nvim_get_current_buf(), { path })
      merge.run({ path }) -- (remembered)
      vim.cmd("tabonly!")
      local elsewhere = vim.fn.tempname()
      vim.fn.mkdir(elsewhere, "p")
      local doc = elsewhere .. "/main_reviewed.docx"
      H.reopen()
      assert(vim.uv.fs_copyfile(H.doc().path, doc))
      H.open(doc)
      vim.cmd("Neoffice track on")
      H.replace(2, "north", "south")
      local cwd = vim.fn.getcwd()
      vim.cmd.cd(elsewhere)
      merge.run({})
      vim.cmd.cd(cwd)
      assert.equals(2, #api.nvim_list_tabpages())
      assert.truthy(vim.fn.bufnr(path) > 0 and vim.fn.win_findbuf(vim.fn.bufnr(path))[1])
      vim.cmd("tabonly!")
    end)

    it("are found when remembered as a file, and as its folder too (the same manuscript)", function()
      local path = setup()
      merge.memory = vim.fn.tempname() .. ".json"
      local dir = vim.fs.dirname(path)
      local other = dir .. "/appendix.tex"
      vim.fn.writefile({ "An appendix, with words of its own." }, other)
      merge.run({ path }) -- (the file)
      vim.cmd("tabonly!")
      merge.run({ dir }) -- (and the folder: both files)
      vim.cmd("tabonly!")
      H.replace(2, "north", "south")
      local asked = false
      local input = vim.ui.input
      vim.ui.input = function() asked = true end
      local cwd = vim.fn.getcwd()
      vim.cmd.cd(vim.fn.tempname():match("^(.*)/"))
      merge.run({})
      vim.cmd.cd(cwd)
      vim.ui.input = input
      assert.is_false(asked)
      assert.equals(2, #api.nvim_list_tabpages())
      vim.cmd("tabonly!")
    end)

    it("aren't those of another manuscript on the same subject", function()
      local path = setup()
      merge.memory = vim.fn.tempname() .. ".json"
      merge.run({ path }) -- (remembered)
      vim.cmd("tabonly!")
      local other = vim.fn.tempname() .. "/other.docx"
      vim.fn.mkdir(vim.fs.dirname(other), "p")
      H.open(other)
      api.nvim_buf_set_lines(0, 0, -1, false, { "Vipers lose water in dry air, and Iberia is warming: a second paragraph of another paper." })
      H.save()
      local asked = false
      local input = vim.ui.input
      vim.ui.input = function() asked = true end
      local cwd = vim.fn.getcwd()
      vim.cmd.cd(vim.fs.dirname(other))
      merge.run({})
      vim.cmd.cd(cwd)
      vim.ui.input = input
      assert.is_true(asked)
      assert.equals(1, #api.nvim_list_tabpages())
    end)

    it("are asked for when none is known: a folder gives its sources", function()
      local path = setup()
      merge.memory = vim.fn.tempname() .. ".json"
      H.replace(2, "north", "south")
      local cwd = vim.fn.getcwd()
      vim.cmd.cd(vim.fn.tempname():match("^(.*)/"))
      H.input = vim.fs.dirname(path)
      merge.run({})
      vim.cmd.cd(cwd)
      assert.equals(2, #api.nvim_list_tabpages())
      local texts = vim.tbl_map(function(w) return table.concat(api.nvim_buf_get_lines(api.nvim_win_get_buf(w), 0, -1, false), "\n") end, api.nvim_tabpage_list_wins(0))
      assert.truthy(vim.iter(texts):any(function(t) return t:find("the south of Iberia", 1, true) end))
      vim.cmd("tabonly!")
      H.input = "typed text"
    end)
  end)

  it("opens the source and the proposal in a diff, and the notes in the quickfix list", function()
    local path = setup()
    H.replace(2, "north", "south")
    H.replace(2, "12 percent", "15 percent")
    merge.run({ path })
    assert.equals(2, #api.nvim_list_tabpages())
    local wins = vim.tbl_filter(function(w) return vim.wo[w].diff end, api.nvim_tabpage_list_wins(0))
    assert.equals(2, #wins)
    local texts = vim.tbl_map(function(w) return table.concat(api.nvim_buf_get_lines(api.nvim_win_get_buf(w), 0, -1, false), "\n") .. "\n" end, wins)
    assert.truthy(vim.tbl_contains(texts, SOURCE))
    assert.truthy(vim.tbl_contains(texts, (SOURCE:gsub("north", "south"))))
    local items = vim.fn.getqflist()
    assert.equals(2, #items) -- (the edit, and the computed value kept)
    local edit = vim.iter(items):find(function(i) return i.user_data == "edit" end)
    assert.equals("change: north → south", edit.text)
    assert.same({ 3, 39 }, { edit.lnum, edit.col })
    vim.cmd("tabonly!")
  end)
end)
