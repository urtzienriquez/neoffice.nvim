-- The :Neoffice command, options, and .doc files.
local H = dofile("tests/helpers.lua")
local neo = require("neoffice")
local api = vim.api

describe(":Neoffice", function()
  it("completes subcommands and their arguments", function()
    assert.same({ "accept" }, vim.fn.getcompletion("Neoffice acc", "cmdline"))
    assert.same({ "on", "off" }, vim.fn.getcompletion("Neoffice track ", "cmdline"))
    assert.same({ "all" }, vim.fn.getcompletion("Neoffice reject ", "cmdline"))
    assert.same({ "auto" }, vim.fn.getcompletion("Neoffice pdf ", "cmdline"))
    assert.same({ "on", "off" }, vim.fn.getcompletion("Neoffice pdf auto ", "cmdline"))
    assert.equals(14, #vim.fn.getcompletion("Neoffice ", "cmdline"))
  end)

  it("runs subcommands", function()
    H.open(H.copy("rich.odt"))
    vim.cmd("Neoffice accept all")
    assert.equals(0, #H.changes())
  end)

  it("reports an unknown subcommand, and outside documents does nothing", function()
    vim.cmd("enew")
    assert.has_no.errors(function()
      vim.cmd("Neoffice nonsense")
      vim.cmd("Neoffice accept")
      vim.cmd("Neoffice track")
      vim.cmd("Neoffice comments")
    end)
  end)

  it("sets no keymaps but g? (its help) in documents, and none with keymaps = false", function()
    local before = #api.nvim_get_keymap("n")
    H.open(H.copy("rich.docx"))
    local function ours(mode) -- (not Neovim's own, for a language server: K for hover...)
      return vim.tbl_map(function(m) return m.lhs end,
        vim.tbl_filter(function(m) return not (m.desc or ""):match("^vim%.lsp") end, api.nvim_buf_get_keymap(0, mode)))
    end
    assert.same({ "g?" }, ours("n"))
    assert.same({}, ours("x"))
    assert.equals(before, #api.nvim_get_keymap("n"))
    local defaults = vim.deepcopy(neo.config.keymaps)
    neo.setup({ keymaps = false })
    H.open(H.copy("rich.docx"))
    assert.same({}, ours("n"))
    neo.config.keymaps = defaults
  end)
end)

describe("options", function()
  after_each(function()
    neo.config.conceal = false
  end)

  it("don't conceal the marks by default", function()
    vim.wo.conceallevel = 0
    H.open(H.copy("rich.docx"))
    assert.equals(0, vim.wo.conceallevel)
  end)

  it("conceal them with conceal = true", function()
    neo.setup({ conceal = true })
    H.open(H.copy("rich.docx"))
    assert.equals(2, vim.wo.conceallevel)
  end)

  it("keep the panel keys configurable (and those set as panel_keys before)", function()
    local defaults = vim.deepcopy(neo.config.keymaps)
    neo.setup({ keymaps = { comments = { delete = false, reply = "R" } } })
    H.open(H.copy("rich.docx"))
    vim.cmd("Neoffice comments")
    local keys = vim.tbl_map(function(m) return m.lhs end, api.nvim_buf_get_keymap(0, "n"))
    assert.is_true(vim.tbl_contains(keys, "R"))
    assert.is_false(vim.tbl_contains(keys, "dd"))
    assert.is_true(vim.tbl_contains(keys, "gq") and vim.tbl_contains(keys, "q")) -- (a list: both)
    vim.cmd("Neoffice comments")
    neo.setup({ panel_keys = { jump = "o", delete = "x" } }) -- (as it was set before: jump is go now)
    assert.same({ "o", "x" }, { neo.config.keymaps.comments.go, neo.config.keymaps.comments.delete })
    neo.config.keymaps = defaults
  end)

  it("open the help with g?, in a document and in the comments panel (at its part)", function()
    H.open(H.copy("rich.docx"))
    local doc_win = api.nvim_get_current_win()
    vim.cmd("normal g?")
    assert.equals("help", vim.bo.buftype)
    assert.truthy(vim.fn.getline("."):find("*neoffice*", 1, true), vim.fn.getline("."))
    vim.cmd("close")
    api.nvim_set_current_win(doc_win)
    vim.cmd("Neoffice comments")
    vim.cmd("normal g?")
    assert.equals("help", vim.bo.buftype)
    assert.truthy(vim.fn.getline("."):find("*neoffice-comments-panel*", 1, true), vim.fn.getline("."))
    vim.cmd("close")
    vim.cmd("Neoffice comments")
  end)
end)

describe(".doc", function()
  it("opens and saves through LibreOffice", function()
    if not H.has_libreoffice then
      return pending("needs LibreOffice")
    end
    local dir = vim.fn.tempname()
    vim.system({ "soffice", "-env:UserInstallation=file://" .. dir .. "/p", "--headless", "--convert-to", "doc", "--outdir", dir, H.FIXTURES .. "rich.docx" }):wait()
    local path = dir .. "/rich.doc"
    local before = H.bytes(path)
    H.open(path)
    assert.equals("# Introduction", H.paras()[1])
    H.save() -- (nothing changed: nothing written)
    assert.equals(before, H.bytes(path))
    H.set(12, "Changed in a .doc.")
    H.reopen()
    assert.equals("Changed in a .doc.", H.paras()[12])
  end)

  it("opens at once the next time: the conversion kept, and what was written read back", function()
    if not H.has_libreoffice then
      return pending("needs LibreOffice")
    end
    local neo = require("neoffice")
    local cache, was = vim.fn.tempname(), neo.config.doc_cache
    neo.config.doc_cache = cache
    local dir = vim.fn.tempname()
    vim.system({ "soffice", "-env:UserInstallation=file://" .. dir .. "/p", "--headless", "--convert-to", "doc", "--outdir", dir, H.FIXTURES .. "rich.docx" }):wait()
    local path = dir .. "/rich.doc"
    H.open(path)
    local function kept()
      local n = 0
      for _ in vim.fs.dir(cache) do
        n = n + 1
      end
      return n
    end
    assert.equals(1, kept())
    local system = vim.system
    local function without_libreoffice(fn) -- (each conversion would fail)
      vim.system = function(cmd, ...)
        assert(cmd[1] ~= "soffice", "LibreOffice was run")
        return system(cmd, ...)
      end
      local ok, err = pcall(fn)
      vim.system = system
      assert(ok, err)
    end
    without_libreoffice(function() H.open(path) end)
    assert.equals("# Introduction", H.paras()[1])
    H.set(12, "Changed in a .doc.")
    H.save()
    assert(vim.wait(60000, function() return kept() == 2 end, 50), "what was written was never read back")
    without_libreoffice(function() H.open(path) end)
    assert.equals("Changed in a .doc.", H.paras()[12])
    neo.config.doc_cache = was
  end)
end)
