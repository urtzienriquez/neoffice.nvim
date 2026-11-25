-- The :Neoffice command, options, and .doc files.
local H = dofile("tests/helpers.lua")
local neo = require("neoffice")
local api = vim.api

describe(":Neoffice", function()
  it("completes subcommands and their arguments", function()
    assert.same({ "accept" }, vim.fn.getcompletion("Neoffice acc", "cmdline"))
    assert.same({ "on", "off" }, vim.fn.getcompletion("Neoffice track ", "cmdline"))
    assert.same({ "all" }, vim.fn.getcompletion("Neoffice reject ", "cmdline"))
    assert.equals(13, #vim.fn.getcompletion("Neoffice ", "cmdline"))
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

  it("sets no keymaps", function()
    local before = #api.nvim_get_keymap("n")
    H.open(H.copy("rich.docx"))
    assert.same({}, api.nvim_buf_get_keymap(0, "n"))
    assert.same({}, api.nvim_buf_get_keymap(0, "x"))
    assert.equals(before, #api.nvim_get_keymap("n"))
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

  it("keep the panel keys configurable", function()
    neo.setup({ panel_keys = { delete = false, reply = "R" } })
    H.open(H.copy("rich.docx"))
    vim.cmd("Neoffice comments")
    local keys = vim.tbl_map(function(m) return m.lhs end, api.nvim_buf_get_keymap(0, "n"))
    assert.is_true(vim.tbl_contains(keys, "R"))
    assert.is_false(vim.tbl_contains(keys, "d"))
    neo.setup({ panel_keys = { delete = "d", reply = "r" } })
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
end)
