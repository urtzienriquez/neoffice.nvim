-- What a document from someone else can't make neoffice do: take all the memory, be read though
-- damaged, write outside the file, run a script through a converter, read files or fetch from the
-- net through pandoc, run macros in LibreOffice.
local H = dofile("tests/helpers.lua")
local zip = require("neoffice.zip")

-- the file with the central directory's record of `entry` changed at `at` (bytes after its
-- signature) to the 4 bytes of `value`
local function patched(name, entry, at, value)
  local path = H.copy(name)
  local f = assert(io.open(path, "rb"))
  local s = f:read("*a")
  f:close()
  local pos = 1
  while true do
    local i = s:find("PK\1\2", pos, true)
    assert(i, "no record of " .. entry)
    local nlen = s:byte(i + 28) + s:byte(i + 29) * 256
    if s:sub(i + 46, i + 45 + nlen) == entry then
      local bytes = string.char(value % 256, math.floor(value / 256) % 256, math.floor(value / 65536) % 256, math.floor(value / 16777216) % 256)
      s = s:sub(1, i + at - 1) .. bytes .. s:sub(i + at + 4)
      break
    end
    pos = i + 4
  end
  f = assert(io.open(path, "wb"))
  f:write(s)
  f:close()
  return path
end

describe("a damaged or made-up file", function()
  if not zip.crc32("x") then
    pending("needs zlib (through LuaJIT's FFI)")
    return
  end

  it("saying a part is huge unpacked isn't believed: nothing that big is made, it says so", function()
    local path = patched("rich.docx", "word/document.xml", 24, 0x30000000) -- (768 MiB)
    local ok, err = pcall(zip.read, path, "word/document.xml")
    assert.is_false(ok)
    assert.truthy(tostring(err):find("too big", 1, true), err)
  end)

  it("whose part isn't what its checksum says isn't read as if it were", function()
    local path = patched("rich.docx", "word/document.xml", 16, 12345)
    local ok, err = pcall(zip.read, path, "word/document.xml")
    assert.is_false(ok)
    assert.truthy(tostring(err):find("checksum", 1, true), err)
    -- (opened: it says the file can't be read, and Neovim goes on)
    local said = {}
    local notify = vim.notify
    vim.notify = function(m) said[#said + 1] = m end
    pcall(H.open, path)
    vim.notify = notify
    assert.is_nil(require("neoffice").docs[vim.api.nvim_get_current_buf()])
  end)

  it("a character that isn't one (&#99999999999;) opens as one that says so", function()
    local xml = require("neoffice.xml")
    local doc = xml.parse("<a>&#99999999999;&#xD800;&#65;</a>")
    assert.equals("\u{FFFD}\u{FFFD}A", doc.children[1].children[1].text)
  end)
end)

describe("writing a file's parts", function()
  it("never writes a part whose name would be a path out of it", function()
    local path = H.copy("rich.docx")
    local before = H.bytes(path)
    local dir = vim.fn.fnamemodify(path, ":h")
    for _, name in ipairs({ "../evil.xml", "/tmp/evil.xml", "word/../../evil.xml", "word\\..\\evil.xml", "" }) do
      local ok, err = pcall(zip.write, path, { [name] = "x" })
      assert.is_false(ok, name)
      assert.truthy(tostring(err):find("not a name for a part", 1, true), err)
    end
    assert.equals(before, H.bytes(path))
    assert.is_nil(vim.uv.fs_stat(vim.fn.fnamemodify(dir, ":h") .. "/evil.xml"))
    zip.write(path, { ["word/new.xml"] = "<a/>" }) -- (a name in it: written)
    assert.equals("<a/>", zip.read(path, "word/new.xml"))
  end)
end)

describe("image previews", function()
  local images = require("neoffice.images")
  it("tell ImageMagick what each file is, and give it no other kinds (a script in disguise)", function()
    local cmd = images.convert_command("png", "/x/a.png", "/x/o.png")
    if cmd and cmd[1] == "magick" then
      assert.truthy(vim.tbl_contains(cmd, "PNG:/x/a.png[0]"), vim.inspect(cmd))
    end
    for _, kind in ipairs({ "mvg", "msl", "ps", "eps", "txt", "html" }) do
      assert.is_nil(images.convert_command(kind, "/x/a." .. kind, "/x/o.png"), kind)
    end
    for _, kind in ipairs({ "pdf", "svg" }) do -- (never through ImageMagick)
      local c = images.convert_command(kind, "/x/a." .. kind, "/x/o.png")
      assert.truthy(not c or c[1] ~= "magick", vim.inspect(c))
    end
  end)
end)

describe("equations through pandoc", function()
  if not H.has_pandoc then
    pending("needs pandoc")
    return
  end
  it("are converted with --sandbox: no file read, nothing fetched", function()
    local calls, system = {}, vim.system
    vim.system = function(cmd, ...)
      if cmd[1] == "pandoc" then
        calls[#calls + 1] = cmd
      end
      return system(cmd, ...)
    end
    local ok, err = pcall(function()
      local path = vim.fn.tempname() .. ".docx"
      H.open(path)
      vim.api.nvim_buf_set_lines(0, 0, -1, false, { "Energy is $E = mc^{2}$." })
      H.reopen()
    end)
    vim.system = system
    assert(ok, err)
    assert.truthy(#calls > 0)
    for _, cmd in ipairs(calls) do -- (each that reads the text, or a document: sandboxed; the one writing the .docx reads pandoc's own JSON)
      local from = cmd[vim.fn.index(cmd, "-f") + 2]
      assert.truthy(from == "json" or vim.tbl_contains(cmd, "--sandbox"), vim.inspect(cmd))
    end
    assert.truthy(vim.iter(calls):any(function(c) return vim.tbl_contains(c, "--sandbox") end))
  end)

  it("aren't written when their LaTeX would make more than an equation (an image, from a file)", function()
    local docx = require("neoffice.docx")
    local ok, err = pcall(docx.equations, {}, { { tex = "x\\) ![a](/etc/hostname) \\(y", display = false } })
    assert.is_false(ok)
    assert.truthy(tostring(err):find("more than an equation", 1, true), err)
  end)
end)

describe("LibreOffice's profiles", function()
  local office = require("neoffice.office")
  it("are made safe for documents from others: no macros, no active content, no links followed", function()
    local dir = vim.fn.tempname()
    local arg = office.profile(dir)
    assert.equals("-env:UserInstallation=file://" .. dir, arg)
    local file = dir .. "/user/registrymodifications.xcu"
    local text = table.concat(vim.fn.readfile(file), "\n")
    local WANT = { -- (each, as it must be: not taken from office.SETTINGS, so none can go unnoticed)
      MacroSecurityLevel = "3", DisableMacrosExecution = "true", DisableActiveContent = "true",
      DisableOLEAutomation = "true", BlockUntrustedRefererLinks = "true",
    }
    for name, value in pairs(WANT) do
      assert.truthy(text:find('<item oor:path="/org.openoffice.Office.Common/Security/Scripting"><prop oor:name="' .. name .. '" oor:op="fuse"><value>' .. value .. "</value>", 1, true), name)
    end
    assert.is_true(require("neoffice.xml").check(text))
  end)

  it("keep LibreOffice's own settings, and put back a safe value where it was changed", function()
    local dir = vim.fn.tempname()
    vim.fn.mkdir(dir .. "/user", "p")
    local file = dir .. "/user/registrymodifications.xcu"
    vim.fn.writefile({ '<?xml version="1.0" encoding="UTF-8"?>',
      '<oor:items xmlns:oor="http://openoffice.org/2001/registry" xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">',
      '<item oor:path="/org.openoffice.Office.Common/Misc"><prop oor:name="FirstRun" oor:op="fuse"><value>false</value></prop></item>',
      '<item oor:path="/org.openoffice.Office.Common/Security/Scripting"><prop oor:name="MacroSecurityLevel" oor:op="fuse"><value>0</value></prop></item>',
      "</oor:items>" }, file)
    office.profile(dir)
    local text = table.concat(vim.fn.readfile(file), "\n")
    assert.truthy(text:find('FirstRun" oor:op="fuse"><value>false', 1, true)) -- (its own: kept)
    assert.is_nil(text:find("<value>0</value>", 1, true)) -- (the unsafe level: gone)
    local _, n = text:gsub('oor:name="MacroSecurityLevel"', "")
    assert.equals(1, n)
    local stamp = vim.uv.fs_stat(file).mtime
    office.profile(dir) -- (as it should be: not written again)
    assert.same(stamp, vim.uv.fs_stat(file).mtime)
  end)
end)
