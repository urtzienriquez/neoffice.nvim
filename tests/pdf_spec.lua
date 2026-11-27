-- PDF exports (:Neoffice pdf; pdf.lua): next to the document, from it as saved, in the background,
-- again after each save when asked; the viewer isn't opened here (config.pdf.open = false).
local H = dofile("tests/helpers.lua")
local neo = require("neoffice")
local api = vim.api

local function said_by(fn)
  local said, notify = {}, vim.notify
  vim.notify = function(msg) said[#said + 1] = msg end
  local ok, err = pcall(fn)
  vim.notify = notify
  assert(ok, err)
  return table.concat(said, "\n")
end
local function stamp(path)
  local s = vim.uv.fs_stat(path)
  return s and (s.mtime.sec * 1e9 + s.mtime.nsec) or nil
end
local function changed(path, before)
  return vim.wait(20000, function() local now = stamp(path) return now ~= nil and now ~= before end, 10)
end
local function head(path)
  local f = assert(io.open(path, "rb"))
  local bytes = f:read(5)
  f:close()
  return bytes
end

if not H.has_libreoffice then
  describe("PDF export", function()
    pending("needs LibreOffice")
  end)
  return
end

describe("PDF export", function()
  before_each(function()
    neo.config.pdf = { auto = false, open = false, quiet = true, timeout = 120 }
  end)

  for _, ext in ipairs({ "docx", "odt" }) do
    it("writes the document as saved next to it, nothing else there (" .. ext .. ")", function()
      local path = H.copy("rich." .. ext)
      local pdf = path:gsub("%." .. ext .. "$", ".pdf")
      H.open(path)
      vim.cmd("Neoffice pdf")
      assert.truthy(changed(pdf, nil), "no PDF")
      assert.equals("%PDF-", head(pdf))
      local names = vim.fn.readdir(vim.fn.fnamemodify(path, ":h"))
      table.sort(names)
      assert.same({ vim.fn.fnamemodify(path, ":t"), vim.fn.fnamemodify(pdf, ":t") }, names) -- (no lock file, no part of one)
    end)
  end

  it("writes it again after each save when asked, and not once turned off", function()
    local path = H.copy("rich.docx")
    local pdf = path:gsub("%.docx$", ".pdf")
    H.open(path)
    vim.cmd("Neoffice pdf auto on")
    assert.truthy(changed(pdf, nil), "no PDF when turned on")
    local before = stamp(pdf)
    H.set(1, "Edited once.")
    H.save()
    assert.truthy(changed(pdf, before), "not written again after a save")
    vim.cmd("Neoffice pdf auto off")
    before = stamp(pdf)
    H.set(1, "Edited twice.")
    H.save()
    vim.wait(1500)
    assert.equals(before, stamp(pdf), "written again though turned off")
  end)

  it("exports the save made while one is being exported, once, after it", function()
    local path = H.copy("rich.docx")
    local pdf = path:gsub("%.docx$", ".pdf")
    H.open(path)
    local pdfmod = require("neoffice.pdf")
    local done = 0
    neo.config.pdf.quiet = false
    local notify = vim.notify
    vim.notify = function(msg) if msg:find("neoffice: PDF ", 1, true) then done = done + 1 end end
    pdfmod.export(path, false)
    pdfmod.export(path, false) -- (while the first is being written: next, once)
    pdfmod.export(path, false)
    local ok = vim.wait(20000, function() return done >= 2 end, 10)
    vim.wait(1500)
    vim.notify = notify
    assert.truthy(ok, "not written")
    assert.equals(2, done)
    assert.equals("%PDF-", head(pdf))
  end)

  it("isn't written for a copy saved under another name (:w other.docx)", function()
    local path = H.copy("rich.docx")
    H.open(path)
    vim.b.neoffice_pdf_auto = true
    local other = vim.fn.fnamemodify(path, ":h") .. "/other.docx"
    vim.cmd("silent write " .. vim.fn.fnameescape(other))
    vim.wait(1500)
    assert.is_nil(vim.uv.fs_stat((path:gsub("%.docx$", ".pdf"))))
    assert.is_nil(vim.uv.fs_stat((other:gsub("%.docx$", ".pdf"))))
  end)

  it("reaches LibreOffice through a pipe only you can use, not a port anyone on the machine could", function()
    local path = H.copy("rich.docx")
    H.open(path)
    vim.cmd("Neoffice pdf")
    assert.truthy(changed((path:gsub("%.docx$", ".pdf")), nil))
    local f = io.open(vim.fn.stdpath("cache") .. "/neoffice-pdf/office.pid")
    local pid = f and tonumber(f:read("*a"))
    if f then
      f:close()
    end
    assert.truthy(pid, "LibreOffice's process")
    local cmdline = table.concat(vim.split(io.open("/proc/" .. pid .. "/cmdline"):read("*a"), "\0", { plain = true }), " ")
    local name = cmdline:match("%-%-accept=pipe,name=([%w_]+);urp;")
    assert.truthy(name, cmdline)
    assert.is_nil(cmdline:find("socket", 1, true))
    local pipes = vim.fn.glob("/tmp/OSL_PIPE_*" .. name, false, true)
    for _, pipe in ipairs(pipes) do -- (yours alone: no one else can connect to it)
      assert.equals("rwx------", vim.fn.getfperm(pipe))
    end
  end)

  it("says the unsaved changes aren't in it", function()
    local path = H.copy("rich.docx")
    H.open(path)
    H.set(1, "Not saved.")
    local said = said_by(function() vim.cmd("Neoffice pdf") end)
    assert.truthy(said:find("unsaved changes", 1, true), said)
    assert.truthy(changed((path:gsub("%.docx$", ".pdf")), nil))
  end)
end)
