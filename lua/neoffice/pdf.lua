-- PDF exports: :Neoffice pdf writes the document, as saved, as a PDF next to it (report.docx ->
-- report.pdf) and opens it; :Neoffice pdf auto [on|off] does it again after each save. A viewer that
-- reloads a file when it changes (zathura, Okular, Evince, sioyek, SumatraPDF...) then shows each
-- save as it is made.
--
-- One LibreOffice, started at the first export and kept running until Neovim quits, converts them
-- (see pdf_server.py, through LibreOffice's Python bridge): an export takes a fraction of a second,
-- not LibreOffice's start. Without that bridge, each export starts LibreOffice (slower). Exports run
-- in the background: typing and saving never wait for one; a save made while its document is being
-- exported is exported next, once, the last one.
local M = {}
local api = vim.api

local function neo()
  return require("neoffice")
end

local function config()
  return neo().config.pdf or {}
end

-- The PDF of a document: next to it, with its name.
function M.path_of(doc_path)
  return vim.fn.fnamemodify(doc_path, ":p:r") .. ".pdf"
end

local here = debug.getinfo(1, "S").source:sub(2):match("^(.*)/[^/]*$") or "."
local cache = vim.fn.stdpath("cache")

-- The server (pdf_server.py): { proc, waiting = { [id] = on_done }, partial = "" }; false when
-- LibreOffice's Python bridge isn't there (each export then starts LibreOffice).
local server
local next_id = 0

local function stop_server()
  if server then
    pcall(function() server.proc:write(nil) end) -- (stdin closed: it stops LibreOffice and quits)
    local proc = server.proc
    vim.defer_fn(function() pcall(function() proc:kill(15) end) end, 6000)
  end
  server = nil
end

local function answer(line)
  local ok, msg = pcall(vim.json.decode, line)
  if not ok or type(msg) ~= "table" then
    return
  end
  if type(msg.error) == "string" and msg.error:match("^no%-uno") then
    server = false -- (no Python bridge: the slower way, from now on)
    return
  end
  local on_done = server and server.waiting[msg.id]
  if on_done then
    server.waiting[msg.id] = nil
    on_done(msg.ok, msg.ok and msg.ms or msg.error)
  end
end

local function start_server()
  if vim.fn.executable("python3") == 0 then
    server = false
    return
  end
  local s = { waiting = {}, partial = "" }
  s.proc = vim.system({ "python3", "-I", here .. "/pdf_server.py", cache .. "/neoffice-libreoffice-pdf", cache .. "/neoffice-pdf" }, {
    stdin = true,
    stdout = function(_, data)
      if not data then
        return
      end
      vim.schedule(function()
        s.partial = s.partial .. data
        while true do
          local line, rest = s.partial:match("^([^\n]*)\n(.*)$")
          if not line then
            break
          end
          s.partial = rest
          answer(line)
        end
      end)
    end,
  }, function(r)
    vim.schedule(function()
      if server == s then -- (it went away: what it was doing failed; the next export starts it again)
        server = nil
        for _, on_done in pairs(s.waiting) do
          on_done(false, "LibreOffice stopped (" .. tostring(r.code) .. ")")
        end
      end
    end)
  end)
  server = s
end

-- (the slower way: LibreOffice started for each export, from a copy in neoffice's cache)
local function convert_alone(src, dst, on_done)
  local work = cache .. "/neoffice-pdf/alone"
  vim.fn.mkdir(work, "p")
  local copy = work .. "/document." .. (src:match("%.(%w+)$") or "docx"):lower()
  if not vim.uv.fs_copyfile(src, copy) then
    return on_done(false, "couldn't copy " .. src)
  end
  local t0 = vim.uv.hrtime()
  vim.system({ "soffice", "-env:UserInstallation=file://" .. cache .. "/neoffice-libreoffice-pdf", "--headless", "--norestore",
    "--convert-to", "pdf", "--outdir", work, copy }, {}, vim.schedule_wrap(function(r)
    local out = work .. "/document.pdf"
    os.remove(copy)
    if r.code ~= 0 or not vim.uv.fs_stat(out) then
      return on_done(false, "LibreOffice (soffice) couldn't convert it")
    end
    local part = vim.fn.fnamemodify(dst, ":h") .. "/." .. vim.fn.fnamemodify(dst, ":t") .. ".neoffice-part"
    local ok = vim.uv.fs_copyfile(out, part) and os.rename(part, dst)
    os.remove(out)
    os.remove(part)
    on_done(ok and true or false, ok and math.floor((vim.uv.hrtime() - t0) / 1e6) or "couldn't write " .. dst)
  end))
end

local function convert(src, dst, on_done)
  if server == nil then
    start_server()
  end
  if not server then
    return convert_alone(src, dst, on_done)
  end
  next_id = next_id + 1
  local id = next_id
  server.waiting[id] = on_done
  local ok = pcall(function() server.proc:write(vim.json.encode({ id = id, src = src, dst = dst }) .. "\n") end)
  if not ok then
    server.waiting[id] = nil
    stop_server()
    return convert_alone(src, dst, on_done)
  end
  -- (it never answers: a LibreOffice stuck on a document; it is stopped, the export failed)
  vim.defer_fn(function()
    if server and server.waiting[id] then
      server.waiting[id] = nil
      on_done(false, "LibreOffice took too long")
      local proc = server.proc
      server = nil
      pcall(function() proc:kill(9) end)
    end
  end, (config().timeout or 120) * 1000)
end

-- path -> { running = bool, again = { open = bool } or nil, opened = bool }
local state = {}

local function open_viewer(pdf)
  local viewer = config().viewer
  if type(viewer) == "string" and viewer ~= "" then
    vim.system({ viewer, pdf }, { detach = true })
  elseif type(viewer) == "table" then
    local cmd = vim.list_extend(vim.deepcopy(viewer), { pdf })
    vim.system(cmd, { detach = true })
  else
    vim.ui.open(pdf)
  end
end

-- Export the document `path` (a .docx, .odt or .doc, as saved) to its PDF; `open`: open it once
-- written (else only the first time it is exported this session, if config.pdf.open).
function M.export(path, open)
  path = vim.fn.fnamemodify(path, ":p")
  local s = state[path] or {}
  state[path] = s
  if s.running then -- (one at a time: this one when that one is done, the last asked)
    s.again = { open = open or (s.again and s.again.open) }
    return
  end
  s.running = true
  local pdf = M.path_of(path)
  convert(path, pdf, function(ok, info)
    s.running = false
    if ok then
      local show = open or (not s.opened and config().open ~= false)
      if show then
        s.opened = true
        open_viewer(pdf)
      end
      if not config().quiet then
        vim.notify(("neoffice: PDF %s (%.1f s)"):format(vim.fn.fnamemodify(pdf, ":~:."), (tonumber(info) or 0) / 1000))
      end
    else
      vim.notify("neoffice: The PDF couldn't be written: " .. tostring(info), vim.log.levels.WARN)
    end
    if s.again then
      local again = s.again
      s.again = nil
      M.export(path, again.open)
    end
  end)
end

-- Whether a document is exported after each save (b:neoffice_pdf_auto; else config.pdf.auto).
function M.auto(buf)
  local own = vim.b[buf].neoffice_pdf_auto
  if own ~= nil then
    return own
  end
  return config().auto == true
end

-- After a document is saved (see M.save in init.lua): exported, if it is to be.
function M.saved(buf, path)
  if M.auto(buf) then
    M.export(path, false)
  end
end

-- :Neoffice pdf [auto [on|off]]
function M.command(arg)
  local buf = api.nvim_get_current_buf()
  local doc = neo().docs[buf]
  if not doc then
    return vim.notify("neoffice: Not a document buffer", vim.log.levels.WARN)
  end
  local sub, value = (arg or ""):match("^(%S*)%s*(%S*)")
  if sub == "auto" then
    local on
    if value == "on" then
      on = true
    elseif value == "off" then
      on = false
    else
      on = not M.auto(buf)
    end
    vim.b[buf].neoffice_pdf_auto = on
    vim.notify("neoffice: PDF after each save " .. (on and "on" or "off"))
    if on then
      M.export(doc.path, not (state[vim.fn.fnamemodify(doc.path, ":p")] or {}).opened)
    end
    return
  elseif sub ~= "" then
    return vim.notify("neoffice: :Neoffice pdf [auto [on|off]]", vim.log.levels.WARN)
  end
  if vim.bo[buf].modified then
    vim.notify("neoffice: The PDF is of the document as saved: its unsaved changes aren't in it (:w, then :Neoffice pdf)", vim.log.levels.WARN)
  end
  M.export(doc.path, true)
end

api.nvim_create_autocmd("VimLeavePre", {
  group = api.nvim_create_augroup("neoffice_pdf", {}),
  callback = function()
    if server then -- (it stops LibreOffice, then quits: waited for a moment, so that it can)
      pcall(function() server.proc:write(nil) end)
      pcall(function() server.proc:wait(3000) end)
    end
  end,
})

return M
