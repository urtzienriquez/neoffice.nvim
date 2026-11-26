-- Images: their type and size, and showing them below their ![alt](path) line with vim.ui.img
-- (Neovim 0.12+, in terminals that support the Kitty graphics protocol: Kitty, Ghostty, WezTerm...).
-- Previews are made in the background (read from the document, made small, other formats turned
-- into PNG: PDF, SVG, EMF...), so opening a document never waits for them.
local api = vim.api
local M = {}

local ns = api.nvim_create_namespace("neoffice_images")

local function be(s, i, n) -- big-endian number of n bytes at i
  local v = 0
  for k = i, i + n - 1 do
    v = v * 256 + s:byte(k)
  end
  return v
end

-- "png" | "jpeg" | "gif", width, height (pixels); nil for anything else.
function M.info(data)
  if data:sub(1, 8) == "\137PNG\r\n\26\n" and #data >= 24 then
    return "png", be(data, 17, 4), be(data, 21, 4)
  elseif data:sub(1, 4) == "GIF8" and #data >= 10 then
    return "gif", data:byte(7) + data:byte(8) * 256, data:byte(9) + data:byte(10) * 256
  elseif data:sub(1, 2) == "\255\216" then -- JPEG: find the frame header (SOFn)
    local i = 3
    while i + 8 <= #data and data:byte(i) == 255 do
      local marker, size = data:byte(i + 1), be(data, i + 2, 2)
      if marker >= 0xC0 and marker <= 0xCF and marker ~= 0xC4 and marker ~= 0xC8 and marker ~= 0xCC then
        return "jpeg", be(data, i + 7, 2), be(data, i + 5, 2)
      end
      i = i + 2 + size
    end
  end
end

-- ── Showing them ────────────────────────────────────────────────────────────

local cache = {} -- path (and modification time, for files on disk) -> PNG bytes, or false
local pending = {} -- the same keys, while being read or converted (in the background)
local shown = {} -- buf -> { key -> image id }
local supported -- (asked once)

local function available()
  if supported == nil then
    local img = vim.ui and vim.ui.img
    local ok, yes = false, false
    if img and #api.nvim_list_uis() > 0 then
      ok, yes = pcall(img._supported, { timeout = 500 })
    end
    supported = ok and yes == true
  end
  return supported and require("neoffice").config.images
end

-- How images are turned into PNG to be shown, by their extension (or kind): a command for a file
-- `src` and the PNG `out` to write; nil if the program isn't there. A preview needs no more than
-- MAX pixels (large photos are also read at a smaller size: much faster).
local MAX = 1000
local tools = {}
local function has(program)
  if tools[program] == nil then
    tools[program] = vim.fn.executable(program) == 1
  end
  return tools[program]
end
local function magick(src, out)
  return has("magick") and { "magick", "-define", "jpeg:size=" .. 2 * MAX .. "x" .. 2 * MAX, src .. "[0]", "-thumbnail", MAX .. "x" .. MAX .. ">", "png:" .. out } or nil
end
local CONVERT = {
  png = magick, jpeg = magick, jpg = magick, gif = magick, bmp = magick, tif = magick, tiff = magick, webp = magick,
  pdf = function(src, out)
    return has("pdftoppm") and { "pdftoppm", "-png", "-singlefile", "-f", "1", "-l", "1", "-scale-to", tostring(MAX), src, (out:gsub("%.png$", "")) }
      or magick(src, out)
  end,
  svg = function(src, out)
    return has("rsvg-convert") and { "rsvg-convert", "-w", tostring(MAX), "-o", out, src } or magick(src, out)
  end,
}
-- (drawings: metafiles and others, through LibreOffice: all those waiting in one call)
local OFFICE = { emf = true, wmf = true, emz = true, wmz = true, svm = true, eps = true, cgm = true, pct = true, pict = true, met = true }

local function done(key, data)
  pending[key] = nil
  cache[key] = data and M.info(data) == "png" and data or false
  for buf in pairs(shown) do
    if api.nvim_buf_is_valid(buf) then
      M.refresh(buf)
    end
  end
end

local function read(file)
  local f = io.open(file, "rb")
  local data = f and f:read("*a")
  if f then
    f:close()
  end
  return data
end

-- LibreOffice's conversions: queued, one call at a time (it doesn't run twice at once).
local office = { queue = {}, running = false }
local function run_office()
  if office.running or #office.queue == 0 then
    return
  end
  local batch = office.queue
  office.queue, office.running = {}, true
  local dir = vim.fn.tempname()
  vim.fn.mkdir(dir, "p")
  local cmd = { "soffice", "-env:UserInstallation=file://" .. vim.fn.stdpath("cache") .. "/neoffice-libreoffice", "--headless", "--convert-to", "png", "--outdir", dir }
  for _, item in ipairs(batch) do
    cmd[#cmd + 1] = item.src
  end
  vim.system(cmd, { timeout = 60000 }, function()
    vim.schedule(function()
      for _, item in ipairs(batch) do
        local out = vim.fs.joinpath(dir, vim.fn.fnamemodify(item.src, ":t:r") .. ".png")
        if has("magick") and vim.uv.fs_stat(out) then -- (drawn on a page: without the page around it)
          vim.system({ "magick", out, "-trim", "+repage", "-thumbnail", MAX .. "x" .. MAX .. ">", "png:" .. out }, { timeout = 30000 }, function()
            vim.schedule(function() item.finish(read(out)) end)
          end)
        else
          item.finish(read(out))
        end
      end
      office.running = false
      run_office()
    end)
  end)
end

-- Turn an image's bytes (of a file with extension `ext`) into a preview: PNG bytes, or nil if it
-- can't be; `finish` gets them (in the background, but for a small PNG: at once).
function M.preview(data, ext, finish)
  local kind = M.info(data)
  if kind == "png" and #data < 1500000 then -- (as it is: nothing to wait for)
    return finish(data)
  end
  ext = kind or ext
  local src = vim.fn.tempname() .. "." .. ext
  local out = vim.fn.tempname() .. ".png"
  local f = io.open(src, "wb")
  if not f then
    return finish(nil)
  end
  f:write(data)
  f:close()
  if OFFICE[ext] and has("soffice") then
    office.queue[#office.queue + 1] = { finish = finish, src = src }
    return run_office()
  end
  local cmd = CONVERT[ext] and CONVERT[ext](src, out)
  if not cmd then
    return finish(kind == "png" and data or nil)
  end
  vim.system(cmd, { timeout = 30000 }, function(r)
    vim.schedule(function()
      finish(r.code == 0 and read(out) or nil)
    end)
  end)
end

-- PNG bytes of an image of the document (as written in the buffer), if it is ready: if not, it is
-- read and converted in the background, and shown when it is (nil meanwhile).
local function png(buf, path)
  local doc = require("neoffice").docs[buf]
  local inside = doc.ctx.entries[path]
  local file = vim.fs.normalize(path)
  file = file:match("^/") and file or vim.fs.joinpath(vim.fs.dirname(doc.path), file)
  local stat = not inside and vim.uv.fs_stat(file)
  local key = inside and doc.path .. "//" .. path or file .. "@" .. (stat and stat.mtime.sec or "")
  if cache[key] == nil and not pending[key] then
    pending[key] = true
    local ext = (path:match("%.(%w+)$") or ""):lower()
    if inside then
      vim.system({ "unzip", "-p", doc.base, path }, {}, function(r)
        vim.schedule(function()
          if r.code == 0 and r.stdout and #r.stdout > 0 then
            M.preview(r.stdout, ext, function(data) done(key, data) end)
          else
            done(key, nil)
          end
        end)
      end)
    else
      local data = stat and read(file)
      if data then
        M.preview(data, ext, function(png_data) done(key, png_data) end)
      else
        done(key, nil)
      end
    end
  end
  return cache[key] or nil
end

-- Size in cells: as wide as allowed, keeping the aspect (a cell is about twice as high as wide).
function M.cells(width, height, max_cols, max_rows)
  local cols = math.max(1, math.min(max_cols, math.floor(width / 8)))
  local rows = math.max(1, math.floor(cols * height / width / 2 + 0.5))
  if rows > max_rows then
    rows, cols = max_rows, math.max(1, math.floor(max_rows * 2 * width / height))
  end
  return cols, rows
end

local function clear(buf)
  for _, id in pairs(shown[buf] or {}) do
    vim.ui.img.del(id)
  end
  shown[buf] = {}
end

-- Reserve lines under each image and draw the visible ones over them.
local function draw(buf)
  if not (api.nvim_buf_is_valid(buf) and available()) then
    return
  end
  local config = require("neoffice").config
  api.nvim_buf_clear_namespace(buf, ns, 0, -1)
  local wanted = {}
  for row, line in ipairs(api.nvim_buf_get_lines(buf, 0, -1, false)) do
    local pos = line:find("![", 1, true)
    while pos do
      local _, path, after = require("neoffice.markup").image_at(line, pos)
      local data = path and png(buf, path)
      if data then
        local _, w, h = M.info(data)
        local cols, rows = M.cells(w, h, config.image_width, config.image_height)
        local blank = {}
        for k = 1, rows do
          blank[k] = { { " " } }
        end
        api.nvim_buf_set_extmark(buf, ns, row - 1, 0, { virt_lines = blank })
        table.insert(wanted, { row = row, data = data, cols = cols, rows = rows, key = row .. ":" .. path })
        break -- (one image per line is shown)
      end
      pos = line:find("![", after or pos + 2, true)
    end
  end
  vim.schedule(function() -- (once the reserved lines are on screen)
    local old, new = shown[buf] or {}, {}
    -- floating windows (pickers, hovers...): an image they cover is hidden while they are open, as
    -- images are drawn on top of everything (a cell more around them, for their borders)
    local floats = {}
    for _, w in ipairs(api.nvim_tabpage_list_wins(0)) do
      local config = api.nvim_win_get_config(w)
      if config.relative ~= "" and not config.hide then
        local top, left = unpack(vim.fn.win_screenpos(w))
        floats[#floats + 1] = { win = w, top = top - 1, left = left - 1,
          bottom = top + api.nvim_win_get_height(w), right = left + api.nvim_win_get_width(w) }
      end
    end
    local function covered(win, top, left, rows, cols)
      return vim.iter(floats):any(function(f)
        return f.win ~= win and top <= f.bottom and top + rows - 1 >= f.top and left <= f.right and left + cols - 1 >= f.left
      end)
    end
    for _, win in ipairs(vim.fn.win_findbuf(buf)) do
      local info = vim.fn.getwininfo(win)[1]
      if info.tabnr == vim.fn.tabpagenr() then
        for _, image in ipairs(wanted) do
          local line = api.nvim_buf_get_lines(buf, image.row - 1, image.row, false)[1] or ""
          local pos = vim.fn.screenpos(win, image.row, math.max(#line, 1))
          local top = pos.row + 1 -- (the reserved lines start right under the line)
          local left = info.wincol + info.textoff
          if pos.row > 0 and top + image.rows - 1 <= info.winrow + info.height - 1 and not covered(win, top, left, image.rows, image.cols) then
            local key = win .. ":" .. image.key
            local opts = { row = top, col = left, width = image.cols, height = image.rows }
            new[key] = old[key] and vim.ui.img.set(old[key], opts) or vim.ui.img.set(image.data, opts)
            old[key] = nil
          end
        end
      end
    end
    for _, id in pairs(old) do
      vim.ui.img.del(id)
    end
    shown[buf] = new
  end)
end

local refreshing = {}
function M.refresh(buf)
  if not refreshing[buf] then
    refreshing[buf] = true
    vim.defer_fn(function()
      refreshing[buf] = nil
      draw(buf)
    end, 50)
  end
end

function M.attach(buf, group)
  local refresh = function() M.refresh(buf) end
  api.nvim_create_autocmd({ "BufWinEnter", "TextChanged", "InsertLeave", "BufWritePost" }, { group = group, buffer = buf, callback = refresh })
  api.nvim_create_autocmd({ "WinScrolled", "WinResized", "VimResized", "TabEnter", "WinNew", "WinClosed", "WinEnter" }, { group = group, callback = refresh })
  api.nvim_create_autocmd({ "BufWinLeave", "BufHidden", "BufWipeout" }, {
    group = group,
    buffer = buf,
    callback = function()
      if vim.ui.img then
        clear(buf)
      end
    end,
  })
  refresh()
end

return M
