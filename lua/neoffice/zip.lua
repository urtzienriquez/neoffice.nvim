-- .docx and .odt files are ZIP archives: their entries read, and replaced or added.
--
-- In Neovim itself, with zlib (through LuaJIT's FFI): an archive's directory is read once (while
-- the file stays as it was), entries are inflated as they are asked for, and writing copies every
-- entry that didn't change byte for byte (its header, its data) and compresses only those that
-- did. Without zlib, or for an archive it doesn't handle (ZIP64, encrypted entries), unzip and zip
-- do it, as before.
local M = {}

local ffi_ok, ffi = pcall(require, "ffi")
local bit = require("bit")
local z -- (zlib, once loaded; false if it can't be)

local function zlib()
  if z == nil then
    z = false
    if ffi_ok then
      local ok = pcall(ffi.cdef, [[
        typedef struct neoffice_z_stream_s {
          const unsigned char *next_in; unsigned int avail_in; unsigned long total_in;
          unsigned char *next_out; unsigned int avail_out; unsigned long total_out;
          const char *msg; void *state; void *zalloc; void *zfree; void *opaque;
          int data_type; unsigned long adler; unsigned long reserved;
        } neoffice_z_stream;
        const char *zlibVersion(void);
        int inflateInit2_(neoffice_z_stream *strm, int windowBits, const char *version, int stream_size);
        int inflate(neoffice_z_stream *strm, int flush);
        int inflateEnd(neoffice_z_stream *strm);
        int deflateInit2_(neoffice_z_stream *strm, int level, int method, int windowBits, int memLevel, int strategy, const char *version, int stream_size);
        int deflate(neoffice_z_stream *strm, int flush);
        int deflateEnd(neoffice_z_stream *strm);
        unsigned long deflateBound(neoffice_z_stream *strm, unsigned long sourceLen);
        unsigned long crc32(unsigned long crc, const unsigned char *buf, unsigned int len);
      ]])
      if ok then
        for _, name in ipairs({ "z", "libz.so.1", "libz.so", "libz.dylib" }) do
          local loaded, lib = pcall(ffi.load, name)
          if loaded then
            z = lib
            break
          end
        end
      end
    end
  end
  return z or nil
end

local function u16(s, i)
  local a, b = s:byte(i, i + 1)
  return a + b * 256
end
local function u32(s, i)
  local a, b, c, d = s:byte(i, i + 3)
  return a + b * 256 + c * 65536 + d * 16777216
end
local function le16(n)
  return string.char(n % 256, math.floor(n / 256) % 256)
end
local function le32(n)
  return string.char(n % 256, math.floor(n / 256) % 256, math.floor(n / 65536) % 256, math.floor(n / 16777216) % 256)
end

-- An archive: its bytes and its entries ({ name, method, flags, crc, csize, size, offset (of its
-- local header, 1-based), central (its directory record) }, in the directory's order). nil if it
-- isn't one this reads (then unzip does).
local cache = {} -- (the last archive read: { key, archive })
local function archive(path)
  local stat = vim.uv.fs_stat(path)
  if not stat then
    return nil
  end
  local key = ("%s:%d:%d.%d"):format(path, stat.size, stat.mtime.sec, stat.mtime.nsec)
  if cache.key == key then
    return cache.archive
  end
  local f = io.open(path, "rb")
  if not f then
    return nil
  end
  local s = f:read("*a")
  f:close()
  -- the end of the central directory: its signature, within the last 64 KiB (a comment may follow)
  local eocd
  for i = #s - 21, math.max(1, #s - 65557), -1 do
    if s:byte(i) == 0x50 and u32(s, i) == 0x06054b50 then
      eocd = i
      break
    end
  end
  if not eocd then
    return nil
  end
  local count, size, start = u16(s, eocd + 10), u32(s, eocd + 12), u32(s, eocd + 16) + 1
  if count == 0xFFFF or size == 0xFFFFFFFF or start == 0x100000000 then
    return nil -- (ZIP64)
  end
  local entries, p = {}, start
  for _ = 1, count do
    if u32(s, p) ~= 0x02014b50 then
      return nil
    end
    local flags, method = u16(s, p + 8), u16(s, p + 10)
    local nlen, xlen, clen = u16(s, p + 28), u16(s, p + 30), u16(s, p + 32)
    local e = {
      flags = flags, method = method, crc = u32(s, p + 16), csize = u32(s, p + 20), size = u32(s, p + 24),
      offset = u32(s, p + 42) + 1, name = s:sub(p + 46, p + 45 + nlen),
      central = s:sub(p, p + 45 + nlen + xlen + clen),
    }
    if bit.band(flags, 1) ~= 0 or (method ~= 0 and method ~= 8) or e.csize == 0xFFFFFFFF or e.size == 0xFFFFFFFF then
      return nil -- (encrypted, another compression, ZIP64)
    end
    entries[#entries + 1] = e
    p = p + 46 + nlen + xlen + clen
  end
  local a = { bytes = s, entries = entries, start = start, eocd = eocd }
  cache.key, cache.archive = key, a
  return a
end

local function data_start(s, e) -- (where an entry's data begins: after its local header)
  return e.offset + 30 + u16(s, e.offset + 26) + u16(s, e.offset + 28)
end

-- The most an entry may unpack to: what it says it is, read before anything is made of that size
-- (a file made to say more would take all the memory: a document's parts are a few MiB).
M.MAX_ENTRY = 512 * 1024 * 1024

local function inflate(src, size)
  local lib = zlib()
  assert(size <= M.MAX_ENTRY, ("a part of the file says it is %d MiB unpacked: too big for a document's part (the file may be damaged, or made to look so)"):format(math.floor(size / 1048576)))
  local out = ffi.new("unsigned char[?]", math.max(size, 1))
  local strm = ffi.new("neoffice_z_stream[1]")
  assert(lib.inflateInit2_(strm, -15, lib.zlibVersion(), ffi.sizeof("neoffice_z_stream")) == 0, "inflate")
  strm[0].next_in, strm[0].avail_in = ffi.cast("const unsigned char *", src), #src
  strm[0].next_out, strm[0].avail_out = out, size
  local r = lib.inflate(strm, 4) -- (Z_FINISH)
  local got = tonumber(strm[0].total_out)
  lib.inflateEnd(strm)
  assert(r == 1 or (r == -5 and size == 0), "inflate failed") -- (Z_STREAM_END)
  assert(got == size, "a part of the file isn't as long as it says: the file is damaged")
  return ffi.string(out, got)
end

local function deflate(data)
  local lib = zlib()
  local strm = ffi.new("neoffice_z_stream[1]")
  assert(lib.deflateInit2_(strm, 6, 8, -15, 8, 0, lib.zlibVersion(), ffi.sizeof("neoffice_z_stream")) == 0, "deflate")
  local bound = tonumber(lib.deflateBound(strm, #data)) + 64
  local out = ffi.new("unsigned char[?]", bound)
  strm[0].next_in, strm[0].avail_in = ffi.cast("const unsigned char *", data), #data
  strm[0].next_out, strm[0].avail_out = out, bound
  local r = lib.deflate(strm, 4) -- (Z_FINISH)
  local n = tonumber(strm[0].total_out)
  lib.deflateEnd(strm)
  assert(r == 1, "deflate failed")
  return ffi.string(out, n)
end

local function crc32(data)
  local lib = zlib()
  return tonumber(lib.crc32(0, ffi.cast("const unsigned char *", data), #data))
end
M.crc32 = function(data) return zlib() and crc32(data) end

function M.read(path, entry)
  local a = zlib() and archive(path)
  if a then
    for _, e in ipairs(a.entries) do
      if e.name == entry then
        local from = data_start(a.bytes, e)
        local raw = a.bytes:sub(from, from + e.csize - 1)
        assert(#raw == e.csize, "a part of the file is cut short: the file is damaged")
        local data = e.method == 0 and raw or inflate(raw, e.size)
        assert(crc32(data) == e.crc, ("a part of the file (%s) isn't what its checksum says: the file is damaged"):format(entry))
        return data ~= "" and data or nil
      end
    end
    return nil
  end
  -- "[" is a wildcard for unzip, so "[Content_Types].xml" needs escaping
  local r = vim.system({ "unzip", "-p", path, (entry:gsub("%[", "[[]")) }):wait()
  return r.code == 0 and r.stdout ~= "" and r.stdout or nil
end

function M.list(path)
  local a = zlib() and archive(path)
  local names = {}
  if a then
    for _, e in ipairs(a.entries) do
      names[e.name] = true
    end
    return names
  end
  local r = vim.system({ "unzip", "-Z1", path }):wait()
  for name in r.stdout:gmatch("[^\n]+") do
    names[name] = true
  end
  return names
end

-- The archive with these entries replaced or added ({ [name] = content }): its bytes, or nil if
-- this can't make it (then zip does).
local function rebuilt(path, entries)
  local a = zlib() and archive(path)
  if not a then
    return nil
  end
  local s = a.bytes
  local order = {} -- (the entries by where they are: each one's bytes run to the next one's)
  for _, e in ipairs(a.entries) do
    order[#order + 1] = e
  end
  table.sort(order, function(x, y) return x.offset < y.offset end)
  local stop = {}
  for k, e in ipairs(order) do
    stop[e] = (order[k + 1] and order[k + 1].offset or a.start) - 1
  end
  local now = os.date("*t")
  local time = bit.bor(bit.lshift(now.hour, 11), bit.lshift(now.min, 5), math.floor(now.sec / 2))
  local date = bit.bor(bit.lshift(math.max(now.year - 1980, 0), 9), bit.lshift(now.month, 5), now.day)
  local parts, central, offset, seen = {}, {}, 0, {}
  local function add(local_bytes, central_record)
    parts[#parts + 1] = local_bytes
    -- (the central record, pointing at where its entry now is)
    central[#central + 1] = central_record:sub(1, 42) .. le32(offset) .. central_record:sub(47)
    offset = offset + #local_bytes
  end
  local function fresh(name, content, old)
    local data = deflate(content)
    local method = 8
    if #data >= #content then -- (stored, when compressing doesn't help)
      data, method = content, 0
    end
    local crc = crc32(content)
    local utf8 = name:find("[\128-\255]") and 0x0800 or 0
    local head = le32(0x04034b50) .. le16(20) .. le16(utf8) .. le16(method) .. le16(time) .. le16(date)
      .. le32(crc) .. le32(#data) .. le32(#content) .. le16(#name) .. le16(0) .. name
    local rec = le32(0x02014b50) .. le16(old and u16(old.central, 5) or 20) .. le16(20) .. le16(utf8) .. le16(method)
      .. le16(time) .. le16(date) .. le32(crc) .. le32(#data) .. le32(#content) .. le16(#name) .. le16(0) .. le16(0)
      .. le16(0) .. le16(0) .. le32(old and u32(old.central, 39) or 0) .. le32(0) .. name
    add(head .. data, rec)
  end
  for _, e in ipairs(a.entries) do -- (in the directory's order, as they were)
    if entries[e.name] ~= nil then
      fresh(e.name, entries[e.name], e)
    else
      add(s:sub(e.offset, stop[e]), e.central)
    end
    seen[e.name] = true
  end
  local new = {}
  for name in pairs(entries) do
    if not seen[name] then
      new[#new + 1] = name
    end
  end
  table.sort(new)
  for _, name in ipairs(new) do
    fresh(name, entries[name])
  end
  local dir = table.concat(central)
  local count = #a.entries + #new
  local tail = le32(0x06054b50) .. le16(0) .. le16(0) .. le16(count) .. le16(count) .. le32(#dir) .. le32(offset)
    .. le16(u16(s, a.eocd + 20)) .. s:sub(a.eocd + 22, a.eocd + 21 + u16(s, a.eocd + 20))
  return table.concat(parts) .. dir .. tail
end

-- A file put in place of `target` only once it is all there, as it should be: written beside it
-- (`tmp`), read back and checked (check(path) -> true, or nil and why), then renamed over it.
-- Through a symbolic link, the file it points to; with the permissions the file had.
local function real(target)
  return vim.uv.fs_realpath(target) or target
end
local function put(tmp, target, check)
  local ok, why = check(tmp)
  if not ok then
    os.remove(tmp)
    error(("the file written (%s) isn't right: %s; nothing was changed"):format(tmp, why or "?"))
  end
  local was = vim.uv.fs_stat(target)
  if was then
    vim.uv.fs_chmod(tmp, bit.band(was.mode, 4095))
  end
  local renamed, err = os.rename(tmp, target)
  if not renamed then
    os.remove(tmp)
    error(("could not put %s in place: %s"):format(target, err or "?"))
  end
end
local function bytes_of(path)
  local f = io.open(path, "rb")
  if not f then
    return nil
  end
  local data = f:read("*a")
  f:close()
  return data
end
local function write_file(path, data) -- (all of it, or an error: a full disk shows when closing too)
  local f, err = io.open(path, "wb")
  if not f then
    error(("could not write %s: %s"):format(path, err or "?"))
  end
  local ok, werr = f:write(data)
  local closed, cerr = f:close()
  if not ok or not closed then
    os.remove(path)
    error(("could not write %s: %s"):format(path, werr or cerr or "?"))
  end
end

-- Copy a file over another, as put() does: the copy compared with the file, byte for byte.
function M.copy(path, target)
  target = real(target)
  local data = assert(bytes_of(path), "could not read " .. path)
  local tmp = target .. ".neoffice~"
  write_file(tmp, data)
  put(tmp, target, function(file)
    local back = bytes_of(file)
    return back == data, "it isn't the same as what was copied"
  end)
end

-- (an archive written: it reads, has the entries it should, and those written read back as they were)
local function readable(file, bytes, entries)
  local back = bytes_of(file)
  if back ~= bytes then
    return nil, "it isn't what was written (a full disk?)"
  end
  cache.key = nil -- (read again: not the archive this was made from)
  local a = archive(file)
  if not a then
    return nil, "it can't be read as a zip archive"
  end
  for name, content in pairs(entries) do
    local ok, got = pcall(M.read, file, name)
    if not ok or (got or "") ~= content then
      return nil, ("%s doesn't read back as it was written"):format(name)
    end
  end
  return true
end

-- Replace or add entries ({ [name] = content }) and write the result to `target` (default: `path`).
-- Works on a copy that is renamed over the target at the end, so a failure never damages the file.
function M.write(path, entries, target)
  for name in pairs(entries) do -- (a part's name, inside the archive: never a path out of it)
    assert(type(name) == "string" and name ~= "" and not name:find("^/") and not name:find("\\", 1, true)
      and not ("/" .. name .. "/"):find("/%.%.?/"), ("not a name for a part of the file: %q"):format(tostring(name)))
  end
  target = real(target or path)
  if next(entries) == nil then -- nothing changed: leave the file alone (or just copy it)
    if target ~= real(path) then
      M.copy(path, target)
    end
    return
  end
  local tmp = target .. ".neoffice~"
  local ok, bytes = pcall(rebuilt, path, entries)
  if ok and bytes then
    write_file(tmp, bytes)
    put(tmp, target, function(file) return readable(file, bytes, entries) end)
    return
  end
  local dir, names = vim.fn.tempname(), {}
  for name, content in pairs(entries) do
    local file = dir .. "/" .. name
    vim.fn.mkdir(vim.fs.dirname(file), "p")
    local f = assert(io.open(file, "wb"))
    f:write(content)
    f:close()
    names[#names + 1] = name
  end
  assert(vim.uv.fs_copyfile(path, tmp))
  local r = vim.system({ "zip", "-q", "-X", "-nw", tmp, unpack(names) }, { cwd = dir }):wait()
  vim.fn.delete(dir, "rf")
  if r.code ~= 0 then
    os.remove(tmp)
    error("zip failed: " .. r.stderr)
  end
  put(tmp, target, function(file) -- (as unzip reads it)
    local t = vim.system({ "unzip", "-tq", file }):wait()
    if t.code ~= 0 then
      return nil, "unzip finds it broken"
    end
    for name, content in pairs(entries) do
      local got = vim.system({ "unzip", "-p", file, (name:gsub("%[", "[[]")) }):wait()
      if got.code ~= 0 or got.stdout ~= content then
        return nil, ("%s doesn't read back as it was written"):format(name)
      end
    end
    return true
  end)
end

return M
