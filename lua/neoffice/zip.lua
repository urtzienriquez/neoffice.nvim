-- .docx and .odt files are ZIP archives; read and replace their entries with unzip/zip.
local M = {}

function M.read(path, entry)
  -- "[" is a wildcard for unzip, so "[Content_Types].xml" needs escaping
  local r = vim.system({ "unzip", "-p", path, (entry:gsub("%[", "[[]")) }):wait()
  return r.code == 0 and r.stdout ~= "" and r.stdout or nil
end

function M.list(path)
  local r = vim.system({ "unzip", "-Z1", path }):wait()
  local names = {}
  for name in r.stdout:gmatch("[^\n]+") do
    names[name] = true
  end
  return names
end

-- Replace or add entries ({ [name] = content }) and write the result to `target` (default: `path`).
-- Works on a copy that is renamed over the target at the end, so a failure never damages the file.
function M.write(path, entries, target)
  target = target or path
  if next(entries) == nil then -- nothing changed: leave the file alone (or just copy it)
    if target ~= path then
      assert(vim.uv.fs_copyfile(path, target .. ".neoffice~"))
      assert(os.rename(target .. ".neoffice~", target))
    end
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
  local tmp = target .. ".neoffice~"
  assert(vim.uv.fs_copyfile(path, tmp))
  local r = vim.system({ "zip", "-q", "-X", "-nw", tmp, unpack(names) }, { cwd = dir }):wait()
  vim.fn.delete(dir, "rf")
  if r.code ~= 0 then
    os.remove(tmp)
    error("zip failed: " .. r.stderr)
  end
  assert(os.rename(tmp, target))
end

return M
