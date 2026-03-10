-- neoffice/zip.lua
-- Helpers for reading/writing entries inside ZIP archives (.docx/.odt are ZIPs).
-- Uses system `unzip` + `zip` binaries. No LuaRocks required.

local M = {}

local function sq(s)
  return "'" .. tostring(s):gsub("'", "'\"'\"'") .. "'"
end

function M.read_entry(zip_path, entry)
  local cmd = "unzip -p " .. sq(zip_path) .. " " .. sq(entry) .. " 2>/dev/null"
  local out = vim.fn.system(cmd)
  if vim.v.shell_error ~= 0 then
    return nil
  end
  if out == "" then
    return nil
  end
  return out
end

function M.list_entries(zip_path)
  local cmd = "unzip -Z1 " .. sq(zip_path) .. " 2>/dev/null"
  local out = vim.fn.system(cmd)
  local entries = {}
  for line in out:gmatch("[^\n]+") do
    if line:match("^[%w/]") then
      table.insert(entries, line)
    end
  end
  return entries
end

function M.has_entry(zip_path, entry)
  for _, e in ipairs(M.list_entries(zip_path)) do
    if e == entry then
      return true
    end
  end
  return false
end

function M.write_entry(zip_path, entry, content)
  local abs_zip = vim.fn.fnamemodify(zip_path, ":p")
  local tmp_dir = vim.fn.tempname()
  vim.fn.mkdir(tmp_dir, "p")

  local dir_part = entry:match("^(.+)/[^/]+$")
  if dir_part then
    vim.fn.mkdir(tmp_dir .. "/" .. dir_part, "p")
  end

  local tmp_file = tmp_dir .. "/" .. entry
  local lines = vim.split(content, "\n", { plain = true })
  vim.fn.writefile(lines, tmp_file)

  if vim.fn.filereadable(tmp_file) ~= 1 then
    vim.fn.system("rm -rf " .. sq(tmp_dir))
    return false
  end

  local cmd = "cd " .. sq(tmp_dir) .. " && zip -u " .. sq(abs_zip) .. " " .. sq(entry) .. " 2>&1"
  local out = vim.fn.system(cmd)
  local exit_code = vim.v.shell_error

  vim.fn.system("rm -rf " .. sq(tmp_dir))

  if exit_code ~= 0 then
    vim.notify(string.format("[neoffice] ZIP: zip command failed (exit %d): %s", exit_code, out), vim.log.levels.ERROR)
    return false
  end
  return true
end

return M
