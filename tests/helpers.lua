-- Shared helpers for the specs: open fixtures, edit like a user, save, and look inside the files.
local neo = require("neoffice")
local zip = require("neoffice.zip")
local xml = require("neoffice.xml")
local api = vim.api
local H = {}

H.FIXTURES = vim.fn.getcwd() .. "/tests/fixtures/"
H.MAIN = { docx = "word/document.xml", odt = "content.xml" }

-- the answers vim.ui.input and confirm() give (tests set H.input before acting)
H.input = "typed text"
vim.ui.input = function(_, callback) callback(H.input) end
vim.fn.confirm = function() return 1 end
vim.notify = function() end -- (keeps the test output readable)

-- A copy of a fixture in a temporary folder. edit(name, text) -> text may change its main parts.
function H.copy(name, edit)
  local dir = vim.fn.tempname()
  vim.fn.mkdir(dir, "p")
  local path = dir .. "/" .. name
  assert(vim.uv.fs_copyfile(H.FIXTURES .. name, path))
  if edit then
    local entries = {}
    for _, entry in ipairs({ "word/document.xml", "word/styles.xml", "word/footnotes.xml", "content.xml" }) do
      local text = zip.read(path, entry)
      if text then
        entries[entry] = edit(entry, text)
      end
    end
    zip.write(path, entries)
  end
  return path
end

neo.config.doc_cache = vim.fn.tempname() -- (.doc files' conversions: not in your cache)
neo.config.styles = false -- (the styles at the top: tested on their own, styles_spec.lua)
neo.config.page = false -- (and the page's settings: pages_spec.lua)

-- A document's plain paragraphs all of its default style (as Word's often are, not the fixtures'
-- First Paragraph and Body Text, pandoc's): for H.copy's edit, `text = H.default_styles(name, text)`.
function H.default_styles(name, text)
  if name == "word/document.xml" then
    return (text:gsub('<w:pStyle w:val="BodyText" ?/>', ""):gsub('<w:pStyle w:val="FirstParagraph" ?/>', ""))
  elseif name == "content.xml" then
    return (text:gsub('text:style%-name="Text_20_body"', 'text:style-name="Standard"'):gsub('text:style%-name="First_20_Paragraph"', 'text:style-name="Standard"'))
  end
  return text
end

function H.open(path)
  vim.cmd("silent! %bwipeout!")
  vim.cmd("edit " .. vim.fn.fnameescape(path))
  vim.wait(0) -- (what opening leaves for once the text shows: its colours)
  return api.nvim_get_current_buf()
end

function H.doc()
  return neo.docs[api.nvim_get_current_buf()]
end

function H.lines()
  return api.nvim_buf_get_lines(0, 0, -1, false)
end

-- A table's line as it would be written unaligned ("| a | b |", "|---|---:|"), others as they are.
function H.compact(line)
  local markup = require("neoffice.markup")
  local format = line:match("^|") and line:find("---", 1, true) and markup.table_format(line) -- (not "| - |")
  if format then
    return markup.format_line(format, select(2, line:gsub("|", "")) - 1)
  end
  local spans = line:match("^|") and markup.cell_spans(line) or {}
  if #spans == 0 then
    return line
  end
  return "| " .. table.concat(vim.tbl_map(function(sp) return line:sub(sp.s, sp.e) end, spans), " | ") .. " |"
end
function H.compact_lines()
  return vim.tbl_map(H.compact, H.lines())
end

function H.paras() -- the paragraphs' texts
  return vim.tbl_map(function(p) return p.text end, neo.blocks(api.nvim_get_current_buf()))
end

function H.row(n) -- the buffer line (1-based) of paragraph n
  return neo.blocks(api.nvim_get_current_buf())[n].first
end

function H.set(n, text) -- replace paragraph n (one line) by text (which may hold several lines)
  local r = H.row(n)
  api.nvim_buf_set_lines(0, r - 1, r, false, vim.split(text, "\n"))
end

function H.replace(n, from, to) -- replace text inside paragraph n, as typing would (marks stay)
  local para = neo.blocks(api.nvim_get_current_buf())[n]
  local r, s
  for row = para.first, para.last do -- (on whichever of its lines it is)
    s = H.lines()[row]:find(from, 1, true)
    if s then
      r = row
      break
    end
  end
  assert(s, "not found: " .. from)
  api.nvim_buf_set_text(0, r - 1, s - 1, r - 1, s - 1 + #from, vim.split(to, "\n"))
  require("neoffice.changes").after_edit(api.nvim_get_current_buf()) -- (what TextChanged does)
end

function H.cursor(n, text) -- put the cursor on `text` in paragraph n
  local r = H.row(n)
  local s = assert(H.lines()[r]:find(text, 1, true), "not found: " .. text)
  api.nvim_win_set_cursor(0, { r, s - 1 })
end

function H.save()
  vim.cmd("silent write")
end

function H.reopen()
  H.save()
  return H.open(H.doc().path)
end

function H.bytes(path)
  local f = assert(io.open(path, "rb"))
  local data = f:read("*a")
  f:close()
  return data
end

function H.entries(path) -- name -> content of every file in the package
  local out = {}
  for name in pairs(zip.list(path)) do
    out[name] = zip.read(path, name) or ""
  end
  return out
end

function H.changed(before, after) -- names of the files that differ (sorted)
  local names = {}
  for name in pairs(vim.tbl_extend("force", {}, before, after)) do
    if before[name] ~= after[name] then
      names[#names + 1] = name
    end
  end
  table.sort(names)
  return names
end

-- The XML of each paragraph of a document, in order.
function H.paragraph_xml(path)
  local ext = path:match("%.(%w+)$")
  local root = xml.parse(zip.read(path, H.MAIN[ext]))
  local tags = ext == "docx" and { ["w:p"] = true } or { ["text:p"] = true, ["text:h"] = true }
  return vim.tbl_map(xml.serialize, xml.collect(root, tags, { ["text:tracked-changes"] = true }))
end

function H.differing(a, b) -- how many positions of two lists differ (and the difference in length)
  local n = math.abs(#a - #b)
  for k = 1, math.min(#a, #b) do
    n = n + (a[k] ~= b[k] and 1 or 0)
  end
  return n
end

function H.comments() -- the document's comments, oldest first
  local list = vim.tbl_values(H.doc().comment)
  table.sort(list, function(a, b) return a.date < b.date end)
  return list
end

function H.changes(author) -- the tracked changes (of one author)
  return vim.tbl_filter(function(ch) return not author or ch.author == author end, vim.tbl_values(H.doc().change))
end

-- Whether LibreOffice opens a file and finds this text in it (converted to text, or to `format`:
-- "html" has the footnotes too).
function H.libreoffice_reads(path, text, format)
  local dir = vim.fn.tempname()
  vim.system({ "soffice", "-env:UserInstallation=file://" .. dir .. "/profile", "--headless", "--convert-to", format or "txt:Text", "--outdir", dir, path }):wait()
  local out = io.open(dir .. "/" .. vim.fn.fnamemodify(path, ":t:r") .. "." .. (format or "txt"))
  local content = out and out:read("*a") or ""
  if out then
    out:close()
  end
  content = content:gsub("%s+", " ") -- (html wraps lines)
  return content:find(text, 1, true) ~= nil, content
end

H.has_libreoffice = vim.fn.executable("soffice") == 1
H.has_pandoc = vim.fn.executable("pandoc") == 1 -- (equations: LaTeX <-> Word)

return H
