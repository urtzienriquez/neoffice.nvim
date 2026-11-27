-- Only autocommands and the :Neoffice command are defined here; the keys are buffer-local (g? in
-- documents, the comments panel's: see neoffice.config.keymaps).

if vim.g.loaded_neoffice then
  return
end
vim.g.loaded_neoffice = true

local api = vim.api
local group = api.nvim_create_augroup("neoffice", {})
local pattern = { "*.docx", "*.odt", "*.doc" }

-- Documents (filetype "neoffice") are highlighted with the markdown parser, also when tree-sitter
-- is started for them without naming a language (see neoffice.languages: as "neoffice", markdown's
-- own queries but for bold, italics and strikethrough, which neoffice shows as it reads them).
vim.treesitter.language.register("markdown", "neoffice")

-- Opening and saving documents go through :edit and :write.
api.nvim_create_autocmd("BufReadCmd", {
  group = group,
  pattern = pattern,
  callback = function(ev)
    require("neoffice").open(ev.buf, ev.match)
  end,
})
api.nvim_create_autocmd("BufWriteCmd", {
  group = group,
  pattern = pattern,
  callback = function(ev)
    require("neoffice").save(ev.buf, ev.match)
  end,
})

-- Neovim's zip browser claims these files too; keep it away from them, whichever loads first.
local function unclaim()
  for _, zip_group in ipairs({ "zip", "nvim.zip" }) do
    pcall(api.nvim_clear_autocmds, { group = zip_group, event = "BufReadCmd", pattern = pattern })
  end
end
unclaim()
api.nvim_create_autocmd("SourcePost", { group = group, pattern = { "*/plugin/zip.lua", "*/plugin/zipPlugin.vim" }, callback = unclaim })

-- The colours say what is what: text in the text's colour; formatting in the text (marks,
-- attributes, alignment, tabs, backslashes) in one gray, readable but quieter than the text
-- (NeofficeMark: half way between the text's colour and the background's); what says what a
-- paragraph is (a heading's #, a title's %, a list's -, a quote's >, a caption's word, a table's |,
-- a box's :::) in that element's colour; what the document makes (numbers) or keeps as it is
-- (fields) as the colourscheme shows special things; comments and changes in their colours.
local function mix(a, b, t) -- (colours as numbers: a, t of the way to b)
  local out = 0
  for shift = 16, 0, -8 do
    local x, y = bit.band(bit.rshift(a, shift), 255), bit.band(bit.rshift(b, shift), 255)
    out = out + bit.lshift(math.floor(x + (y - x) * t + 0.5), shift)
  end
  return out
end
local function mark_colour()
  local normal = api.nvim_get_hl(0, { name = "Normal", link = false })
  if not (normal.fg and normal.bg) then -- (the terminal's colours: not known)
    return { link = "Comment" }
  end
  return { fg = mix(normal.fg, normal.bg, 0.45) }
end

local function highlights()
  for name, hl in pairs({
    NeofficeMark = mark_colour(), -- formatting, as opposed to text
    NeofficeTitle = { link = "Title" },
    NeofficeNumber = { link = "Number" },
    NeofficeAlign = { link = "NeofficeMark" },
    NeofficeTableOptions = { link = "NeofficeMark" },
    NeofficeBox = { link = "@punctuation.special" }, -- (a box's :::, as a table's |)
    NeofficeCaption = { link = "@label" },
    NeofficeBold = { link = "@markup.strong" }, -- (bold, italics, strikethrough: as neoffice reads them)
    NeofficeItalic = { link = "@markup.italic" },
    NeofficeStrike = { link = "@markup.strikethrough" },
    -- (a link or an image, all of it: "[text](url)", "![alt](path)"; as a colour scheme shows links)
    NeofficeLink = { link = next(api.nvim_get_hl(0, { name = "@markup.link.label.markdown_inline" })) and "@markup.link.label.markdown_inline" or "@markup.link.label" },
    NeofficeStyleName = { link = "@label" }, -- (a style's name, in the styles at the top)
    NeofficeSection = { link = "NeofficeBox" }, -- (where a section starts: "\\newpage {...}")
    NeofficeEscape = { link = "NeofficeMark" }, -- (a backslash that makes the next character plain text)
    NeofficeChangeInfo = { link = "Comment" },
    NeofficeFormat = { underdashed = true, sp = "#61afef" },
    NeofficeComment = { underdotted = true, sp = "#e5c07b" },
    NeofficeCommentResolved = {},
    NeofficeCommentSign = { link = "DiagnosticWarn" },
    NeofficeCommentVirt = { link = "Comment" },
    NeofficeCommentCurrent = { link = "Search" }, -- (the text of the comment the panel is on)
    -- (++underline++, ^{superscripts}, _{subscripts}, ^[footnotes], ⟦fields⟧, \newpage: drawn by
    -- neoffice, or by syntax/neoffice.vim when the marks are concealed)
    neofficeUnderline = { link = "Underlined" },
    neofficeHighlight = { link = "Search" },
    neofficeSuperscript = { link = "Special" },
    neofficeSubscript = { link = "Special" },
    neofficeFootnote = { link = "Comment" },
    neofficeFootnoteMark = { link = "NeofficeMark" },
    neofficeField = { link = "Special" },
    neofficeBreak = { link = "NeofficeMark" },
    neofficeScriptMark = { link = "NeofficeMark" },
  }) do
    hl.default = true
    api.nvim_set_hl(0, name, hl)
  end
  -- a colour scheme's colours for markdown (headings, links, code...), for documents too (their
  -- languages: "neoffice", "neoffice_inline", see neoffice.languages)
  for name in pairs(api.nvim_get_hl(0, {})) do
    local base, lang = name:match("^(@.+)%.(markdown[_inline]*)$")
    if base and (lang == "markdown" or lang == "markdown_inline") then
      api.nvim_set_hl(0, base .. (lang == "markdown" and ".neoffice" or ".neoffice_inline"), { link = name, default = true })
    end
  end
  -- tracked changes: a colour for each author (insertions underlined, deletions struck through)
  for n, color in ipairs({ "#98c379", "#61afef", "#c678dd", "#e5c07b", "#56b6c2", "#e06c75" }) do
    api.nvim_set_hl(0, "NeofficeInsert" .. n, { underline = true, fg = color, default = true })
    api.nvim_set_hl(0, "NeofficeDelete" .. n, { strikethrough = true, fg = color, default = true })
  end
end
highlights()
api.nvim_create_autocmd("ColorScheme", { group = group, callback = highlights })

-- :Neoffice {subcommand} [argument]; alone, it lists the subcommands to pick from.
local function comments()
  return require("neoffice.comments")
end
local function changes()
  return require("neoffice.changes")
end
local SUBCOMMANDS = {
  comments = function(arg) -- the comments panel; text [on|off]: their text at the end of the line
    if arg:match("^text") then
      return comments().text(arg:match("^text%s+(%S+)"))
    end
    comments().toggle_panel()
  end,
  comment = function(_, o) comments().add(o.range > 0) end, -- comment the selection / cursor
  reply = function() comments().reply() end, -- (to the comment under the cursor)
  edit = function() comments().edit() end, -- (the text of the comment under the cursor)
  resolve = function() comments().resolve() end,
  delete = function() comments().delete() end,
  changes = function(arg) changes().list(arg ~= "" and arg or nil) end, -- tracked changes (of an author)
  -- at the cursor, a selection, all, or all of an author: accept all Bob Smith
  accept = function(arg, o) changes().resolve("accept", arg:match("^all") ~= nil, o.range > 0, arg:match("^all%s+(.+)")) end,
  reject = function(arg, o) changes().resolve("reject", arg:match("^all") ~= nil, o.range > 0, arg:match("^all%s+(.+)")) end,
  next = function() changes().jump(1) end,
  prev = function() changes().jump(-1) end,
  track = function(arg) changes().track(arg ~= "" and arg or nil) end, -- record your edits: on/off
  numbers = function(arg) require("neoffice").numbers(arg ~= "" and arg or nil) end, -- number headings: on/off
  pdf = function(arg) require("neoffice.pdf").command(arg) end, -- export it as a PDF; auto [on|off]: after each save
  styles = function(arg) -- reset: the styles as a new document's
    if arg ~= "reset" then
      return vim.notify("neoffice: :Neoffice styles reset", vim.log.levels.WARN)
    end
    require("neoffice").reset_styles(api.nvim_get_current_buf())
  end,
}
local ARGUMENTS = { accept = { "all" }, reject = { "all" }, track = { "on", "off" }, numbers = { "on", "off" }, comments = { "text" }, pdf = { "auto" }, styles = { "reset" } }
local NAMES = vim.tbl_keys(SUBCOMMANDS)
table.sort(NAMES)

api.nvim_create_user_command("Neoffice", function(o)
  require("neoffice").finish_all() -- (what the documents are as opened: worked out before any change)
  local sub, arg = o.args:match("^%s*(%S*)%s*(.-)%s*$")
  if sub == "" then
    return vim.ui.select(NAMES, { prompt = "Neoffice" }, function(choice)
      if choice then
        SUBCOMMANDS[choice]("", o)
      end
    end)
  elseif not SUBCOMMANDS[sub] then
    return vim.notify("Unknown :Neoffice subcommand: " .. sub, vim.log.levels.ERROR)
  end
  SUBCOMMANDS[sub](arg, o)
end, {
  nargs = "*",
  range = true,
  complete = function(lead, cmdline)
    local words = vim.split(cmdline:sub(1, #cmdline - #lead), "%s+", { trimempty = true })
    local list = #words == 1 and NAMES or ARGUMENTS[words[2]] or {}
    local by_author = (words[2] == "accept" or words[2] == "reject") and words[3] == "all" or words[2] == "changes" and #words == 2
    if words[2] == "comments" and words[3] == "text" or words[2] == "pdf" and words[3] == "auto" then
      list = { "on", "off" }
    end
    if by_author then -- the authors of the changes (names may have spaces: the rest of the line)
      local typed = cmdline:match("^%s*%S+%s+%S+%s+all%s+(.*)$") or cmdline:match("^%s*%S+%s+changes%s+(.*)$") or lead
      return vim.tbl_filter(function(name) return vim.startswith(name, typed) end, changes().authors())
    end
    return vim.tbl_filter(function(item) return vim.startswith(item, lead) end, list)
  end,
  desc = "neoffice: comments [text [on|off]], comment, reply, edit, resolve, delete, changes [author], accept [all [author]], reject [all [author]], next, prev, track [on|off], numbers [on|off], pdf [auto [on|off]], styles reset",
})
