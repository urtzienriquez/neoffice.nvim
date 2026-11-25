-- Only autocommands and the :Neoffice command are defined here; no keymaps (see the README for some).

if vim.g.loaded_neoffice then
  return
end
vim.g.loaded_neoffice = true

local api = vim.api
local group = api.nvim_create_augroup("neoffice", {})
local pattern = { "*.docx", "*.odt", "*.doc" }

-- Documents (filetype "neoffice") are highlighted with the markdown parser, also when tree-sitter
-- is started for them without naming a language.
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

local function highlights()
  for name, hl in pairs({
    NeofficeTitle = { link = "Title" },
    NeofficeNumber = { link = "Number" },
    NeofficeAlign = { link = "@punctuation.special" },
    NeofficeChangeInfo = { link = "Comment" },
    NeofficeFormat = { underdashed = true, sp = "#61afef" },
    NeofficeComment = { underdotted = true, sp = "#e5c07b" },
    NeofficeCommentResolved = {},
    NeofficeCommentSign = { link = "DiagnosticWarn" },
    NeofficeCommentVirt = { link = "Comment" },
    NeofficeCommentAuthor = { link = "Title" },
    NeofficeCommentCurrent = { link = "Search" }, -- (the text of the comment the panel is on)
  }) do
    hl.default = true
    api.nvim_set_hl(0, name, hl)
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
  -- the document's edits into its LaTeX/Rnoweb source (files): a diff view to review them
  merge = function(arg) require("neoffice.merge").run(vim.tbl_map(vim.fn.expand, vim.split(arg, "%s+", { trimempty = true }))) end,
}
local ARGUMENTS = { accept = { "all" }, reject = { "all" }, track = { "on", "off" }, numbers = { "on", "off" }, comments = { "text" } }
local NAMES = vim.tbl_keys(SUBCOMMANDS)
table.sort(NAMES)

api.nvim_create_user_command("Neoffice", function(o)
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
    if words[2] == "comments" and words[3] == "text" then
      list = { "on", "off" }
    end
    if words[2] == "merge" then -- (source files)
      return vim.fn.getcompletion(lead, "file")
    end
    if by_author then -- the authors of the changes (names may have spaces: the rest of the line)
      local typed = cmdline:match("^%s*%S+%s+%S+%s+all%s+(.*)$") or cmdline:match("^%s*%S+%s+changes%s+(.*)$") or lead
      return vim.tbl_filter(function(name) return vim.startswith(name, typed) end, changes().authors())
    end
    return vim.tbl_filter(function(item) return vim.startswith(item, lead) end, list)
  end,
  desc = "neoffice: comments [text [on|off]], comment, reply, resolve, delete, changes [author], accept [all [author]], reject [all [author]], next, prev, track [on|off], numbers [on|off], merge {files}",
})
