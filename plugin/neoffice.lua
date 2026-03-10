if vim.g.loaded_neoffice == 1 then
  return
end
vim.g.loaded_neoffice = 1

vim.api.nvim_create_user_command("DocOpen", function(args)
  require("neoffice").open(args.args ~= "" and args.args or nil)
end, { nargs = "?", complete = "file", desc = "Open .docx/.odt in Neovim" })

vim.api.nvim_create_user_command("DocSave", function()
  require("neoffice").save()
end, { desc = "Save proxy buffer back to original document" })

vim.api.nvim_create_user_command("DocChanges", function()
  require("neoffice").show_changes()
end, { desc = "Show floating track-changes summary" })

vim.api.nvim_create_user_command("DocAccept", function(args)
  require("neoffice").accept(args.args)
end, { nargs = "?", desc = "Accept change at cursor (or 'all')" })

vim.api.nvim_create_user_command("DocReject", function(args)
  require("neoffice").reject(args.args)
end, { nargs = "?", desc = "Reject change at cursor (or 'all')" })

vim.api.nvim_create_user_command("DocComments", function()
  require("neoffice").toggle_comments()
end, { desc = "Toggle comments panel" })

vim.api.nvim_create_user_command("DocAddComment", function()
  require("neoffice").add_comment()
end, { desc = "Add comment at cursor line" })

vim.api.nvim_create_user_command("DocDebug", function(args)
  require("neoffice.debug").run(args.args ~= "" and args.args or nil)
end, { nargs = "?", complete = "file", desc = "Dump neoffice diagnostic report" })

vim.api.nvim_create_user_command("DocDiagnoseAnnotations", function(args)
  require("neoffice.diagnose_annotations").show_saved_annotations(args.args ~= "" and args.args or nil)
end, { nargs = "?", complete = "file", desc = "Show annotation structure from ODT file" })
