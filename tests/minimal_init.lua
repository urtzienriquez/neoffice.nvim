-- nvim --headless -u tests/minimal_init.lua -c "PlenaryBustedDirectory tests {minimal_init='tests/minimal_init.lua'}"
vim.opt.rtp:prepend(vim.fn.getcwd())
for _, dir in ipairs({
  vim.env.PLENARY_DIR or "",
  vim.fn.stdpath("data") .. "/lazy/plenary.nvim",
  vim.fn.stdpath("data") .. "/site/pack/test/start/plenary.nvim",
}) do
  if vim.fn.isdirectory(dir) == 1 then
    vim.opt.rtp:append(dir)
  end
end
vim.cmd("runtime plugin/plenary.vim")
vim.cmd("runtime plugin/neoffice.lua")
