# neoffice.nvim

Edit Word (`.docx`, `.doc`) and LibreOffice (`.odt`) documents in Neovim as plain text with
Markdown-style marks: `nvim report.docx`, edit, `:w`. Only what you change is written: everything
else in the file stays exactly as it was.

**This is a project under development. Please, feel free to open issues or pull requests.**

## Features

- Paragraphs separated by blank lines, as in Markdown; new documents with `nvim new.docx`.
- Formatting: `**bold**`, `*italic*`, `++underline++`, `~~strike~~`, `==highlight==` (other
  colours: `==highlight=={green}`), `x^{2}`, `W_{food}`, `[links](https://…)`.
- Captions: `Table: …` and `Figure: …`, real captions numbered by the document (Word's and
  LibreOffice's own numbering, so cross-references work).
- Paragraphs: `# ` headings (numbered on demand), `% ` title, `> ` quote, `- ` and `1. ` lists,
  and alignment `<|> centered`, `->> right`, `<=> justified`.
- Line breaks (`\` at the end of a line) and page breaks (`\newpage`).
- Tables as pipe tables (`| a | b |`): edit cells, add or remove rows and columns, new tables.
  Tables typeset themselves, as in LaTeX: booktabs rules, columns as wide as their content,
  compact rows, numbers aligned right. A line above a table changes its look
  (`{width = full, height = tall}`); colons in the line of dashes align columns (`|:---:|---:|`).
- Floating images and boxes, placed by meaning, not coordinates: `{width = 6cm, float = right}`;
  a box (a text box, a frame) as `::: {float = right} … :::`, its paragraphs editable (a figure with
  its caption, a sidebar). Charts, shapes and other drawings show by name (`⟦chart⟧`) and are kept.
- Footnotes (`^[note]`), images (`![alt](path){width = 50%}`, shown with `vim.ui.img`), equations
  as LaTeX (`$…$`, `$$…$$`, `.docx`), and fields and citations, shown as `⟦…⟧` and kept.
- Comments: add, reply, resolve and delete, in a side panel.
- Tracked changes: shown by author, accepted or rejected (at the cursor, in a selection, all, or
  someone's), and your own edits tracked. Undo works.
- A built-in language server (`neoffice_ls`): headings as symbols (`gO`, pickers, outlines),
  formatting (tables aligned, with conform.nvim's `lsp_format = "fallback"`), completion of
  table options and highlight colours, sections as folds,
  hover on comments and changes.

One command, `:Neoffice` (alone, it lets you pick):

- `:Neoffice comments` — the comments panel (`comments text` shows them at the end of their lines); `comment`, `reply`, `resolve`, `delete`
- `:Neoffice track [on|off]` — record your edits as tracked changes
- `:Neoffice accept|reject [all [author]]` — the change at the cursor, the selection, or all
- `:Neoffice next|prev` — jump between changes; `:Neoffice changes [author]` lists them
- `:Neoffice numbers [on|off]` — number the headings

No keymaps are set. See `:help neoffice`.

Writing in LaTeX or Rnoweb? [manuscript.nvim](https://github.com/urtzienriquez/manuscript.nvim)
exports your manuscript to Word, and brings your collaborators' revisions back into your sources.

## Requirements

- Neovim 0.10+ (to show images: 0.12+, in a terminal like Kitty, Ghostty or WezTerm)
- `zip` and `unzip` (as for Neovim's own zip plugin)
- Optional: [pandoc](https://pandoc.org) to edit equations (without it they are kept as they
  are, shown as `￼`), LibreOffice for `.doc` files, ImageMagick to show JPEG and GIF images

## Installation

```lua
vim.pack.add({ "https://github.com/urtzienriquez/neoffice.nvim" })
```

No `setup()` needed. Optional:

```lua
require("neoffice").setup({
  author = "Your Name", -- on your comments and changes (default: $USER)
  track_changes = nil, -- true/false: always/never track your edits (default: as the document says)
  conceal = false, -- hide the marks
  images = true, -- show images
  comment_text = false, -- each comment's text at the end of its line (or :Neoffice comments text)
  captions = { table = "Table", figure = { "Figure", "Fig." }, separator = ". " }, -- caption words (first written), after the number
  tables = { rules = "booktabs", width = "auto", height = "compact", position = "center" }, -- how tables look
})
vim.api.nvim_create_autocmd("FileType", {
  pattern = "neoffice",
  callback = function(ev)
    vim.keymap.set("n", "<leader>wc", "<Cmd>Neoffice comments<CR>", { buffer = ev.buf, desc = "Comments" })
    vim.keymap.set({ "n", "x" }, "<leader>wy", "<Cmd>Neoffice accept<CR>", { buffer = ev.buf, desc = "Accept change" })
  end,
})
```

## License

GNU General Public License v3.0 — see [LICENSE](LICENSE).
