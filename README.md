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
- The document's paragraph styles at its top (`Normal: {font = Times New Roman, size = 11pt}`):
  edit a line to change the style, and every paragraph of it follows.
- Page settings at its top (`page: {size = A4, margins = 2.5cm}`), and sections where a line
  says what changes (`\newpage {orientation = landscape}`): a landscape page for a wide figure.
- Paragraph formats set for many paragraphs at once, as in LaTeX: a line `{align = justify, first =
  0.5cm}` sets what the paragraphs after it have of their own, against their style, until the next;
  one that differs says how (`<|> A centred one.`, `{before = 12pt} More space.`). A few words in
  another font or colour: `[words]{font = Arial}`, `[words]{color = red}` (and `{color = darkblue}`
  for paragraphs, `color = ...` for a style). A code action moves a format into its style.
- Paragraphs: `# ` headings (numbered on demand), `% ` title, `> ` quote, `- ` and numbered lists
  written as their numbering's pattern (`1.` for every item, `1.1.` or `1.a.` for sub-items: the
  document counts),
  and alignment `<|> centered`, `->> right`, `<=> justified`.
- Line breaks (`\` at the end of a line) and page breaks (`\newpage`).
- Tables as pipe tables (`| a | b |`): edit cells, add or remove rows and columns, new tables.
  Tables typeset themselves, as in LaTeX: booktabs rules, columns as wide as their content,
  compact rows, numbers aligned right. A line above a table changes its look
  (`{width = full, height = tall}`); colons in the line of dashes align columns (`|:---:|---:|`).
- Cropped images show and change their crop: `{width = 8cm, crop-top = 39%, crop-right = 11%}`.
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

- `:Neoffice comments` — the comments panel (`comments text` shows them at the end of their lines); `comment`, `reply`, `edit`, `resolve`, `delete`
- `:Neoffice track [on|off]` — record your edits as tracked changes
- `:Neoffice accept|reject [all [author]]` — the change at the cursor, the selection, or all
- `:Neoffice next|prev` — jump between changes; `:Neoffice changes [author]` lists them
- `:Neoffice numbers [on|off]` — number the headings
- `:Neoffice pdf` — the document as a PDF next to it, opened; `pdf auto [on|off]` writes it again after each save
  (one LibreOffice kept running in the background: a fraction of a second each; a viewer that reloads files,
  like zathura, shows each save)

One key is set in documents: `g?` opens the help. The comments panel has its own, two keys a family as in
fugitive (`aa` reply, `ee` edit, `rr` resolve, `dd` delete, `<CR>` go to it, `gq` close, `g?` its help).
Every key can be changed or dropped (`keymaps` in `setup()`); the rest are yours to set: see `:help
neoffice-keymaps`.

Writing in LaTeX or Rnoweb? [manuscript.nvim](https://github.com/urtzienriquez/manuscript.nvim)
exports your manuscript to Word, and brings your collaborators' revisions back into your sources.

## Requirements

- Neovim 0.10+ (to show images: 0.12+, in a terminal like Kitty, Ghostty or WezTerm)
- zlib (on every Linux and macOS system: documents are read and written in Neovim itself); without
  it, the `zip` and `unzip` commands
- Optional: [pandoc](https://pandoc.org) to edit equations (without it they are kept as they
  are, shown as `￼`), LibreOffice for `.doc` files and PDFs (fastest with its Python bridge,
  `python3-uno`); to show images other than PNG (made in the
  background, so opening never waits): ImageMagick (JPEG, GIF, TIFF, WebP...), `pdftoppm` (PDF),
  `rsvg-convert` (SVG), LibreOffice (EMF, WMF, SVM)

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
  styles = true, -- the document's paragraph styles at its top
  page = true, -- its page settings at its top, and its sections
  show_tabs = true, -- show tabs as a faint →, non-breaking spaces as a faint ·
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
