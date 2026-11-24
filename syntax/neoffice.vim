" Tree-sitter (markdown) highlights the other marks (**bold**, ~~strikethrough~~, $equations$...);
" this adds <u>underline</u>, ==highlight==, superscript^{2}, subscript_{2}, ^[footnotes],
" ⟦fields⟧ and line and page breaks. (The alignment marks are highlighted by init.lua: M.decorate.)
" (The marks are hidden only with 'conceallevel' set, i.e. setup({ conceal = true }).)
syntax region neofficeUnderline matchgroup=neofficeTag start="\\\@<!<u>" end="</u>" concealends keepend
syntax region neofficeHighlight matchgroup=neofficeScriptMark start="\\\@<!==" end="\\\@<!==" concealends keepend oneline
syntax region neofficeSuperscript matchgroup=neofficeScriptMark start="\\\@<![\^]{" end="\\\@<!}" concealends keepend
syntax region neofficeSubscript matchgroup=neofficeScriptMark start="\\\@<!_{" end="\\\@<!}" concealends keepend
syntax region neofficeFootnote matchgroup=neofficeFootnoteMark start="\\\@<![\^]\[" end="\\\@<!]" keepend oneline
syntax match neofficeField "\\\@<!⟦.\{-}\\\@<!⟧"
syntax match neofficeBreak "\\\@<!\\newpage"
syntax match neofficeBreak "\\\@<!\\$"
highlight default link neofficeUnderline Underlined
highlight default link neofficeHighlight Search
highlight default link neofficeSuperscript Special
highlight default link neofficeSubscript Special
highlight default link neofficeTag Comment
highlight default link neofficeFootnote Comment
highlight default link neofficeFootnoteMark Comment
highlight default link neofficeField Special
highlight default link neofficeBreak NonText
" (neofficeScriptMark, the ^{ _{ } and == around superscripts, subscripts and highlights, has no
" highlight of its own)
