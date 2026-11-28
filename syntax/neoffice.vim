" Tree-sitter (markdown) highlights the other marks (**bold**, ~~strikethrough~~, $equations$...);
" this adds ++underline++, ==highlight== (and ==highlight=={green}), superscript^{2},
" subscript_{2}, ^[footnotes], ⟦fields⟧, line and page breaks and vertical space (\vspace). (The alignment marks, highlight
" colours, table options and caption numbers are drawn by init.lua: M.decorate.)
" (The marks are hidden only with 'conceallevel' set, i.e. setup({ conceal = true }).)
syntax region neofficeUnderline matchgroup=neofficeScriptMark start="\\\@<!++" end="\\\@<!++" concealends keepend oneline
syntax region neofficeHighlight matchgroup=neofficeScriptMark start="\\\@<!==" end="\\\@<!==\({\s*\(color\s*=\s*\)\?[#[:alnum:]]\+\s*}\)\?" concealends keepend oneline
syntax region neofficeSuperscript matchgroup=neofficeScriptMark start="\\\@<![\^]{" end="\\\@<!}" concealends keepend
syntax region neofficeSubscript matchgroup=neofficeScriptMark start="\\\@<!_{" end="\\\@<!}" concealends keepend
syntax region neofficeFootnote matchgroup=neofficeFootnoteMark start="\\\@<![\^]\[" end="\\\@<!]" keepend oneline
syntax match neofficeField "\\\@<!⟦.\{-}\\\@<!⟧"
syntax match neofficeBreak "\\\@<!\\newpage"
syntax match neofficeBreak "\\\@<!\\$"
syntax match neofficeBreak "^\\vspace\s*\({[^{}]*}\)\?\s*$"
syntax match neofficeBreak "^\\\(tableofcontents\|listoffigures\|listoftables\)\s*$"
highlight default link neofficeUnderline Underlined
highlight default link neofficeHighlight Search
highlight default link neofficeSuperscript Special
highlight default link neofficeSubscript Special
highlight default link neofficeFootnote Comment
highlight default link neofficeFootnoteMark NeofficeMark
highlight default link neofficeField Special
highlight default link neofficeBreak NeofficeMark
highlight default link neofficeScriptMark NeofficeMark
" (neofficeScriptMark: the ++ ^{ _{ } and == around underlines, superscripts, subscripts and
" highlights)
