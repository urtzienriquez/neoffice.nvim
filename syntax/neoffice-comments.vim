" The comments panel (:Neoffice comments), as fugitive's status window: what it is about, its help,
" then its sections, Open and Resolved, each comment its author and date, its text and its replies
" indented under it.
if exists("b:current_syntax")
  finish
endif

syn sync fromstart

syn match neofficeCommentsHeader /^[A-Z][a-z][^:]*:/
syn match neofficeCommentsHelpHeader /^Help:/ nextgroup=neofficeCommentsHelpTag skipwhite
syn match neofficeCommentsHelpTag /\S\+/ contained

syn match neofficeCommentsHeading /^\%(Open\|Resolved\)\ze (\d\+)$/ nextgroup=neofficeCommentsCount skipwhite
syn match neofficeCommentsCount /(\d\+)/hs=s+1,he=e-1 contained

" (a comment's author and date, at the start of its line; a reply's, two spaces in)
syn match neofficeCommentsAuthor /^ \{0,2}\S.\{-}\ze, \d\{4}-\d\d-\d\d$/ nextgroup=neofficeCommentsDate
syn match neofficeCommentsDate /, \d\{4}-\d\d-\d\d$/ contained

" (the resolved ones: quieter, but for their section's heading)
syn region neofficeCommentsResolved start=/^\%(Resolved (\d\+)$\)\@=/ end=/\%$/ contains=neofficeCommentsHeading

hi def link neofficeCommentsHelpHeader neofficeCommentsHeader
hi def link neofficeCommentsHeader Label
hi def link neofficeCommentsHelpTag Tag
hi def link neofficeCommentsHeading PreProc
hi def link neofficeCommentsCount Number
hi def link neofficeCommentsAuthor Identifier
hi def link neofficeCommentsDate Comment
hi def link neofficeCommentsResolved Comment

let b:current_syntax = "neoffice-comments"
