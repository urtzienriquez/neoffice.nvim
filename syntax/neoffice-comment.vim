" A comment's buffer (a reply, a new comment, one edited), as git's commit message: its text, then
" what it answers and what to do, in lines starting with "#" (left out when written).
if exists("b:current_syntax")
  finish
endif

syn match neofficeCommentHelp /^#.*/

hi def link neofficeCommentHelp Comment

let b:current_syntax = "neoffice-comment"
