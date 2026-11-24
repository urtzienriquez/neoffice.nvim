#!/bin/sh
# Regenerates the test documents. Needs pandoc, LibreOffice (soffice), ImageMagick, python3 and zip.
#   rich.docx    pandoc: heading, formatting, link, comment, tracked changes, image, list, table
#   rich.odt     LibreOffice's conversion of rich.docx
#   office.docx  LibreOffice's .docx of rich.odt with a reply added: comment threads as Word has them
#   empty.docx / empty.odt   documents without any paragraph
#   math.docx / math.odt     equations, superscripts and subscripts
#   layout.docx / layout.odt alignment, quote, line and page breaks, strikethrough, highlight,
#                            footnotes and fields
set -e
cd "$(dirname "$0")"
tmp=$(mktemp -d)
soffice="soffice -env:UserInstallation=file://$tmp/profile --headless"
magick -size 40x20 xc:tomato "$tmp/dot.png" # images (also used by the tests: 40x20 pixels)
magick -size 30x30 xc:steelblue "$tmp/square.jpg"
cat > "$tmp/rich.md" <<'MD'
# Introduction

This is **bold**, *italic*, [underlined]{.underline} and a [link](https://example.org). Tom & Jerry <3, 2*3.

[A comment]{.comment-start id="0" author="Alice" date="2024-01-01T10:00:00Z"}Commented text[]{.comment-end id="0"} and more.

Text with [inserted words]{.insertion author="Bob" date="2024-01-02T10:00:00Z"} and [deleted words]{.deletion author="Bob" date="2024-01-02T10:00:00Z"} inside.

An image ![](dot.png) here.

- first item
- second item

| A | B |
|---|---|
| cell one | cell two |

Last paragraph.
MD
: > "$tmp/empty.md"
cat > "$tmp/math.md" <<'MD'
Area in cm^2^ and a rate in mL h^-1^ of W~food~.

Inline $x^2 + y_1$ here.

$$B = W_{food} + W_{met} - E$$

Prices: $5 and $10.
MD
cat > "$tmp/layout.md" <<'MD'
Centered text.

Right text.

Justified text.

> A quote.

First line\
second line.

Before the page break.

After the page break.

Some ~~struck~~ and [lit]{.mark} words.

A claim[^1] and another[^2].

See CITATION for that.

[^1]: The first note.

[^2]: The second note.
MD
(
  cd "$tmp"
  pandoc rich.md -o rich.docx
  pandoc layout.md -o layout.docx
  # alignment, a page break and a field (a citation, as Zotero writes it)
  mkdir layout && cd layout && unzip -q ../layout.docx
  python3 - <<'PY'
import re
s = open("word/document.xml").read()
paras = re.findall(r"<w:p>.*?</w:p>", s)
def change(text, f):
    global s
    p = next(p for p in paras if text in p)
    s = s.replace(p, f(p), 1)
for text, jc in (("Centered text", "center"), ("Right text", "right"), ("Justified text", "both")):
    change(text, lambda p: p.replace("</w:pPr>", '<w:jc w:val="%s"/></w:pPr>' % jc, 1))
change("After the page break", lambda p: p.replace("</w:pPr>", '</w:pPr><w:r><w:br w:type="page"/></w:r>', 1))
field = ('</w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r>'
         '<w:r><w:instrText xml:space="preserve"> ADDIN ZOTERO_ITEM {} </w:instrText></w:r>'
         '<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>(Smith, 2020)</w:t></w:r>'
         '<w:r><w:fldChar w:fldCharType="end"/></w:r><w:r><w:t xml:space="preserve">')
change("CITATION", lambda p: p.replace("CITATION", field))
open("word/document.xml", "w").write(s)
PY
  zip -q -X -r ../layout.docx . && cd ..
  pandoc empty.md -o empty.docx
  pandoc math.md -o math.docx
  $soffice --convert-to odt rich.docx empty.docx math.docx layout.docx > /dev/null
  # add a reply to the comment, then let LibreOffice write a .docx with it
  mkdir unpacked && cd unpacked && unzip -q ../rich.odt
  python3 - <<'PY'
import re
s = open("content.xml").read()
name = re.search(r'office:name="([^"]+)"', s).group(1)
reply = ('<office:annotation loext:parent-name="%s"><dc:creator>Carol</dc:creator>'
         '<dc:date>2024-01-03T10:00:00</dc:date><text:p>A reply</text:p></office:annotation>' % name)
s = s.replace('<office:annotation-end office:name="%s"/>' % name, '<office:annotation-end office:name="%s"/>' % name + reply)
open("content.xml", "w").write(s)
PY
  zip -q -X -0 ../reply.odt mimetype && zip -q -X -r ../reply.odt . -x mimetype
  cd .. && $soffice --convert-to docx reply.odt > /dev/null && mv reply.docx office.docx
)
cp "$tmp/rich.docx" "$tmp/rich.odt" "$tmp/office.docx" "$tmp/empty.docx" "$tmp/empty.odt" "$tmp/math.docx" "$tmp/math.odt" "$tmp/layout.docx" "$tmp/layout.odt" "$tmp/dot.png" "$tmp/square.jpg" .
rm -rf "$tmp"
