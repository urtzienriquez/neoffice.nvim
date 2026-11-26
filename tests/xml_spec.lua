local xml = require("neoffice.xml")

local function roundtrip(s)
  return xml.serialize(xml.parse(s))
end

describe("xml", function()
  it("round-trips entities without double escaping", function()
    assert.equals("<a>Tom &amp; Jerry &lt;3</a>", roundtrip("<a>Tom &amp; Jerry &lt;3</a>"))
    assert.equals("Tom & Jerry <3", xml.text(xml.parse("<a>Tom &amp; Jerry &lt;3</a>")))
  end)

  it("keeps whitespace-only text", function()
    assert.equals("<r><t> </t></r>", roundtrip("<r><t> </t></r>"))
    assert.equals("<a>\n  <b/>\n</a>", roundtrip("<a>\n  <b/>\n</a>"))
  end)

  it("keeps attribute order and decodes values", function()
    local doc = xml.parse([[<p z="1" a='x &quot;y&quot;' m="&lt;"/>]])
    assert.equals('x "y"', xml.attr(doc.children[1], "a"))
    assert.equals('<p z="1" a="x &quot;y&quot;" m="&lt;"/>', xml.serialize(doc))
  end)

  it("reads attribute values containing > and quotes", function()
    local doc = xml.parse([[<a href="x>y" t='say "hi"'>z</a>]])
    assert.equals("x>y", xml.attr(doc.children[1], "href"))
    assert.equals('say "hi"', xml.attr(doc.children[1], "t"))
    assert.equals("z", xml.text(doc))
  end)

  it("keeps declarations, comments and CDATA verbatim", function()
    local s = '<?xml version="1.0"?><a><!-- note > here --><![CDATA[<raw>]]></a>'
    assert.equals(s, roundtrip(s))
  end)

  it("decodes numeric entities", function()
    assert.equals("é€", xml.text(xml.parse("<a>&#233;&#x20AC;</a>")))
  end)

  it("keeps unknown entities and unicode as they are", function()
    assert.equals("<a>&amp;nbsp; ñ 日本 😀</a>", roundtrip("<a>&nbsp; ñ 日本 😀</a>"))
  end)

  it("handles namespaces, self-closing tags with spaces, and deep nesting", function()
    local s = '<w:p xmlns:w="u"><w:r><w:t xml:space="preserve"> a </w:t></w:r><w:br /></w:p>'
    assert.equals('<w:p xmlns:w="u"><w:r><w:t xml:space="preserve"> a </w:t></w:r><w:br/></w:p>', roundtrip(s))
    local deep = ("<a>"):rep(200) .. "x" .. ("</a>"):rep(200)
    assert.equals(deep, roundtrip(deep))
  end)

  it("builds, finds, copies and removes elements", function()
    local doc = xml.parse("<r><a/><b><c n='1'/></b></r>")
    local c = xml.find(doc, "c")
    assert.equals("1", xml.attr(c, "n"))
    local copy = xml.copy(xml.find(doc, "b"))
    xml.set_attr(xml.find(copy, "c"), "n", "2")
    assert.equals("1", xml.attr(c, "n"))
    xml.remove(xml.find(doc, "a"))
    xml.append(xml.find(doc, "r"), xml.elem("d", { { "k", "v" } }, { { text = "t" } }), 1)
    assert.equals('<r><d k="v">t</d><b><c n="1"/></b></r>', xml.serialize(doc))
    assert.equals(1, #xml.collect(doc, { c = true }))
  end)

  it("tells a part Word and LibreOffice can read from one they can't", function()
    assert.is_true(xml.check('<?xml version="1.0"?><w:document xmlns:w="u"><w:body w:x="1"/></w:document>'))
    assert.is_true(xml.check('<a xmlns:p="u"><p:b xml:space="preserve"/></a>'))
    assert.is_nil(xml.check("<a/><b/>")) -- (two roots)
    assert.is_nil(xml.check('<a xmlns:p="u"><q:b/></a>')) -- (a prefix not declared)
    assert.is_nil(xml.check('<a><b p:x="1"/></a>'))
    assert.is_nil(xml.check('<a xmlns:p="u"/><p:b/>'))
  end)
end)
