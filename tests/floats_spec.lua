-- Floating images and boxes: where they go, as attributes ({float = right}); boxes as ::: blocks.
local H = dofile("tests/helpers.lua")
local zip = require("neoffice.zip")
local floats = require("neoffice.floats")
local api = vim.api

local DOT = H.FIXTURES .. "dot.png"
local PAGE = { text = 16, left = 2.5, top = 2.5 }

describe("floats", function()
  it("reads where a float is as what it means, and makes the place again from it", function()
    local function means(geo) return floats.classify(geo, PAGE) end
    assert.same({ float = "right" }, means({ h = { rel = "column", offset = 10 }, v = { rel = "paragraph", offset = 0.05 }, wrap = "square", gap = 0.3, width = 6 }))
    assert.same({ float = "center" }, means({ h = { rel = "column", align = "center" }, v = { rel = "paragraph", offset = 0 }, wrap = "none", gap = 0.3, width = 12 }))
    assert.same({ float = "left", x = "3cm", wrap = "tight", gap = "0cm" }, means({ h = { rel = "column", offset = 3 }, v = { rel = "paragraph" }, wrap = "tight", gap = 0, width = 4 }))
    assert.same({ float = "left", place = "top" }, means({ h = { rel = "column", offset = 0 }, v = { rel = "margin", align = "top" }, wrap = "square", gap = 0.3, width = 4 }))
    assert.same({ float = "left", place = "page", x = "1cm", y = "2cm", wrap = "front" }, means({ h = { rel = "page", offset = 1 }, v = { rel = "page", offset = 2 }, wrap = "front", gap = 0.3, width = 4 }))
    for _, attrs in ipairs({ { float = "right" }, { float = "center", wrap = "tight" }, { float = "left", x = "3cm", y = "1cm" }, { float = "left", place = "bottom" } }) do
      local geo = floats.geometry(attrs, PAGE, 4)
      if geo.h.align then -- (as a format writes it: an offset)
        geo.h = { rel = geo.h.rel, offset = ({ left = 0, right = PAGE.text - 4, center = (PAGE.text - 4) / 2 })[geo.h.align] }
      end
      assert.same(attrs, floats.classify(geo, PAGE), vim.inspect(attrs))
    end
    assert.is_nil(floats.geometry({ width = "3cm" }, PAGE, 3)) -- (in the line)
  end)

  it("reads an image's attributes, names or values alone, and says which are wrong", function()
    assert.same({ width = "6cm", float = "right" }, floats.parse("width = 6cm, float = right"))
    assert.same({ width = "6cm", float = "right", wrap = "tight" }, floats.parse("6cm right tight"))
    assert.is_nil(floats.parse("float = up"))
    assert.is_nil(floats.parse("colour = red"))
    assert.equals("width = 6cm, float = center", floats.render({ width = "6cm", float = "center", wrap = "none", gap = "0.3cm", place = "here" }))
  end)
end)

for _, ext in ipairs({ "docx", "odt" }) do
  describe(ext, function()
    local function main() return zip.read(H.doc().path, H.MAIN[ext]) end

    it("floats an image, and puts it back in the line", function()
      H.open(H.copy("rich." .. ext))
      H.replace(5, "{width = 1.4cm}", "{width = 1.4cm, float = right, wrap = tight}")
      H.reopen()
      assert.truthy(H.paras()[5]:find("{width = 1.4cm, float = right, wrap = tight}", 1, true), H.paras()[5])
      assert.truthy(main():find(ext == "docx" and "<wp:anchor" or 'text:anchor%-type="paragraph"'))
      H.replace(5, ", float = right, wrap = tight", "")
      H.reopen() -- (where it went: at the start of its paragraph)
      assert.truthy(H.paras()[5]:match("^!%[%]%(.-%){width = 1.4cm}An image  here%.$"), H.paras()[5])
      assert.falsy(main():find(ext == "docx" and "<wp:anchor" or 'text:anchor%-type="paragraph"'))
    end)

    it("shows a floating image at the start of its paragraph", function()
      H.open(H.copy("rich." .. ext))
      H.replace(5, "{width = 1.4cm}", "{width = 1.4cm, float = left}")
      H.reopen()
      assert.truthy(H.paras()[5]:match("^!%[%]%(.-%){width = 1.4cm, float = left}An image  here%.$"), H.paras()[5])
    end)

    it("makes a box with paragraphs in it, edits it, and takes it away", function()
      H.open(H.copy("rich." .. ext))
      local r = H.row(12)
      api.nvim_buf_set_lines(0, r - 1, r - 1, false, {
        "::: {width = 6cm, float = right}", "![A dot](" .. DOT .. ")", "", "Figure: A dot in a box.", ":::", "",
      })
      H.reopen()
      local lines = H.lines()
      local open = vim.fn.index(lines, "::: {width = 6cm, float = right}") + 1
      assert.truthy(open > 0, table.concat(lines, "\n"))
      assert.truthy(lines[open + 1]:match("^!%[A dot%]%(.-%){width = 6cm}$"), lines[open + 1]) -- (as wide as its box)
      assert.same({ "", "Figure: A dot in a box.", ":::", "", "Last paragraph." }, vim.list_slice(lines, open + 2, open + 6))
      assert.truthy(main():find(ext == "docx" and "<w:txbxContent>" or "<draw:custom%-shape"))
      -- its caption is numbered, and its text and place can be edited
      api.nvim_buf_set_lines(0, open - 1, open, false, { "::: {width = 5cm, float = left}" })
      local k = vim.fn.index(H.paras(), "Figure: A dot in a box.") + 1
      H.replace(k, "A dot in a box.", "A dot, edited.")
      H.reopen()
      assert.truthy(vim.tbl_contains(H.lines(), "::: {width = 5cm, float = left}"))
      assert.truthy(vim.tbl_contains(H.paras(), "Figure: A dot, edited."))
      -- without its lines, it is gone (and its paragraphs with it)
      lines = H.lines()
      open = vim.fn.index(lines, "::: {width = 5cm, float = left}") + 1
      api.nvim_buf_set_lines(0, open - 1, open + 4, false, {})
      H.reopen()
      assert.falsy(vim.iter(H.lines()):any(function(l) return l:match("^:::") end))
      assert.falsy(main():find(ext == "docx" and "<w:txbxContent>" or "<draw:custom%-shape"))
      assert.equals("Last paragraph.", H.paras()[#H.paras()])
    end)

    it("doesn't write a box's attribute it doesn't know", function()
      local path = H.copy("rich." .. ext)
      H.open(path)
      local before = H.bytes(path)
      local r = H.row(12)
      api.nvim_buf_set_lines(0, r - 1, r - 1, false, { "::: {float = up}", "Text.", ":::", "" })
      H.save()
      assert.equals(before, H.bytes(path))
    end)
  end)
end

describe("objects neoffice doesn't edit", function()
  it("show by name, and are kept", function()
    local path = H.copy("rich.docx", function(name, text)
      if name ~= "word/document.xml" then
        return text
      end
      local shape = '<w:p><w:r><w:t xml:space="preserve">A drawing </w:t></w:r><w:r><w:pict><v:rect xmlns:v="urn:schemas-microsoft-com:vml" style="width:20pt;height:20pt"/></w:pict></w:r><w:r><w:t xml:space="preserve"> here.</w:t></w:r></w:p>'
      return (text:gsub("<w:body>", "<w:body>" .. shape, 1))
    end)
    H.open(path)
    assert.equals("A drawing ⟦drawing⟧ here.", H.paras()[1])
    H.replace(1, "A drawing", "Still a drawing")
    H.reopen()
    assert.equals("Still a drawing ⟦drawing⟧ here.", H.paras()[1])
    assert.truthy(zip.read(path, "word/document.xml"):find("<v:rect", 1, true))
  end)
end)

describe("the language server, on floats and boxes", function()
  local function request(method, params)
    local buf = api.nvim_get_current_buf()
    assert(vim.wait(2000, function() return #vim.lsp.get_clients({ bufnr = buf, name = "neoffice_ls" }) > 0 end))
    params = vim.tbl_extend("force", { textDocument = { uri = vim.uri_from_bufnr(buf) } }, params or {})
    local result
    for _, r in pairs(vim.lsp.buf_request_sync(buf, method, params, 2000) or {}) do
      result = r.result
    end
    return result
  end

  it("offers an image's attributes, and a box's", function()
    H.open(H.copy("rich.docx"))
    local r = H.row(12)
    api.nvim_buf_set_lines(0, r - 1, r - 1, false, { "![](x.png){width = 3cm, ", "", "::: {float = ", "" })
    local got = vim.tbl_map(function(i) return i.label end, request("textDocument/completion", { position = { line = r - 1, character = #"![](x.png){width = 3cm, " } }).items)
    assert.is_true(vim.tbl_contains(got, "float = right"))
    assert.is_false(vim.tbl_contains(got, "width = 6cm"))
    got = vim.tbl_map(function(i) return i.label end, request("textDocument/completion", { position = { line = r + 1, character = #"::: {float = " } }).items)
    table.sort(got)
    assert.same({ "center", "left", "right" }, got)
  end)

  it("boxes a floating figure with its caption", function()
    H.open(H.copy("rich.docx"))
    local r = H.row(12)
    api.nvim_buf_set_lines(0, r - 1, r - 1, false, { "![A dot](" .. DOT .. "){width = 4cm, float = right}", "", "Figure: A dot.", "" })
    local actions = request("textDocument/codeAction", {
      range = { start = { line = r - 1, character = 0 }, ["end"] = { line = r - 1, character = 0 } }, context = { diagnostics = {} },
    })
    local action = vim.iter(actions):find(function(a) return a.title:find("Box this figure", 1, true) end)
    vim.lsp.util.apply_workspace_edit(action.edit, "utf-8")
    assert.same({ "::: {width = 4cm, float = right}", "![A dot](" .. DOT .. "){width = 4cm}", "", "Figure: A dot.", ":::" }, vim.list_slice(H.lines(), r, r + 4))
  end)
end)
