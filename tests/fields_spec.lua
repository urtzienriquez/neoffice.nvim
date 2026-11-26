-- Fields (citations from Zotero or Mendeley, dates, references...): shown as ⟦their text⟧, kept as
-- they are; the text around them is edited, they aren't, and they can't be typed.
local H = dofile("tests/helpers.lua")
local zip = require("neoffice.zip")
local api = vim.api

local NAME = 'ZOTERO_ITEM CSL_CITATION {&quot;citationID&quot;:&quot;x&quot;} RNDa1'
local CITED = {
  odt = function(name, text)
    if name ~= "content.xml" then
      return text
    end
    local p = '<text:p>Energy matters <text:reference-mark-start text:name="' .. NAME .. '"/>(1)'
      .. '<text:reference-mark-end text:name="' .. NAME .. '"/> a lot.</text:p><text:p>Another one.</text:p>'
    return (text:gsub("</office:text>", p .. "</office:text>", 1))
  end,
  docx = function(name, text)
    if name ~= "word/document.xml" then
      return text
    end
    local p = '<w:p><w:r><w:t xml:space="preserve">Energy matters </w:t></w:r>'
      .. '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> ADDIN ZOTERO_ITEM CSL_CITATION {"citationID":"x"} </w:instrText></w:r>'
      .. '<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>(1)</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>'
      .. '<w:r><w:t xml:space="preserve"> a lot.</w:t></w:r></w:p><w:p><w:r><w:t>Another one.</w:t></w:r></w:p>'
    return (text:gsub("<w:sectPr", p .. "<w:sectPr", 1))
  end,
}
local ZOTERO = { odt = "ZOTERO_ITEM CSL_CITATION", docx = "ADDIN ZOTERO_ITEM" }

for _, ext in ipairs({ "odt", "docx" }) do
  describe(ext .. " citations (Zotero)", function()
    local path
    local function at(text) return vim.fn.index(H.paras(), text) + 1 end
    local function citations()
      local _, n = zip.read(path, H.MAIN[ext]):gsub(vim.pesc(ZOTERO[ext]), "")
      return ext == "odt" and n / 2 or n -- (an .odt's: a start and an end)
    end
    before_each(function()
      path = H.copy("rich." .. ext, CITED[ext])
      H.open(path)
    end)

    it("shows a citation as ⟦its text⟧, and keeps it when the text around it is edited", function()
      local k = at("Energy matters ⟦(1)⟧ a lot.")
      assert.truthy(k > 0, table.concat(H.paras(), "\n"))
      H.replace(k, "a lot", "a great deal")
      H.reopen()
      assert.truthy(at("Energy matters ⟦(1)⟧ a great deal.") > 0)
      assert.equals(1, citations())
    end)

    it("keeps it when its line is typed again (cc), or moved with dd and p", function()
      local k = at("Energy matters ⟦(1)⟧ a lot.")
      H.set(k, "Energy now matters ⟦(1)⟧ a lot.") -- (the whole line replaced)
      H.save()
      assert.equals(1, citations())
      local r = H.row(k)
      local line = H.lines()[r]
      api.nvim_buf_set_lines(0, r - 1, r + 1, false, {}) -- (dd: it and the blank line after it)
      local other = H.row(at("Another one."))
      api.nvim_buf_set_lines(0, other, other, false, { "", line })
      H.reopen()
      assert.same({ "Another one.", "Energy now matters ⟦(1)⟧ a lot." }, vim.list_slice(H.paras(), #H.paras() - 1))
      assert.equals(1, citations())
    end)

    it("doesn't write a field typed by hand, nor a second copy: the save says so and changes nothing", function()
      local before = H.bytes(path)
      local errors = {}
      local notify = vim.notify
      vim.notify = function(msg) errors[#errors + 1] = msg end
      H.set(at("Another one."), "Another one ⟦(2)⟧.")
      H.save()
      assert.equals(before, H.bytes(path))
      assert.truthy(table.concat(errors, "\n"):find("can't be typed", 1, true))
      H.set(at("Another one ⟦(2)⟧."), "Another one \\⟦(2)\\⟧.") -- (the characters themselves: escaped)
      H.save()
      vim.notify = notify
      H.reopen()
      assert.truthy(at("Another one \\⟦(2)\\⟧.") > 0)
      assert.equals(1, citations())
    end)
  end)
end
