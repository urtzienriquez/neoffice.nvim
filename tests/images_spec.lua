-- Images: shown as ![alt](path); replaced, added and removed by editing that.
local H = dofile("tests/helpers.lua")
local images = require("neoffice.images")
local zip = require("neoffice.zip")
local api = vim.api

local DOT = H.FIXTURES .. "dot.png" -- 40x20
local SQUARE = H.FIXTURES .. "square.jpg" -- 30x30

describe("image files", function()
  it("are recognised with their size", function()
    assert.same({ "png", 40, 20 }, { images.info(H.bytes(DOT)) })
    assert.same({ "jpeg", 30, 30 }, { images.info(H.bytes(SQUARE)) })
    assert.same({ "gif", 3, 5 }, { images.info("GIF89a\3\0\5\0rest") })
    assert.is_nil(images.info("not an image"))
    assert.is_nil(images.info(""))
  end)

  it("take as many cells as fit, keeping their shape", function()
    assert.same({ 60, 15 }, { images.cells(800, 400, 60, 20) }) -- wide: limited by width
    assert.same({ 40, 20 }, { images.cells(400, 400, 60, 20) }) -- square: limited by height
    assert.same({ 5, 1 }, { images.cells(40, 10, 60, 20) }) -- small: not enlarged beyond 8 pixels a cell
  end)
end)

for _, ext in ipairs({ "docx", "odt" }) do
  describe(ext, function()
    local path

    before_each(function()
      path = H.copy("rich." .. ext)
      H.open(path)
    end)

    local function image_path(n)
      return H.paras()[n]:match("!%[.-%]%((.-)%)")
    end

    it("shows images as ![alt](path to the file inside the document)", function()
      local inside = image_path(5)
      assert.truthy(inside)
      assert.truthy(zip.list(path)[inside])
    end)

    it("keeps an image untouched when its paragraph is edited", function()
      local before = H.paragraph_xml(path)[5]:match(ext == "docx" and "<w:drawing>.-</w:drawing>" or "<draw:frame.-</draw:frame>")
      H.replace(5, "An image", "A picture")
      H.reopen()
      assert.equals("A picture ![](" .. image_path(5) .. "){width=1.4cm} here.", H.paras()[5])
      local after = H.paragraph_xml(path)[5]:match(ext == "docx" and "<w:drawing>.-</w:drawing>" or "<draw:frame.-</draw:frame>")
      assert.equals(before, after)
    end)

    it("replaces an image with a file from disk (relative to the document or absolute)", function()
      vim.uv.fs_copyfile(SQUARE, vim.fs.dirname(path) .. "/square.jpg")
      H.replace(5, image_path(5), "square.jpg")
      H.reopen()
      local inside = image_path(5)
      assert.are_not.equals("square.jpg", inside)
      assert.equals(H.bytes(SQUARE), zip.read(path, inside))
      H.replace(5, inside, DOT)
      H.reopen()
      assert.equals(H.bytes(DOT), zip.read(path, image_path(5)))
    end)

    it("adds a new image, which then belongs to the document", function()
      api.nvim_buf_set_lines(0, H.row(12) - 1, H.row(12) - 1, false, { "![A red dot](" .. DOT .. ")", "" })
      H.save()
      local first = H.bytes(path)
      H.save() -- (saving again doesn't add it twice)
      assert.equals(first, H.bytes(path))
      H.open(path)
      local inside = image_path(12)
      assert.equals("![A red dot](" .. inside .. "){width=1.1cm}", H.paras()[12]) -- (40 pixels)
      assert.equals(H.bytes(DOT), zip.read(path, inside))
      if ext == "docx" then
        assert.truthy(zip.read(path, "[Content_Types].xml"):find('Extension="png"'))
      else
        assert.truthy(zip.read(path, "META-INF/manifest.xml"):find(inside, 1, true))
      end
    end)

    it("adds a JPEG image", function()
      H.set(12, "![](" .. SQUARE .. ")")
      H.reopen()
      assert.equals(H.bytes(SQUARE), zip.read(path, image_path(12)))
      if ext == "docx" then
        assert.truthy(zip.read(path, "[Content_Types].xml"):find('Extension="jpeg"'))
      end
    end)

    it("reuses an image already in the document", function()
      local inside = image_path(5)
      local files = vim.tbl_count(zip.list(path))
      H.set(12, "Again: ![](" .. inside .. ")")
      H.reopen()
      assert.equals("Again: ![](" .. inside .. "){width=1.1cm}", H.paras()[12])
      assert.equals(files, vim.tbl_count(zip.list(path)))
    end)

    it("resizes an image, keeping its shape, and only that", function()
      local function size()
        local main = zip.read(path, H.MAIN[ext])
        if ext == "docx" then
          local cx, cy = main:match('<wp:extent cx="(%d+)" cy="(%d+)"')
          return tonumber(cx) / 360000, tonumber(cy) / 360000
        end
        local frame = main:match("<draw:frame.->")
        return require("neoffice.markup").cm(frame:match('svg:width="([^"]+)"')), require("neoffice.markup").cm(frame:match('svg:height="([^"]+)"'))
      end
      local w, h = size()
      local before = H.paragraph_xml(path)
      H.replace(5, "{width=1.4cm}", "{width=7cm}")
      H.reopen()
      assert.truthy(H.paras()[5]:find("{width=7cm}", 1, true))
      local w2, h2 = size()
      assert.is_true(math.abs(w2 - 7) < 0.01)
      assert.is_true(math.abs(h2 / w2 - h / w) < 0.01)
      assert.equals(1, H.differing(before, H.paragraph_xml(path)))
    end)

    it("gives a new image the width you write (in cm, mm or in)", function()
      H.set(12, "![](" .. DOT .. "){width=50mm} and ![](" .. SQUARE .. "){width=2in}")
      H.reopen()
      assert.truthy(H.paras()[12]:find("{width=5cm}.*{width=5.1cm}"))
    end)

    it("takes a width in % of the text's width (the page's, without its margins)", function()
      local text_width = require("neoffice." .. ext).text_width(H.doc().ctx)
      local function width()
        local main = zip.read(path, H.MAIN[ext])
        if ext == "docx" then
          return tonumber(main:match('<wp:extent cx="(%d+)"')) / 360000
        end
        return require("neoffice.markup").cm(main:match("<draw:frame.-svg:width=\"([^\"]+)\""))
      end
      H.replace(5, "{width=1.4cm}", "{width=100%}")
      H.reopen()
      assert.is_true(math.abs(width() - text_width) < 0.01)
      H.paras() -- (shown in cm when opened again)
      assert.truthy(H.paras()[5]:find("{width=" .. require("neoffice.markup").width(text_width) .. "}", 1, true))
      H.replace(5, H.paras()[5]:match("{width=[^}]*}"), "{width=50%}")
      H.reopen()
      assert.is_true(math.abs(width() - text_width / 2) < 0.01)
    end)

    it("refuses a width it can't read, changing nothing", function()
      local before = H.bytes(path)
      H.replace(5, "{width=1.4cm}", "{width=big}")
      H.save()
      assert.equals(before, H.bytes(path))
      assert.is_true(vim.bo.modified)
    end)

    it("changes the alt text", function()
      H.replace(5, "![]", "![A tiny dot]")
      H.reopen()
      assert.truthy(H.paras()[5]:find("![A tiny dot](", 1, true))
    end)

    it("removes an image when its text is deleted", function()
      H.replace(5, "![](" .. image_path(5) .. "){width=1.4cm} ", "")
      H.reopen()
      assert.equals("An image here.", H.paras()[5])
      assert.falsy(zip.read(path, H.MAIN[ext]):find(ext == "docx" and "w:drawing" or "draw:frame"))
    end)

    it("refuses to save a file that isn't an image, changing nothing", function()
      local before = H.bytes(path)
      H.set(12, "![](" .. H.FIXTURES .. "make.sh)")
      H.save()
      assert.equals(before, H.bytes(path))
      assert.is_true(vim.bo.modified)
    end)

    if H.has_libreoffice then
      it("LibreOffice shows the added and replaced images", function()
        H.set(12, "![](" .. DOT .. ")")
        H.replace(5, image_path(5), SQUARE)
        H.save()
        local dir = vim.fn.tempname()
        local to = ext == "docx" and "odt" or "docx"
        vim.system({ "soffice", "-env:UserInstallation=file://" .. dir .. "/p", "--headless", "--convert-to", to, "--outdir", dir, path }):wait()
        local pictures = vim.tbl_filter(function(name) return name:match("^Pictures/") or name:match("^word/media/") end, vim.tbl_keys(zip.list(dir .. "/rich." .. to)))
        assert.equals(2, #pictures)
      end)
    end
  end)
end

describe("showing images", function()
  local calls, real_uis

  before_each(function()
    calls = {}
    vim.ui.img = { -- a stand-in for the terminal: records what would be drawn
      _supported = function() return true end,
      set = function(data, opts)
        table.insert(calls, { set = data, opts = opts })
        return #calls
      end,
      del = function(id) table.insert(calls, { del = id }) end,
    }
    real_uis = api.nvim_list_uis
    api.nvim_list_uis = function() return { {} } end
    vim.o.lines, vim.o.columns = 40, 100
  end)

  after_each(function()
    api.nvim_list_uis = real_uis
  end)

  it("reserves lines under the image and draws it there", function()
    local path = H.copy("rich.docx")
    H.open(path)
    api.nvim_buf_set_lines(0, 0, -1, false, { "Text above.", "", "![](" .. DOT .. ")", "", "Text below." })
    calls = {} -- (from here on)
    images.refresh(api.nvim_get_current_buf())
    vim.wait(500, function() return #calls > 0 end)
    assert.equals(1, #calls)
    local opts = calls[1].opts
    assert.same({ 5, 1 }, { opts.width, opts.height }) -- 40x20 pixels: 5x1 cells of about 8x16
    assert.equals(H.bytes(DOT), calls[1].set)
    local marks = api.nvim_buf_get_extmarks(0, api.nvim_get_namespaces().neoffice_images, 0, -1, { details = true })
    assert.equals(1, #marks)
    assert.equals(2, marks[1][2]) -- (under the image's line)
    assert.equals(1, #marks[1][4].virt_lines)
  end)

  it("shows nothing for a missing image, and stops when the buffer is hidden", function()
    H.open(H.copy("rich.docx"))
    api.nvim_buf_set_lines(0, 0, -1, false, { "![](missing.png)", "", "![](" .. DOT .. ")" })
    calls = {}
    images.refresh(api.nvim_get_current_buf())
    vim.wait(500, function() return #calls > 0 end)
    assert.equals(1, #calls)
    vim.cmd("enew")
    assert.truthy(calls[2] and calls[2].del)
  end)
end)
