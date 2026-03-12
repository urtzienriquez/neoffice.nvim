-- neoffice/conceal.lua
-- XML tag concealment using extmarks (inspired by render-markdown)

local M = {}

---@class neoffice.Conceal
---@field private buf integer
---@field private ns integer
---@field private enabled boolean
local Conceal = {}
Conceal.__index = Conceal

---Create new Conceal object for buffer
---@param buf integer
---@return neoffice.Conceal
function Conceal.new(buf)
  local self = setmetatable({}, Conceal)
  self.buf = buf
  self.ns = vim.api.nvim_create_namespace("neoffice_conceal")
  self.enabled = false
  return self
end

---Scan buffer and apply concealment using extmarks
function Conceal:apply()
  if not self.enabled then
    return
  end

  vim.api.nvim_buf_clear_namespace(self.buf, self.ns, 0, -1)
  local lines = vim.api.nvim_buf_get_lines(self.buf, 0, -1, false)

  local row = 0
  while row < #lines do
    local line = lines[row + 1]
    local col = 1

    while col <= #line do
      -- Find the start of any tag
      local tag_start = line:find("<", col)
      if not tag_start then
        break
      end

      -- 1. Identify what kind of tag we are dealing with
      local is_ann_start = line:find("^office:annotation[%s>]", tag_start + 1)
      local is_ann_end = line:find("^office:annotation%-end", tag_start + 1)

      -- 2. Define our target closing marker
      -- If it's a comment start, we hide EVERYTHING until the final closing tag
      -- Otherwise, we just hide until the next '>'
      local target = is_ann_start and "</office:annotation>" or ">"

      local end_row, end_col = -1, -1

      -- 3. Search for the target across this and subsequent lines
      for r = row, #lines - 1 do
        local search_line = lines[r + 1]
        local start_search = (r == row) and tag_start or 1
        local _, match_end = search_line:find(target, start_search, true)

        if match_end then
          end_row, end_col = r, match_end
          break
        end
      end

      -- 4. Apply the concealment if the closing marker was found
      if end_row ~= -1 then
        pcall(vim.api.nvim_buf_set_extmark, self.buf, self.ns, row, tag_start - 1, {
          end_line = end_row,
          end_col = end_col,
          conceal = (is_ann_start or is_ann_end) and "💬" or "",
          priority = (is_ann_start or is_ann_end) and 200 or 100,
        })

        -- Update our loop positions to skip past the concealed block
        row = end_row
        line = lines[row + 1]
        col = end_col + 1
      else
        -- Fallback: if no closing tag found, move past the '<' to avoid infinite loops
        col = tag_start + 1
      end
    end
    row = row + 1
  end
end

---Enable concealment
function Conceal:enable()
  self.enabled = true

  -- Set window options for concealment
  local wins = vim.fn.win_findbuf(self.buf)
  for _, win in ipairs(wins) do
    vim.api.nvim_set_option_value("conceallevel", 2, { win = win })
    vim.api.nvim_set_option_value("concealcursor", "", { win = win }) -- Show in all modes
  end

  self:apply()
end

---Disable concealment
function Conceal:disable()
  self.enabled = false
  vim.api.nvim_buf_clear_namespace(self.buf, self.ns, 0, -1)

  -- Reset window options
  local wins = vim.fn.win_findbuf(self.buf)
  for _, win in ipairs(wins) do
    vim.api.nvim_set_option_value("conceallevel", 0, { win = win })
    vim.api.nvim_set_option_value("concealcursor", "", { win = win })
  end
end

---Refresh concealment (re-scan and re-apply)
function Conceal:refresh()
  if self.enabled then
    self:apply()
  end
end

---Toggle concealment on/off
function Conceal:toggle()
  if self.enabled then
    self:disable()
  else
    self:enable()
  end
end

---Get statistics about concealment
---@return table
function Conceal:stats()
  local extmarks = vim.api.nvim_buf_get_extmarks(self.buf, self.ns, 0, -1, { details = false })

  return {
    enabled = self.enabled,
    total_extmarks = #extmarks,
  }
end

-- Module functions

---@type table<integer, neoffice.Conceal>
local conceals = {}

---Get or create Conceal object for buffer
---@param buf integer
---@return neoffice.Conceal
function M.get(buf)
  if not conceals[buf] then
    conceals[buf] = Conceal.new(buf)
  end
  return conceals[buf]
end

---Enable concealment for buffer
---@param buf integer
function M.enable(buf)
  M.get(buf):enable()
end

---Disable concealment for buffer
---@param buf integer
function M.disable(buf)
  M.get(buf):disable()
end

---Toggle concealment for buffer
---@param buf integer
function M.toggle(buf)
  M.get(buf):toggle()
end

---Refresh concealment for buffer
---@param buf integer
function M.refresh(buf)
  M.get(buf):refresh()
end

---Get stats for buffer
---@param buf integer
---@return table
function M.stats(buf)
  return M.get(buf):stats()
end

---Clean up when buffer is deleted
---@param buf integer
function M.cleanup(buf)
  if conceals[buf] then
    conceals[buf]:disable()
    conceals[buf] = nil
  end
end

return M
