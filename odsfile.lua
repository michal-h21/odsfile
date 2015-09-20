module(...,package.seeall)
require "zip"
local xmlparser = require ("luaxml-mod-xml")
local handler = require("luaxml-mod-handler")

local namedRanges = {}

function load(filename)
  local p = {
    file = zip.open(filename),
    content_file_name = "content.xml",
    loadContent = function(self,filename)
      local treehandler = handler.simpleTreeHandler()
      local filename = filename or self.content_file_name  
      local xmlfile = self.file:open(filename)
      local text = xmlfile:read("*a")
      local xml = xmlparser.xmlParser(treehandler)
      xml:parse(text)
      return treehandler
    end
  }
  return p
end

function getTable(x,table_name)
  local t = getTable0(x,table_name)
  local t2 = {}

  for key, val in pairs(t) do
    if key == "table:table-row" then
      local rows = {}
      
      for i = 1, #val do
        local r = val[i]
        local rowRep = r["_attr"]["table:number-rows-repeated"] or 1

        row = {}
        row["_attr"] = r["_attr"]
        local cc = r["table:table-cell"] or {}
        if #cc == 0 then
          cc = {cc}
        end
        
        local columns = {}
        --for j = 1, #cc do
         -- local c = cc[j]
				for _, c in ipairs(cc) do
					c["_attr"] = c["_attr"] or {}
          local colRep = c["_attr"]["table:number-columns-repeated"] or 1
          for k = 1, colRep, 1 do
            table.insert(columns, c)
          end
        end
        row["table:table-cell"] = columns
        
        for j = 1, rowRep, 1 do
          table.insert(rows, row)
        end
      end
      
      t2[key] = rows
    else
      t2[key] = val
    end
  end

  return t2
end

function getTable0(x,table_name)
  local tables = x.root["office:document-content"]["office:body"]["office:spreadsheet"]["table:table"]
  namedRanges = loadNameRanges(x, table_name)
  if #tables > 1 then
    if type(tables) == "table" and table_name ~= nil then 
      for k,v in pairs(tables) do
          if(v["_attr"]["table:name"]==table_name) then
            return v, k
          end 
        end
    elseif type(tables) == "table" and table_name == nil then
      return tables[1], 1  
    else 
      return tables  
    end
  else 
    return tables
  end
end

function getColumnCount(tbl)
  local tbl = tbl or {}
  local columns = tbl["table:table-column"] or {}
  local x = 0
  for _, c in pairs(columns) do
    local rep = c["table:number-columns-repeated"] or 1
    x = x + rep
  end
  return x
end

function loadNameRanges(root, tblname)
  local tblname = tblname or ""
  local t = {}
  local ranges = root.root["office:document-content"]["office:body"]["office:spreadsheet"]["table:named-expressions"]
  if not ranges then return {} end
  ranges = ranges["table:named-range"] or {}
  if #ranges == 0 then 
    ranges = {ranges}
  end
  for _,r in ipairs(ranges) do
    local a = r["_attr"] or {}
    local range = a["table:cell-range-address"] or ""
    local name = a["table:name"] 
    if name and range:match("^"..tblname) then
      range = range:gsub("^[^%.]*",""):gsub("[%$%.]","")
      print("named range", name, range)
      t[name] = range
    end
  end
  return t
end





function tableValues(tbl,x1,y1,x2,y2)
  local t= {}
  local x1 = x1 or 1
  local x2 = x2 or getColumnCount(tbl)
  if type(tbl["table:table-row"])=="table" then
    local rows = table_slice(tbl["table:table-row"],y1,y2)
    for k,v in pairs(rows) do
      -- In every sheet, there are two rows with no data at the bottom, we need to strip them
      if(v["_attr"] and v["_attr"]["table:number-rows-repeated"] and tonumber(v["_attr"]["table:number-rows-repeated"])>10000) then break end
      local j = {}
      if #v["table:table-cell"] > 1 then
        local r = table_slice(v["table:table-cell"],x1,x2)
        for p,n in pairs(r) do
          local attr = n["_attr"]
          local cellValue = n["text:p"] or ""
          table.insert(j,{value=cellValue,attr=attr})
        end
      else
        local p = {value=v["table:table-cell"]["text:p"],attr=v["table:table-cell"]["_attr"]} 
        table.insert(j,p) 
      end
      table.insert(t,j)
    end
  end
  return t
end

function getRange(range)
  local range = namedRanges[range] or range
  local r = range:lower()
  local function getNumber(s)
    if s == "" or s == nil then return nil end
    local f,ex = 0,0
    for i in string.gmatch(s:reverse(),"(.)") do
      f = f + (i:byte()-96) * 26 ^ ex
      ex = ex + 1 
    end
    return f
  end
  for x1,y1,x2,y2 in r:gmatch("(%a*)(%d*):*(%a*)(%d*)") do
    return getNumber(x1),tonumber(y1),getNumber(x2),tonumber(y2) 
   --print(string.format("%s, %s, %s, %s",getNumber(x1),y1,getNumber(x2),y2))
  end
end

function table_slice (values,i1,i2)
  -- Function from http://snippets.luacode.org/snippets/Table_Slice_116
  local res = {}
  local n = #values
  -- default values for range
  i1 = i1 or 1
  i2 = i2 or n
  if i2 < 0 then
    i2 = n + i2 + 1
  elseif i2 > n then
    i2 = n
  end
  if i1 < 1 or i1 > n then
    return {}
  end
  local k = 1
  for i = i1,i2 do
    res[k] = values[i]
    k = k + 1
  end
  return res
end

function interp(s, tab)
  return (s:gsub('(-%b{})', 
    function(w) 
      s = w:sub(3, -2)
      s = tonumber(s) or s
      return tab[s] or w 
    end)
  )
end

get_link = function(val)
  local k = val["text:a"][1]
  local href = val["text:a"]["_attr"]["xlink:href"]
  return "\\odslink{"..href.."}{"..k.."}"
end

function escape(s)
  return string.gsub(s, "([#%%$&])", "\\%1")
end

function get_cell(val, delim)
  local val = val or ""
  local typ = type(val)
  if typ == "string" then
    return escape(val)
  elseif typ == "table" then
    if val["text:a"] then
      return get_link(val)
    elseif val["text:span"] then
      return get_cell(val["text:span"], delim)
    elseif val["text:s"] then
      -- return get_cell(val["text:s"], delim)
      return table.concat(val, " ")
    else
      local t = {}
      for _,v in ipairs(val) do
        local c = get_cell(v, delim)
        table.insert(t, c)
      end
      return table.concat(t,delim)
    end
  end
end

-- Interface for adding new rows to the spreadsheet

function newRow()
  local p = {
    pos = 0,
    cells = {},
    -- Generic function for inserting cell
    addCell = function(self,val, attr,pos)
      if pos then
        table.insert(self.cells,pos,{["text:p"] = val, ["_attr"] = attr})
        self.pos = pos
      else
        self.pos = self.pos + 1
        table.insert(self.cells,self.pos,{["text:p"] = val, ["_attr"] = attr})
      end
    end, 
    addString = function(self,s,attr,pos)
      local attr = attr or {}
      attr["office:value-type"] = "string"
      self:addCell(s,attr,pos)
    end,
    addFloat = function(self,i,attr,pos)
      local attr = attr or {}
      local s = tonumber(i) or 0
      s = tostring(s)
      attr["office:value-type"] = "float"
      attr["office:value"] = s
      self:addCell(s,attr,pos)
    end, 
    findLastRow = function(self,sheet)
      for i= #sheet["table:table-row"],1,-1 do
        if sheet["table:table-row"][i]["_attr"]["table:number-rows-repeated"] then
          return i
        end
      end
      return #sheet["table:table-row"]+1
    end,
    insert = function(self, sheet, pos)
      local t = {}
      local pos = pos or self:findLastRow(sheet)
      print("pos je: ",pos)
      if sheet["table:table-column"]["_attr"] and sheet["table:table-column"]["_attr"]["table:number-columns-repeated"] then
        table_columns = sheet["table:table-column"]["_attr"]["table:number-columns-repeated"]
      else 
        table_columns = #sheet["table:table-column"]
      end
      for i=1, table_columns do
        table.insert(t,self.cells[i] or {})  
      end
      t = {["table:table-cell"]=t}
      table.insert(sheet["table:table-row"],pos,t)
    end
  }
  return p
end


-- function for updateing the archive. Depends on external zip utility
function updateZip(zipfile, updatefile)
  local command  =  string.format("zip %s %s",zipfile, updatefile)
  print ("Updating an ods file.\n" ..command .."\n Return code: ", os.execute(command))  
end
