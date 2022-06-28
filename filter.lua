local myTableBorderType  = "dashSmallGap"
local myTableBorderColor = "898989"
local myTableWidthPct    = 9000 -- table table percent: [0, 10000)

function addWordStyle(text, style)
    local word = string.match(text, "^<w:r><w:rPr>(.*)</w:r>$")
    if word then
        -- printf("addWordStyle, text: ".. text .. ", to: "..'<w:r><w:rPr>' .. style .. word .. '</w:r>')
        return pandoc.RawInline("openxml", '<w:r><w:rPr>' .. style .. word .. '</w:r>')
    end
    return nil
end

function aroundWordStyle(text, style)
    -- printf("aroundWordStyle, text: "..text..", to: "..'<w:r><w:rPr>' .. style .. '</w:rPr><w:t>' .. text .. "</w:t></w:r>")
    return pandoc.RawInline("openxml", '<w:r><w:rPr>' .. style .. '</w:rPr><w:t>' .. text .. "</w:t></w:r>")
end

function formatColor(color)
    if color == "red" then
        return "ff0000"
    elseif color == "green" then
        return "00ff00"
    elseif color == "blue" then
        return "0000ff"
    else
        return color
    end
end

local fontStack = {}
function RawInline(elem)
    if FORMAT ~= "docx" then
        return elem
    end

    local fontStart = string.match(elem.text, '<font')
    if fontStart then
        local color = string.match(elem.text, 'color%s*=%s*[\'"]#?([%w]+)[\'"]')
        if color then
            table.insert(fontStack, formatColor(color))
        else
            color = string.match(elem.text, 'color%s*:%s*#?([%w]+)')
            if color then
                table.insert(fontStack, formatColor(color))
            else
                table.insert(fontStack, "")
            end
        end
    else
        local fontEnd = string.match(elem.text, '</font')
        if fontEnd then
            table.remove(fontStack, #fontStack)
        end
    end
    return elem
end

function Str(elem)
    if FORMAT ~= "docx" then
        return elem
    end

    elem.text = string.gsub(elem.text, "<", "&lt;")
    elem.text = string.gsub(elem.text, ">", "&gt;")
    if #fontStack and fontStack[#fontStack] ~= '' then
        local color = fontStack[#fontStack]
        if color then
            return aroundWordStyle(elem.text, '<w:color w:val="' .. color .. '"/>')
        end
    end
    return elem
end

function Strong(elem)
    if FORMAT ~= "docx" then
        return elem
    end

    for i, elemContent in ipairs(elem.content) do
        if elemContent.t == "Str" then
            elem.content[i] = aroundWordStyle(elemContent.text, '<w:b/><w:bCs/>')
        elseif elemContent.t == "RawInline" then
            local replace = addWordStyle(elemContent.text, '<w:b/><w:bCs/>')
            if replace then
                elem.content[i] = replace
            end
        end
    end
    return elem
end

function Emph(elem)
    if FORMAT ~= "docx" then
        return elem
    end

    for i, elemContent in ipairs(elem.content) do
        if elemContent.t == "Str" then
            elem.content[i] = aroundWordStyle(elemContent.text, '<w:i/><w:iCs/>')
        elseif elemContent.t == "RawInline" then
            local replace = addWordStyle(elemContent.text, '<w:i/><w:iCs/>')
            if replace then
                elem.content[i] = replace
            end
        end
    end
    return elem
end

function Header(elem)
    if FORMAT ~= "docx" then
        return elem
    end

    for i, content in ipairs(elem.content) do
        if content.t == "RawInline" then
            local replace = addWordStyle(content.text, '<w:pStyle w:val="Heading' .. elem.level .. '"/>')
            if replace then
                elem.content[i] = replace
            end

        elseif content.t == "Str" then
            elem.content[i] = aroundWordStyle(content.text, '<w:pStyle w:val="Heading' .. elem.level .. '"/>')

        elseif content.t == "Emph" then
            for j = 1, #content.content do
                local emphContent = content.content[j]
                local replace     = addWordStyle(emphContent.text, '<w:pStyle w:val="Heading' .. elem.level .. '"/>')
                if replace then
                    elem.content[i].content[j] = replace
                end
            end
        elseif content.t == "Strong" then
            for j = 1, #content.content do
                local emphContent = content.content[j]
                local replace     = addWordStyle(emphContent.text, '<w:pStyle w:val="Heading' .. elem.level .. '"/>')
                if replace then
                    elem.content[i].content[j] = replace
                end
            end
        end
    end
    return elem
end

local MyTableBlock = {}

function MyTableBlock:create()
    local tmp = { blocks = {} }
    setmetatable(tmp, self)
    self.__index = self
    return tmp
end

function MyTableBlock:tableStart()
    local xml = [[
        <w:tbl>
            <w:tblPr>
                <w:tblStyle w:val="TableGrid"/>
                <w:tblW w:w="{myTableWidthPct}" w:type="pct"/>
                <w:jc w:val="center"/>
                <w:tblBorders>
                    <w:top w:val="{myTableBorderType}" w:space="0" w:sz="4" w:color="{myTableBorderColor}"/>
                    <w:left w:val="{myTableBorderType}" w:space="0" w:sz="4" w:color="{myTableBorderColor}"/>
                    <w:bottom w:val="{myTableBorderType}" w:space="0" w:sz="4" w:color="{myTableBorderColor}"/>
                    <w:right w:val="{myTableBorderType}" w:space="0" w:sz="4" w:color="{myTableBorderColor}"/>
                    <w:insideH w:val="{myTableBorderType}" w:space="0" w:sz="4" w:color="{myTableBorderColor}"/>
                    <w:insideV w:val="{myTableBorderType}" w:space="0" w:sz="4" w:color="{myTableBorderColor}"/>
                </w:tblBorders>
            </w:tblPr>
        ]]

    xml       = string.gsub(xml, "{myTableWidthPct}", myTableWidthPct)
    xml       = string.gsub(xml, "{myTableBorderType}", myTableBorderType)
    xml       = string.gsub(xml, "{myTableBorderColor}", myTableBorderColor)
    table.insert(self.blocks, pandoc.RawBlock("openxml", xml))
end

function MyTableBlock:tableEnd()
    table.insert(self.blocks, pandoc.RawBlock("openxml", '</w:tbl>'))
end

function MyTableBlock:trHeadStart()
    table.insert(self.blocks, pandoc.RawBlock("openxml", [[
        <w:tr>
            <w:tblPrEx>
            </w:tblPrEx>
            <w:trPr>
                <w:tblHeader w:val="true"/>
            </w:trPr>
        ]]))
end

function MyTableBlock:trStart()
    table.insert(self.blocks, pandoc.RawBlock("openxml", [[
        <w:tr>
        <w:trPr/>
        ]]))
end

function MyTableBlock:trEnd()
    table.insert(self.blocks, pandoc.RawBlock("openxml", '</w:tr>'))
end

function MyTableBlock:cellStart(cellNum)
    local cellStart    = [[
        <w:tc>
            <w:tcPr>
                <w:tcW w:w="{cellWidth}" w:type="pct"/>
            </w:tcPr>
            <w:p>
            <w:pPr>
                <w:jc w:val="center"/>
            </w:pPr>
        ]]
    local cellWidth    = math.floor(myTableWidthPct / cellNum)
    local cellStartXml = string.gsub(cellStart, "{cellWidth}", cellWidth)
    table.insert(self.blocks, pandoc.RawBlock("openxml", cellStartXml))
end

function MyTableBlock:cellEnd()
    table.insert(self.blocks, pandoc.RawBlock("openxml", '</w:p></w:tc>'))
end

function MyTableBlock:addNewRawBlock(format, text)
    table.insert(self.blocks, pandoc.RawBlock(format, text))
end

function MyTableBlock:addRawBlock(rawblock)
    table.insert(self.blocks, rawblock)
end

function MyTableBlock:getBlocks()
    return self.blocks
end

function handleTableCells(myTable, cells)
    for _, cell in ipairs(cells) do
        myTable:cellStart(#cells)
        for _, cellContent in ipairs(cell.contents) do
            for _, cellContentContent in ipairs(cellContent.content) do
                if cellContentContent.t == "Str" then
                    myTable:addNewRawBlock("openxml", '<w:r><w:rPr/><w:t>' .. cellContentContent.text .. "</w:t></w:r>")

                elseif cellContentContent.t == "RawInline" then
                    myTable:addNewRawBlock(cellContentContent.format, cellContentContent.text)

                elseif cellContentContent.t == "RawBlock" then
                    myTable:addRawBlock(cellContentContent)

                elseif cellContentContent.t == "Strong" or cellContentContent.t == "Emph" then
                    for _, cellContentContentContent in ipairs(cellContentContent.content) do
                        if cellContentContentContent.t == "RawInline" then
                            cellContentContentContent = pandoc.RawBlock(cellContentContentContent.format, cellContentContentContent.text)
                        end
                        myTable:addRawBlock(cellContentContentContent)
                    end
                end
            end
        end
        myTable:cellEnd()
    end
end

function printf(s)
    local out = io.open("/privateFile/config/Pandoc/lua-out.txt", "a")
    out:write(os.date("%H:%M:%S") .. "  " .. s .. "\n")
    out:flush()
end

function Table(dTable)

    if FORMAT ~= "docx" then
        return dTable
    end

    local myTable = MyTableBlock:create()
    myTable:tableStart()

    -- head
    if dTable.head and #dTable.head.rows > 0 then
        for _, row in ipairs(dTable.head.rows) do
            myTable:trHeadStart()
            handleTableCells(myTable, row.cells)
            myTable:trEnd()
        end
    end

    if dTable.bodies and #dTable.bodies > 0 then
        for _, tBody in ipairs(dTable.bodies) do
            if #tBody.body > 0 then
                for _, row in ipairs(tBody.body) do
                    myTable:trStart()
                    handleTableCells(myTable, row.cells)
                    myTable:trEnd()
                end
            end
        end
    end

    myTable:tableEnd()
    return myTable:getBlocks()
end

