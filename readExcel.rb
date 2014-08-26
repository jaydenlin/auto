require 'spreadsheet'
require 'nokogiri'
require 'open-uri'


Spreadsheet.client_encoding = 'UTF-8'

pathRoot="/mnt/taiwan-rakuten/Service_Platform_Division/Development_Department/Department\ _CrossTeam"
excelPath="#{pathRoot}/Hare/Event/data/EventPage_Single.xls"
htmlPath="#{pathRoot}/Hare/Event/html/pc_template.html"

book = Spreadsheet.open "#{excelPath}"

allSheets=book.worksheets

def getValueFromCell(cell)
    if cell.class==Spreadsheet::Formula
       return cell.value
    else
       return cell
    end
end

for sheet in allSheets
    #p sheet.name

    if sheet.name=="TopProducts"
        for cell in sheet.row(6)
             if cell!=nil
                p getValueFromCell cell
             end
        end
    end
    #sheet.each 2 do |row|
    #     p row[3]
         # do something interesting with a row
    #end

end
#s = SimpleSpreadsheet::Workbook.read("#{excelPath}")
