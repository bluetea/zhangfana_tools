require "spreadsheet"
require 'roo'



book_in = Spreadsheet.open 'summary.xls'
sheet_in = book_in.worksheet 0 #指定读取哪个sheet

book_out = Spreadsheet::Workbook.new #创建一个新的输出book
sheet_out = book_out.create_worksheet
sheet_out.name = "zhangfana"



i = 0
  sheet_in.each do |row|
    object_row = row
    if i == 0
      object_row[13] = "PO Amount (RMB)"#生成的新文件row名字变了

    end
    object_row.delete_at(10)
    object_row.delete_at(10)

    ary = object_row.to_a  #把类型必须转换
    ary.each do |item|
      sheet_out.row(i).push item
    end
    i += 1
   
  end




book_out.write "test_out.xls"

