require "spreadsheet"

unless ARGV[0] || ARGV[1] 
  print "使用方法：\n命令行下输入: ruby handle.rb input_file(源文件) output_file(目标文件)\n"
else

begin
  book_in = Spreadsheet.open ARGV[0]
rescue => open_source_file_error
  print "妞，你确认有这个文件吗\n"
end

sheet_in = book_in.worksheet 0 #指定读取哪个sheet

book_out = Spreadsheet::Workbook.new #创建一个新的输出book
sheet_out = book_out.create_worksheet
sheet_out.name = "zhangfana"

other_row_format = Spreadsheet::Format.new :bottom => :thin,
                                  :top => :thin,
                                  :left => :thin,
                                  :right => :thin,
                                  :bottom_color => :black,
                                  :top_color => :black,
                                  :left_color => :black,
                                  :right_color => :black

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
    sheet_out.row(i).default_format = other_row_format #设置其它row 默认格式
    i += 1
   
  end
firet_row_format = Spreadsheet::Format.new :weight => :bold,
                                  :pattern_fg_color => :lime,
                                  :pattern => 1,

                                  :bottom => :thin,
                                  :top => :thin,
                                  :left => :thin,
                                  :right => :thin,
                                  :bottom_color => :black,
                                  :top_color => :black,
                                  :left_color => :black,
                                  :right_color => :black

sheet_out.row(0).default_format = firet_row_format
print "报表转换程序 for zhanfana Version 0.1\n"
book_out.write ARGV[1]

end

