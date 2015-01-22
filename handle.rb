require "spreadsheet"

class HandleSheet
  def initialize(*args)
    @input = ARGV[0]
    @output = ARGV[1]
    @book_in = nil
    @book_out = nil
    @sheet_in = nil
    @sheet_out_us = nil
    @sheet_out_rmb = nil
    @other_row_format = Spreadsheet::Format.new :bottom => :thin,
                                  :top => :thin,
                                  :left => :thin,
                                  :right => :thin,
                                  :bottom_color => :black,
                                  :top_color => :black,
                                  :left_color => :black,
                                  :right_color => :black

    @firet_row_format = Spreadsheet::Format.new :weight => :bold,
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

  end

  def open_file #测试看能否打开源文件和生成目标文件
    begin
      @book_in = Spreadsheet.open @input  #打开一个源文件
    rescue => open_source_file_error
      print "妞，你确认有这个文件吗\n"
    end

    begin
      @book_out = Spreadsheet::Workbook.new #创建一个新的输出book
    rescue => open_destination_file_error
      print "大姐，无法生成目标表格文件，我也不晓得为啥"
    end
  end

  def read_sheet(sheet_name_rmb, sheet_name_us) #指定读取数据源 哪个sheet
    @sheet_in_rmb = @book_in.worksheet sheet_name_rmb 
    @sheet_in_us = @book_in.worksheet sheet_name_us
  end

  def adjust_sheet_rmb
    @sheet_out_rmb = @book_out.create_worksheet
    @sheet_out_rmb.name = "RMB sheet" #生成rmb的sheet

    i = 0
    @sheet_in_rmb.each do |row|
      object_row = row
      if i == 0
        object_row[13] = "PO Amount (RMB)"#生成的新文件row名字变了
      end
      object_row.delete_at(10)
      object_row.delete_at(10)

      ary = object_row.to_a  #把类型必须转换
      ary.each do |item|
        @sheet_out_rmb.row(i).push item
      end
      i += 1
      @sheet_out_rmb.row(i).default_format = @other_row_format #设置其它row 默认格式

    end
    @sheet_out_rmb.row(0).default_format = @firet_row_format #设置默认row的默认格式
  end

  def adjust_sheet_us
    @sheet_out_us = @book_out.create_worksheet
    @sheet_out_us.name= "US sheet"

  end

  def wirte_file
    @book_out.write @output
  end

  def print_version
    print "报表转换程序 for zhanfana 支持us和rmb同时转换\nVersion 0.2\n"
  end
end


unless ARGV[0] || ARGV[1] 
  print "使用方法：\n命令行下输入: ruby handle.rb input_file(源文件) output_file(目标文件)\n"
end

report = HandleSheet.new
report.open_file
report.read_sheet('NB', 'NA')
report.print_version
report.adjust_sheet_rmb
report.adjust_sheet_us
report.wirte_file


