require "spreadsheet"

class HandleSheet
  def initialize(*args)
    @pass_coherence_checked = nil
    @input = ARGV[0]
    @output = ARGV[1] + ".xls" unless ARGV[1] =~ /\.xls$/ if ARGV[1] != nil
    @book_in = nil
    @book_out = nil
    @sheet_in = nil
    @sheet_out_us = nil
    @sheet_out_rmb = nil
    @sheet_name_rmb =nil
    @sheet_name_us = nil
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
      print "妞，无法生成目标表格文件，我也不晓得为啥!"
    end
  end

  def read_sheet(sheet_name_rmb, sheet_name_us) #指定读取数据源 哪个sheet
    @sheet_name_rmb = sheet_name_rmb
    @sheet_name_us = sheet_name_us
    @sheet_in_rmb = @book_in.worksheet sheet_name_rmb 
    @sheet_in_us = @book_in.worksheet sheet_name_us
  end

  def check__coherence_rmb
    j = 2# 循环行数用的
    @sheet_in_rmb.each 1 do |row|
      row[16].class.to_s =~ /Formula/i ? row[16] = row[16].value : row[16] = row[16] 
      #如果是Spread::Fomula类型就取它的value
      check_sellout = (row[4] =~ /out/i) && (row[16] <= 0)
      check_sellin = (row[4] =~ /in/i) && (row[16] >= 0)
      unless check_sellout || check_sellin
         puts "**注意数据源文件表:#{@sheet_name_rmb}的第#{j}行的价格正负不一致**"
         @pass_coherence_checked = true
      end
       # break j == 2000
      j += 1
    end
  end

  def adjust_sheet_rmb
    @sheet_out_rmb = @book_out.create_worksheet
    @sheet_out_rmb.name = "RMB sheet" #生成rmb的sheet

    i = 0
    @sheet_in_rmb.each do |row|
      object_row = row
      if i == 0
        object_row[13] = "PO Amount (RMB)"#原的row名字不对，生成对的row名字
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
    print "just for zhanfana 表格转换工具----Version 0.3\n\n"
  end

  def check_choherence_all
    open_file
    read_sheet('NB', 'NA')
    print_version
    puts "正在检查正负匹配 >>>>"
    check__coherence_rmb
    puts "检测完毕 >>>>"
  end

  def conver_all
    open_file
    read_sheet('NB', 'NA')
    print_version
    puts "正在生成文件 >>>>"
    adjust_sheet_rmb
    adjust_sheet_us
    wirte_file
    puts "<<#{@output}>>文件已经生成!"
  end
end

report = HandleSheet.new

if !(ARGV[0] || ARGV[1])
  print "使用方法:命令行下输入: \n检查价格一致性: ruby handle.rb input_file.xls(源文件)\n根据数据源生成新表: ruby handle.rb input.xls output.xls\n"
  exit
elsif ARGV[0] && ARGV[1] #有输入和输出的时候，进行转换
  report.conver_all
else
  report.check_choherence_all#只有输入的时候进行检查数据源的格式一致性
end

    


