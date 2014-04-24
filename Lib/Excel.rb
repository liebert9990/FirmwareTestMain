require 'win32ole'
#
#  Excel class to get and set testcase infomation.
#
# @example
# ex = Excel.new('d:\abc.xlsx')
#
class Excel

  attr_reader :pass_fail
  attr_reader :username
  attr_reader :password
  attr_writer :write_starttime

  def initialize(path_name)
    @path_name = path_name
    @sheet = 1
  end

  def username

    user =  @worksheet.Range('b4').value
    return user

  end

  def password

    pass=  @worksheet.Range('b5').value
    return pass

  end

  def pass_fail

    pass=  @worksheet.Range('b9').value
    return pass

  end

  def write_state(state)
    @worksheet.Range('b9').value = state
  end
  
  def write_starttime(starttime)
    @worksheet.Range("b4").value = starttime
  end

  #
  # Creat an excel.
  #
  # @example
  #   ex = Excel.new('d:\abc.xlsx')
  #   ex.creat_xls
  #

  def creat_xls

    @spreatsheet = WIN32OLE::new('excel.Application')
    @workbook = @spreatsheet.Workbooks.add
    @worksheet = @workbook.Worksheets(@sheet)
    @spreatsheet.visible = true

  end

  #
  # Open an excel.
  #
  # @example
  #   ex = Excel.new('d:\abc.xlsx')
  #   ex.open
  #

  def open_xls

    @spreatsheet = WIN32OLE::new('excel.Application')
    @workbook = @spreatsheet.Workbooks.open(@path_name)
    @worksheet = @workbook.Worksheets(@sheet)
    @spreatsheet.visible = true

  end

  #
  # Save an excel.
  #
  # @example
  #   ex = Excel.new('d:\abc.xlsx')
  #   ex.save_xls
  #

  def save_xls
    @workbook.save
  end

  #
  # Save as an excel to an other path.
  #
  # @example
  #   ex = Excel.new('d:\abc.xlsx')
  #   ex.save_as_xls('c:\abc.xlsx')
  #

  def save_as_xls(save_path)
    @workbook.saveas(save_path)
  end

  #
  # Close excel
  #
  # @example
  #   ex = Excel.new('d:\abc.xlsx')
  #   ex.close_xls
  #

  def close_xls
    @workbook.close
    @spreatsheet.quit
  end

  #
  # Active xls
  #
  # @example
  #   ex = Excel.new('d:\abc.xlsx')
  #   ex.open_xls
  #   ex.active_xls
  # 

  def active_xls
    @workbook.activate
  end

  #
  # Read testcase infomation from excel, return array
  #
  # @example
  #   ex = Excel.new('d:\abc.xlsx')
  #   a.open_xls
  #   ex.read_testcase
  # return [{"Test Steps"=>1.0, "Type"=>"; Navigate to the test page", "PageName"=>nil, "Component"=>nil, "Action"=>nil, "ARG"=>[nil, nil, nil, nil], "Actual Output"=>[nil, nil]}, {"Test Steps"=>2....
  #
  def read_testcase

    temp = []
    rows = 1 +  @worksheet.Range('b2').value.to_i #end rows of test case
    header = get_header #get testcase header

    end_colum = 66 + header.length
   
    if end_colum.to_i <= 90 # 90 is 'Z' ASCII
      testcase_array = @worksheet.Range("c2:#{end_colum.chr}#{rows}").value
    else
      first_character  = 65 + (end_colum.to_i - 90)/26 # 26 is English letters number
      second_character = 65 + (end_colum.to_i - 90)%26 - 1
      testcase_array =  @worksheet.Range("c2:#{first_character.chr}#{second_character.chr}#{rows}").value
    end
  
    testcase_array.each{|x| #get every line info
      x.each_index{|index| x[index] = x[index].to_s.gsub(/\.0*$/,'') if (x[index] =~ /[0-9]+\.[0-9]+\.[0-9]+\.[0-9]/) == nil} #deleted .0 for int class from excel value except for ip address

      number = 0
      temp_hash = Hash.new
      arg_arry = []
      act_output_arry = []
      header.each{|t| #get every header to a hash
        if t.include?("ARG")  #get all the arg to array and save as ARG hash
          arg_arry << x[number]
          temp_hash["ARG"] = arg_arry
        elsif  t.include?("Actual Output")  #get all the Actual output to array and save as Actual output hash
          act_output_arry << x[number]
          temp_hash["Actual Output"] = act_output_arry
        elsif  t.include?("Expected Output") || t.include?("Result")
          
        else
          temp_hash[t] = x[number]
        end
        number += 1
      }
      temp << temp_hash #put every step to arry
    }
   
    return temp
  end

  #
  # Read testsite form excel
  #
  # @example
  #   ex = Excel.new('d:\abc.xlsx')
  #   ex.open_xls
  #   ex.read_testsite
  #

  def read_testsite

    testsite =  @worksheet.Range('b3').value
    return testsite

  end

  #
  # Write web infomation to excel
  #
  # @example
  #   ex = Excel.new('d:\abc.xlsx')
  #   ex.open_xls
  #   ex.write_output(2,[1,2,3])
  #
  
  def write_output(step = 1,output=nil)
    row = step + 1

    number =@worksheet.Range('1:1').value.flatten.compact.rindex("Actual Output1")
    start_colum = "a"
    number.times{|t| start_colum  = start_colum.next } #Count Actual output1's colum

    output.each_index{|x|
      @worksheet.Range("#{start_colum}#{row}").value = output[x]
      start_colum  = start_colum.next
    }

  end

  #
  # Get excel header and return to array
  #
  # @example
  #   ex = Excel.new('d:\abc.xlsx')
  #   ex.open_xls
  #   ex.get_header
  #

  def get_header

    data = @worksheet.Range("1:1").value.flatten.compact.slice(2..-2)
    return data

  end

  #
  # Get launch excel index and script name
  #
  # @example
  #   ex = Excel.new('d:\abc.xlsx')
  #   ex.open_xls
  #   ex.get_running_list
  #

  def get_running_list

    rows =  @worksheet.Range("b2").value.to_i
    index = 2
    list = Hash.new

    for i in index..rows+1
      if @worksheet.Range("e#{i}").value
        list[@worksheet.Range("c#{i}").value.to_i] =  @worksheet.Range("f#{i}").value
      end
      i += 1
    end
    return list
  end

  #
  # Write result of each script to launch spreadsheet
  #  Arguments - [row, elapsed time, result]
  # @example
  #   ex = Excel.new('d:\abc.xlsx')
  #   ex.open_xls
  #   ex.write_lauch([2,'12:30','Fail'])
  #
  
  def write_lauch(arg=[])
    temp = arg
    if temp == []
      puts "Please provide 'index', 'Elapsed time' and 'result'."
    else
      @worksheet.Range("g#{temp[0]+1}:h#{temp[0]+1}").value = [temp[1],temp[2]]
    end

    if temp[2] == 'Pass'
      @worksheet.Range("h#{temp[0]+1}").Interior.ColorIndex = 10
    elsif temp[2] == 'Fail'
      @worksheet.Range("h#{temp[0]+1}").Interior.ColorIndex = 3
    end
  end


  #
  # Write card information form launch spreadsheet
  # @example
  #   ex = Excel.new('d:\abc.xlsx')
  #   ex.open_xls
  #   ex.read_card_info
  #
  def read_card_info
    info = []
    info << @worksheet.Range("b4").value
    info << @worksheet.Range("b7:b18").value
    return info.flatten
  end

  def write_card_info(info)
    row = 8
    for i in 0...info.size
      @worksheet.Range("b#{row}").value = info[i]
      row += 1
    end
  end

end

#a = Excel.new('D:\02 automation test\01 Keyword driven\01 svn\source code\Keywordscripts\unity_case_template.xls')
##a = Excel.new('D:\02 automation test\01 Keyword driven\01 svn\source code\Keywordscripts\unity_case_example_prototype.xlsx')
#
#
#a = Excel.new('E:\Launch.xls')
#a.open_xls
#p a.username
#p a.password
#a.save_as_xls('e:\ddddsds.xlsx')
#a.active_xls
#a.close_xls
#a.read_testcase
#a.write_output(2,[1,2,3])
#a.write_output(3,[4,5,6,7,8,9,10])
##a.save_xls
#a.close_xls
#sleep 100
#a.get_running_list
#a.write_lauch([3,'12.30','Fail'])
#a.write_lauch
#a.read_card_info