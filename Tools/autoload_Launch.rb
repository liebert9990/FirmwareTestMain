require 'win32ole'

#  - createand return new instance of excel
def new_xls(s_s) #wb name and sheet number
  ss = WIN32OLE::new('excel.Application')
  wb = ss.Workbooks.Open(s_s)
  ws = wb.Worksheets(1)
  ss.visible = true # For debug
  xls = [ss,wb,ws]
end

#Search a target file in the given path, return the index of the array of the filenames if exist, otherwise return -1
def file_search(path)
  if path =~ /.*\/Keywordscripts\/\z/
    fl_list = Dir.entries(path).delete_if{ |e| e=~ /^\..*/ || e=~ /^.*\.xls/ ||e=='temp'}
  else
    fl_list = Dir.entries(path).delete_if{ |e| e=~ /^\..*/ }
  end
end

#get test case folder list for Launch file
def get_testcaselist
  keywordscripts = (__FILE__).gsub(/Tools.*/,'Keywordscripts/')
  addr = []
  testsuit_path = file_search(keywordscripts)
  testsuit_path.each{|x|
    #"D:/02 automation test/01 Keyword driven/01 svn/source code/Keywordscripts/"
    temp = keywordscripts + x + "/"
    #"D:/02 automation test/01 Keyword driven/01 svn/source code/Keywordscripts/710.050 Time Service Settings/"
    temp2 = file_search(temp)
    temp2.each{|s|
      temp3 = temp + s
      temp4 = temp3.gsub(/.*\/Keywordscripts\//,'').gsub(/\//,'\\')
      addr << temp4
    }
  }
  return addr
 
end

def write_launch(list=nil)
  lauch_path =  (__FILE__).gsub(/Tools.*/,'') + "Launch.xlsx"
  p lauch_path

  @ss = new_xls(lauch_path)
  @wb = @ss[1]
  @ws = @ss[2]

  row = 2
  list.each_index {|x| 
    @ws.Range("F#{row+x}").Value =list[x]
  }
  
  #Count for test cases
  @ws.Range("B2").Value =list.size

  @wb.save
end

def select_script(num)
  rows = @ws.Range('B2').value.to_i
  puts rows
  rows.times{|x|
    if num == '1'
      @ws.Range("E#{2+x}").Value = 'TRUE'
    elsif num == '2'
      @ws.Range("E#{2+x}").Value = 'FALSE'
    end
  }
  @wb.save
end

begin
  #script running
  p "Start Running"
  # get test case list
  tst = get_testcaselist
  #write list to Lauch.xlsx
  write_launch(tst)
  #select run script, support all select and not select all.
  print "\nPlease select the number, if you want run all the script or not, \n1.all    2.:no script run \n "
  run_number = gets.chomp
  p run_number
  select_script(run_number)
  p "End Running"
end