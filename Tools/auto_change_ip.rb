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
      p temp3
      addr << temp3
    }
  }
  return addr
end

def change_ip(list=nil,ip=nil)
  if ip == "" then ip = "test_site" end
  
  if list
    list.each{|x|
      p x
      ss = new_xls(x)
      wb = ss[1]
      ws = ss[2]
      ws.Range("B3").Value = ip
      wb.save
      ss[0].Workbooks.close
      ss[0].quit
    }
  end
end

begin
  #script running
  p "Start Running"
  # get test case list
  list = get_testcaselist
  #get ip
  print "\nPlease type card ip by <Enter>: "
  card_ip = gets.chomp
  #p card_ip
  #Change every test case ip
  change_ip(list,card_ip)
  p "End Running"
end