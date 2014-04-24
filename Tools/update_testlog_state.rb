require 'win32ole'
require "rexml/document"
include REXML

  #write the result into testlog.
  #arg have two elements: 1-case name, 2-result
  def write_testlog(arg=nil)
    test_result = arg[1]
    file_name = arg[0].sub('.xls','.tlg')
    testlog_path = "I:/lmg test engineering/TestLog/Projects/157182 - IS-Unity_Charlie/Project Test Cases/IPv4/IE8/710. Configure"
    testlog_file = (testlog_path<<"/"<<file_name).gsub('/','\\')

    #puts testlog_file
    
    doc = Document.new(File.open(testlog_file))

    doc.elements.each("project_test_case"){  |elem|
          elem.elements["test_result"].add_text( "#{test_result}")        
    }
    file_target = File.open(testlog_file, "w")
    file_target.write(doc)
    file_target.close
  end 
    
  #  - createand return new instance of excel
  def new_xls(s_s) #wb name and sheet number
    ss = WIN32OLE::new('excel.Application')
    wb = ss.Workbooks.Open(s_s)
    ws = wb.Worksheets(1)
    ss.visible = true # For debug
    xls = [ss,wb,ws]
  end
  
  # Get launch excel index and script name
  #
  # @example
  #   ex = Excel.new('d:\abc.xlsx')
  #   ex.open_xls
  #   ex.get_running_list
  #
  def get_running_list(worksheet)
    rows =  worksheet.Range("b2").value.to_i
    index = 2
    testcases = Hash.new

    for i in index..rows+1
      if worksheet.Range("e#{i}").value
        testcases["#{worksheet.Range("f#{i}").value}"] =  worksheet.Range("h#{i}").value
      end
      i += 1
    end
    return testcases
  end
  

begin
  #script running
  p "Start Running"
  
  s_s =File.dirname(__FILE__) + "/Launch.xlsx"
  spreadsheet = new_xls(s_s)
  worksheet = spreadsheet[2]
  # get test case list
  running_list = get_running_list(worksheet)

  #write result into testlog
  p "Write the result into testlog..."
  
  running_list.each do |key,value|
    arg = [key,value]
    write_testlog(arg)
  end
  
  p "End Running"
end