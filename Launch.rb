# To change this template, choose Tools | Templates
# and open the template in the editor.


$:.unshift File.dirname(__FILE__)<<"/Lib"
require 'Excel.rb'
require 'Report.rb'
require 'Errorhandler.rb'
$:.unshift File.dirname(__FILE__)<<"/Driver"
require 'Engine.rb'

begin
  start_time = Time.now
  launch_file = __FILE__.gsub(".rb",".xlsx")
  exe_result = "failed"
  card_info = Array.new
  
  #Create launch report folder
  launch_report_folder = (File.dirname(__FILE__)<<"/Report/Launch Report/"<<start_time.strftime("%m-%d_%H-%M-%S")).gsub('/','\\')
  Dir.mkdir File.dirname(launch_report_folder) unless File.exist?File.dirname(launch_report_folder)
  Dir.mkdir launch_report_folder unless File.exist? launch_report_folder

  #copy the report generator to new launch folder
  report_generator = File.dirname(__FILE__)<<"/Lib/TestReportGenerator.xslt"
  if File.exists?(report_generator)
    FileUtils.copy report_generator,launch_report_folder
  end

  #Create the testcase result folder
  case_result_folder = (File.dirname(__FILE__)<<"/Report/Testcase Result/"<<start_time.strftime("%m-%d_%H-%M-%S")).gsub('/','\\')
  Dir.mkdir File.dirname(case_result_folder) unless File.exist?File.dirname(case_result_folder)
  Dir.mkdir case_result_folder unless File.exist? case_result_folder

  #Create the log report folder
  log_link_folder = (File.dirname(__FILE__)<<"/Report/Log Recording/"<<start_time.strftime("%m-%d_%H-%M-%S")).gsub('/','\\')
  Dir.mkdir File.dirname(log_link_folder) unless File.exist?File.dirname(log_link_folder)
  Dir.mkdir log_link_folder unless File.exist? log_link_folder
  #the spreadsheet path
  testcase_path = File.dirname(__FILE__)<<"/Keywordscripts"

  #initialize the engine class
  engine = Engine.new
  browser_type = engine.select_browser
  errorhandler = Errorhandler.new

  #Save the source luanch into a new file, which can be used to record the test result.
  FileUtils.copy launch_file, launch_report_folder
  new_launch = (launch_report_folder+ "/" + File.basename(launch_file)).gsub('/','\\')

  #Open the launch file, read the selected testcases
  @excel = Excel.new(new_launch)
  @excel.open_xls
  @excel.write_starttime(start_time)
  running_list = @excel.get_running_list
  
  if running_list.size == 0
    puts "There is no selected spreadsheet, stop to run."
  end

  #initialize the report class
  report = Report.new
  launch_xml =  (launch_report_folder + "/" + File.basename(launch_file).gsub('.xlsx', '.xml')).gsub('/','\\')
  report.creat_launch_xml(launch_xml)
  report.create_xml_nodes(running_list)

  #Falg for first running testcase
  first_testcase = true

  #execute the selected testcases one by one
  running_list.each do |running_case|
    index = running_case[0]
    file_name = running_case[1]

    time_stamp = start_time.strftime("%m-%d_%H-%M-%S")

    begin
      #Create the log file and write the information into log file.
      log_file = (log_link_folder + "/"  + File.basename(file_name).chomp(".xls")+'-'+time_stamp+(".txt")).gsub('/','\\')
      log_report = report.create_log(log_file)

      #Save the source excel into a new file, which can be used to record the test result.
      excel_file = (testcase_path + "/" + file_name).gsub('/','\\')
      report.write_log(log_report, "Starting to Execute:   " + excel_file)
      puts "Starting to Execute:   " + excel_file
      FileUtils.copy excel_file, case_result_folder
      Dir.chdir(case_result_folder)

      excel_copy = (case_result_folder+ '/' + File.basename(excel_file)).gsub('/','\\')
      new_excel = (case_result_folder+ '/' + File.basename(excel_file).chomp(".xls")+'-'+time_stamp+(".xls")).gsub('/','\\')
      File.rename(excel_copy, new_excel)
  
      #The process to execute one testcase
      engine.run(browser_type, new_excel, first_testcase)
      if first_testcase
        card_info = engine.card_info
        first_testcase = false
      end

      #The argument of testcase executed result for Log
      if engine.case_result == "Pass"
        exe_result = "successed"
      else
        exe_result = "failed"
      end
      run_command = engine.run_command
    rescue Exception => e
      puts "Launch: Executing failed: #{e}\n"
      report.write_log(log_report, "Launch: Executing failed: #{e}\n")
      run_command = errorhandler.errorhandle($!.to_s, browser_type)
      report.write_log(log_report, errorhandler.log_info)
      exe_result = "failed"
    ensure
      #Write the log information into Log file.
      report.write_log(log_report, engine.log_info)
      report.write_log(log_report, "-----------------Testcase running #{exe_result}.------------------------\n")
      puts "-----------------Testcase running #{exe_result}.------------------------\n\n"

      #write the executed result into launch.xsl
      exe_result = [index, engine.elapsed_time, engine.case_result]
      @excel.write_lauch(exe_result)
      @excel.write_card_info(card_info)
      #write the result into xml file
      spread_sheet_link = (File.basename(file_name).chomp(".xls")+'-'+time_stamp+(".xls")).gsub('/','\\')
      case_status = [index, engine.elapsed_time,spread_sheet_link,File.basename(log_file),engine.case_result]
      #report.creat_xml_node(case_status)
      report.edit_xml_node(case_status)
      card_info << browser_type
      report.write_card_info(start_time, card_info)
      
      #write the result into testlog.
      #arg_testlog = [running_case,engine.case_result]
      #report.write_testlog(arg_testlog)
      
      #Script will execute next script? "Next_script" -- next csript. "Stop_run" -- stop all script running.
      case run_command
      when "Next_script"
        next
      when "Stop_run"
        break
      end
    end
  end

rescue Exception => e
  puts "Launch: Executing failed: #{e}\n\n"
  #puts $@.to_s
ensure
  # Close the active test case before exit.
  @excel.save_xls
  @excel.close_xls
end
