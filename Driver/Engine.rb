#The detail process for one test case.
#
#
#
#
require 'fileutils'
$:.unshift File.dirname(__FILE__).sub('Driver','Lib')
require 'Excel.rb'
require 'Browser.rb'
require 'Keyword.rb'
require 'Errorhandler.rb'
class Engine
  attr_reader :log_info, :case_result, :elapsed_time, :run_command, :card_info

  def initialize()
    @run_command = "Next_script"
    @elapsed_time = 0
    @case_result = "Fail"
    @log_info = ""
    @card_info = Array.new
    $username = "Liebert"
    $password = "Liebert"
  end

  # select which browser instance to create
  # 1 - ie
  # 2 - firefox
  # 3 - chrome
  # pick a number to select browser
  def select_browser
    bro_arr = ['ie','firefox','chrome']
    while (1)
      for i in 1...bro_arr.length + 1
        puts i.to_s + ' - ' + bro_arr[i-1]
      end
      puts "pick a number to select browser - "
      ch = gets
      if ch.to_i>=1&&ch.to_i<=bro_arr.length
        puts "running on #{bro_arr[ch.to_i-1]}"
        return bro_arr[ch.to_i-1]
      else
        puts "please select a correct number."
      end
    end
  end

  #execute one piece of test action
  def execute_test_action(one_case)
    component = one_case["Component"]
    action = one_case["Action"]
    arguments = one_case["ARG"]
    result = one_case["Actual Output"]

    action = action.downcase
    #get element information
    if action =~ /error_message|get_tip|get_status|get_description/
      element = component
    else
      element = @browser.get_element(component)
    end

    #call the method for this action
    @keyword.method("#{action}").call(element,arguments,result)
  end

  #execute one piece of function action
  def execute_function_action(one_case)
    action = one_case["Action"]
    arguments = one_case["ARG"]
    result = one_case["Actual Output"]

    #call the method for this action
    action = action.downcase
    if action =~ /login/
      if arguments == "" || arguments == nil # no ARG column in testcase
        usr = $username
        psw = $password
      else
        usr = arguments[0]
        psw = arguments[1]
        usr = $username if arguments[0] == "" || arguments[0] == nil
        psw = $password if arguments[1] == "" || arguments[1] == nil
      end
      actual_result = @keyword.method("#{action}").call(usr, psw,result)
      @keyword.login_flag = true
    elsif action =~ /popup|reboot/
      actual_result = @keyword.method("#{action}").call(arguments, result)
    else
      use_info = [$username, $password]
      actual_result = @keyword.method("#{action}").call(use_info, arguments, result)
    end
    return actual_result
  end

  #execute one piece of command action
  def execute_cmd_action(one_case)
    action = one_case["PageName"]
    result = one_case["Actual Output"]
    arguments = one_case["ARG"]
    action = action.downcase
    if arguments == "" || arguments == nil # no ARG column in testcase
      arg = nil
    else
      arg = arguments[0]
    end
    @keyword.method("#{action}").call(arg, result)
  end

  # main process to execute one test case
  def run(browser_type, excel_file, first_testcase)
    start_time = Time.now
    @log_info = "" #clean up the string to record current script's log

    #initialize the class, create the instance for each class
    @excel = Excel.new(excel_file)
    @browser = Browser.new(browser_type)
    @keyword = Keyword.new(@browser)
    @errorhandler = Errorhandler.new

    #1. Open the input file.
    begin
      @excel.open_xls
      #2. read user information from new excel
      ip_address = @excel.read_testsite
      $username = @excel.username
      $password = @excel.password
      #3. open ie
      @browser.open_browser(ip_address)
      #4. read all test cases into memory
      test_cases = @excel.read_testcase
      next_step = true # use to judge run the next step or stop running.

      #5. take one piece of test case from the array
      test_cases.each do |one_case|
        cmd_type = one_case["Type"]
        #if this row is a comment, continue to execute next row
        if cmd_type.to_s =~ /^\;/
          @log_info<<one_case.to_s<<"\n"
          puts one_case
          next
        end
        next_step = run_one_step(one_case,browser_type,ip_address)
        break if next_step == false
      end

      #First running testcase, write the card information into launch
      if first_testcase
        @card_info = @keyword.get_card_info
      end

    rescue Exception => e     
      @log_info<<"Engine: ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~Jump to expection"<<"\n"
      @log_info<< "Engine: Executing failed: #{e}"<<"\n"
      #@log_info<<$@.to_s<<"\n"
      puts "Engine: ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~Jump to expection"
      puts "Engine: Executing failed: #{e}"
      #puts $@.to_s
        
      #error handle
      error_str = e.to_s
      @run_command = @errorhandler.errorhandle(error_str, browser_type, ip_address)
      @log_info<< @errorhandler.log_info
    ensure
      # save the result and clean the enviroment
      
      if @browser.browser_entry != nil 
        @browser.close_browser
      else
        @log_info<<"Engine: ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~Browser has some problem, stop the next running."<<"\n"
        puts "Engine: ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~Browser has some problem, stop the next running."
        @run_command = "Stop_run"
      end
      @excel.write_state("Error") if !next_step #if script occus error, wirte the error state into excel
      @excel.save_xls
      @case_result = @excel.pass_fail
      @excel.close_xls
      elapsed_time = Time.now - start_time
      @elapsed_time = elapsed_time
    end
  end

  def run_one_step(one_case,browser_type,ip_address)
    cmd_type = one_case["Type"]
    step_index = one_case["Test Steps"]
    next_step = true
    #7. execute each row according to its type
    begin
      case cmd_type.to_s
      when "F"
        action_result = execute_function_action(one_case)
      when "C"
        action_result = execute_cmd_action(one_case)
      when "T"
        #8. jump to execute the method of execute_test_action
        action_result = execute_test_action(one_case)
      else

      end
    rescue Exception => e
      action_result = one_case["Actual Output"]
      #set every actual output value as false
      (0..action_result.size-1).each{|i| action_result[i] = false}
      @log_info<<"Engine: ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~Jump to expection"<<"\n"
      @log_info<< "Engine: Executing failed: #{e}"<<"\n"
      #@log_info<<$@.to_s<<"\n"
      puts "Engine: ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~Jump to expection"
      puts "Engine: Executing failed: #{e}"
      #puts $@.to_s

      #error handle
      error_str = e.to_s
      @run_command = @errorhandler.errorhandle(error_str, browser_type, ip_address)
      @log_info << @errorhandler.log_info
      next_step = false if @run_command == "Next_script" || @run_command == "Stop_run"
    ensure
      # write the executed result for one step
      @excel.write_output(step_index.to_i, action_result)
      #wirte the "Actual Output" by using its actual result.
      case_result = ["Test Steps"=>one_case["Test Steps"], "Type"=>one_case["Type"], "PageName"=>one_case["PageName"], "Component"=>one_case["Component"], "Action"=>one_case["Action"], "ARG"=>one_case["ARG"], "Actual Output"=>action_result]
      @log_info<<case_result.to_s<<"\n"
      puts case_result
    end
    next_step
  end

end

#excel_file = (File.dirname(__FILE__).sub('Driver','Keywordscripts')<<'\\'<<'unity_case_example_prototype.xlsx').gsub('/','\\')
#excel_file = (File.dirname(__FILE__).sub('Driver','Keywordscripts')<<'\\'<<'unity_case_snmpv1_access.xlsx').gsub('/','\\')
#excel_file = (File.dirname(__FILE__).sub('Driver','Keywordscripts')<<'\\'<<'unity_case_snmpv1_trap.xlsx').gsub('/','\\')
#excel_file = (File.dirname(__FILE__).sub('Driver','Keywordscripts')<<'\\'<<'unity_case_snmpv3.xlsx').gsub('/','\\')
#excel_file = (File.dirname(__FILE__).sub('Driver','Keywordscripts')<<'\\'<<'unity_case_firmware_update.xlsx').gsub('/','\\')
#result_path = File.dirname(__FILE__).sub('Driver','Report') + "/Testcase Result"
#engine = Engine.new(excel_file, result_path)
#engine.run
#puts engine.case_result
#puts engine.elapsed_time
