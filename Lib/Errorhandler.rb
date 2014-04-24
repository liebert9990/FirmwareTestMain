# The class for error handle
# There are three types of error.
# -- "". -----Card connection error.
# -- Noconnect. ----Close browser error(No connection error).
# -- Nonetwork. ----Network error.
#
$:.unshift File.dirname(__FILE__)
require 'Browser.rb'
class Errorhandler
  attr_reader :error_type, :log_info
  def initialize
    @log_info = ""  
    @error_type = ""
  end
  #handle all kinds of error according to its error type
  def errorhandle(error_str, browser_type, test_site=nil)
    @log_info = "" #clean up the string to record current error's log
    handle_result = nil
    judge_error_type(error_str)
    case @error_type
    when "Noconnect","Nonetwork"
      handle_result = connect_handle(browser_type ,test_site)
    when "Nofile"
      puts "Errorhandler: This script is not existed, start the next script."
      @log_info << "Errorhandler: This script is not existed, start the next script.\n"
      handle_result = "Next_script"
    else
      puts "Errorhandler: Other exception, next step..."
      @log_info << "Errorhandler: Other exception, next step...\n"
      handle_result = "Continue_running"
    end
    #return result: Stop_run, Next_script,Continue_running
    return handle_result
  end

  #handle the kind of network error
  #browser_type: browser' type
  #test_site: card's ip address
  def connect_handle(browser_type,test_site)
    if card_available?(1, test_site)
      if browser_opened?(1, browser_type, test_site)
        if browser_open_card?
          #card can connect, and browser can open its webpage
          puts "Errorhandler: Card can connect, and browser can open its webpage, start the next script."
          @log_info << "Errorhandler: Card can connect, and browser can open its webpage, start the next script.\n"
          return "Next_script"
        else
          #card can connect, but browser cannot open its webpage
          puts "Errorhandler: Card can connect, but browser cannot open its webpage, stop the next running."
          @log_info << "Errorhandler: Card can connect, but browser cannot open its webpage, stop the next running.\n"
          return "Stop_run"
        end
        @browser.close_browser
      else
        #browser has some problem
        puts "Errorhandler: Browser has some problem, stop the next running."
        @log_info << "Errorhandler: Browser has some problem, stop the next running.\n"
        return "Stop_run"
      end
    else
      #card refuse to connect
      puts "Errorhandler: Card refuse to connect, stop the next running."
      @log_info << "Errorhandler: Card refuse to connect, stop the next running.\n"
      return "Stop_run"
    end
  end

  #judge error type
  def judge_error_type(error_str)
 if error_str =~ /No connection could be made because the target machine actively refused it/
      @error_type = "Noconnect"
    elsif error_str =~ /chrome not reachable/
      @error_type = "Nonetwork"
    elsif error_str =~ /No such file or directory/
      @error_type = "Nofile"
    end
  end

  # ping card
  def ping_card(test_site)
    reply_from = Regexp.new('Reply from')
    results = `ping #{test_site}`
    return true if (reply_from.match(results))
    return false
  end

  # Is card available?
  def card_available?(tries,test_site)
    result = false
    result = ping_card(test_site)
    if result == false # card is still unavailable - retry
      puts "Errorhandler: Retry #{tries} times to ping the card of #{test_site}..."
      @log_info << "Errorhandler: Retry #{tries} times to ping the card of #{test_site}...\n"
      tries += 1
      if tries <= 3
        result = card_available?(tries, test_site)
      else
        puts "Errorhandler: Retry limit reached, card is still unavailable!"
        @log_info << "Errorhandler: Retry limit reached, card is still unavailable!\n"
        return result
      end
    else
      puts "Errorhandler: Retry #{tries} times to ping the card of #{test_site}..."
      puts "Errorhandler: The card of #{test_site} is available."
      @log_info << "Errorhandler: Retry #{tries} times to ping the card of #{test_site}...\n"
      @log_info << "Errorhandler: The card of #{test_site} is available.\n"
    end
    result
  end

  #Browser can be opened?
  def browser_opened?(tries, browser_type, test_site)
    result = false
    result = open_browser(browser_type, test_site)
    if result == false # browser is still unavailable - retry
      puts "Errorhandler: Retry #{tries} times to open browser..."
      @log_info << "Errorhandler: Retry #{tries} times to open browser...\n"
      tries += 1
      if tries <= 3
        result = browser_opened?(tries, browser_type, test_site)
      else
        puts "Errorhandler: Retry limit reached, browser is still unavailable!"
        @log_info << "Errorhandler: Retry limit reached, browser is still unavailable!\n"
        return result
      end
    else
      puts "Errorhandler: Retry #{tries} times to  to open browser..."
      puts "Errorhandler: Browser can be opened."
      @log_info << "Errorhandler: Retry #{tries} times to  to open browser...\n"
      @log_info << "Errorhandler: Browser can be opened.\n"
    end
    result
  end

  def open_browser(browser_type, test_site)
    browser_result = true
    #open a new browser, make sure the browser can be opend.
    begin
      @browser = Browser.new(browser_type)
      @browser.open_browser(test_site)
    rescue
      browser_result = false
    ensure
    end
    return browser_result
  end

  #this method's pre-condition: card is available, and browser has opened
  #judge browser can open the card webpage or not.
  def browser_open_card?
    # judge the webpage's address field is currect or not
    browser_url = @browser.url.to_s
    if browser_url.include?("default.html?devId=4")
      card_result = true
    else
      card_result = false
    end
    return card_result
  end
end

#str = "http://test_site/default.html?devId=4"
#str_new = str.to_s
#if str_new.include?("default.html?devId=4")
#  puts "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!"
#else
#  puts "&&&&&&&&&&&&&&"
#end
#errorhandle = Errorhandler.new
#error_str_noconnection = "No connection could be made because the target machine actively refused it"
#error_str_nonetwork = "#<Errno::ENOENT: No such file or directory - D:\\private\\keyword_unity\\Keywordscripts\\710.040System\\System_710.040.10.110.xls>"
#errorhandle.judge_error_type(error_str_nonetwork)
#puts  errorhandle.error_type
#error_str_noconnection = "No connection could be made because the target machine actively refused it"
#error_str_nonetwork = "chrome not reachable"
#browser_type = "chrome"
#ip_right = "10.146.64.43"
#ip_wrong = "10.146.64.55"
#command_noconnection = errorhandle.errorhandle(error_str_noconnection, browser_type, ip_right)
#puts "++++++++++++++++++++++++++"
#puts "Next command(should be Next_script): ----------------#{command_noconnection}---------------"
#
#command_nonetwork = errorhandle.errorhandle(error_str_nonetwork, browser_type, ip_right)
#puts "++++++++++++++++++++++++++"
#puts "Next command(should be Stop_run): ----------------#{command_nonetwork}---------------"
