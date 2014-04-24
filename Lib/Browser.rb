require 'watir-webdriver'
require 'win32ole'
require 'win32API'
Watir.always_locate = false #http://watirmelon.com/2011/08/14/telling-watir-webdriver-to-always-locate-elements-upon-use/
class Browser

  attr_accessor :brow_type

  def initialize(type)
    @brow_type = type
    $accuelems = Hash.new # repository of already-found elements
    $tabframe = 'tabArea'
    $navframe = 'navigationFrame'
    $detframe = 'detailArea'
    $devframe = 'deviceStatus'
    $text_col = 1 # colum index of the label name, in unity card, it is 1.
    $father_link = '' # indicate which page the button or fields belong to
    $login_times = 1 # record the login times (for ie second+ login).
  end

  # get the browser version
  # eg. browser.version => 9.0.8112.16421|25.0.1364.97|18.0
  def version
    case @brow_type
    when 'internet_explorer'
      temp =  `reg query \"HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\Internet Explorer\" \/v svcVersion"`
    when 'firefox'
      temp =  `reg query \"HKEY_LOCAL_MACHINE\\SOFTWARE\\Wow6432Node\\Mozilla\\Mozilla Firefox\" \/v CurrentVersion"`
    when 'chrome'
      temp =  `reg query \"HKEY_LOCAL_MACHINE\\SOFTWARE\\Wow6432Node\\Google\\Update\\ClientState\\{8A69D345-D564-463C-AFF1-A69D9E530F96}\" \/v pv"`
    end
    @version = /(\d+\.)+\d+/.match(temp)
  end

  # get the browser language
  # eg. browser.language => en-US
  def language
    case @brow_type
    when 'internet_explorer'
      temp =  `reg query \"HKEY_LOCAL_MACHINE\\SYSTEM\\CurrentControlSet\\Control\\MUI\\UILanguages" \/s"`
    when 'firefox'
      temp =  `reg query \"HKEY_LOCAL_MACHINE\\SOFTWARE\\Wow6432Node\\Mozilla\\Mozilla Firefox\" \/v CurrentVersion"`
    when 'chrome'
      temp =  `reg query \"HKEY_LOCAL_MACHINE\\SOFTWARE\\Wow6432Node\\Google\\Update\\ClientState\\{8A69D345-D564-463C-AFF1-A69D9E530F96}\" \/v lang"`
    end
    @language = /zh-CN|en-US/.match(temp)
  end

  # get the browser url
  # eg.
  #   browser.goto "http://www.google.com"
  #   browser.url
  #   #=> "http://www.google.com"
  def url
    @url = $brow.url
  end

  def browser_entry
    $brow
  end

  # create an browser instance
  def open_browser(site)
    $brow = Watir::Browser.new @brow_type
    $brow.goto(site)
    $brow.driver.manage.window.maximize # maximize the browser
  end

  # close the browser
  def close_browser
    $brow.close
  end

  # get the element based on the label
  # eg. browser.get_element('System Name')
  #  => element object
  def get_element(label) # label comes from excel
    flag = false
    if $accuelems.has_key?(label) #check elements repo first
      flag = true
      is_what,type,elem_id = $accuelems[label]
    elsif $accuelems.has_key?($father_link+label) # bind the field element with its page name to be specific.
      flag = true
      is_what,type,elem_id = $accuelems[$father_link+label]
    else
      while 1
        if $brow.frame(:id, $navframe).link(:text, label).exist? then is_what = 'link'; break end # link should be ahead of field. ref - snmpv3 users setting (20) page
        if $brow.frame(:id, $detframe).td(:text, label).exist? then is_what = 'field';break end
        if $brow.frame(:id, $detframe).button(:value, label).exist? then is_what = 'button'; break end
        if $brow.frame(:id, $tabframe).link(:text, label).exist? then is_what = 'tab'; break end
        if $brow.form(:name, 'firmwareHttpForm').exist? then is_what = 'file_upload'; break end
      end
    end
    
    case is_what
    when 'tab'
      $accuelems[label] = ['tab'] unless flag
      get_tabs(label)
    when 'link'
      $father_link = label # remember the father link, because in sms and email, mstp and ethernet page, some elements has the same label name with different id
      $accuelems[label] = ['link'] unless flag
      get_links(label)    
    when 'button'
      $accuelems[label] = ['button']unless flag
      get_buttons(label)
    when 'file_upload'
      $accuelems[label] = ['file_upload']unless flag
      get_file_upload(label)
    when 'field'
      if flag
        get_fields(label,type,elem_id)
      else
        fied = get_fields(label)
        $accuelems[$father_link+label] = ['field',fied.type.to_s,fied.id.to_s]
        fied
      end
    end
  end

  # get the fields element  (textbox, checkbox, select list, file_field) by a cell's label
  # browser.get_fields('System Name')
  #  => textbox object (behind this label cell)
  def get_fields(label,type = nil,id = nil)
    case type.to_s
    when 'checkbox'
      field = $brow.frame(:id, $detframe).checkbox(:id, id)
    when 'textarea','text', 'password'# textarea , text, or password
      field = $brow.frame(:id, $detframe).text_field(:id, id)
    when 'select-one'
      field = $brow.frame(:id, $detframe).select_list(:id, id)
    when ''
      if $brow.frame(:id, $detframe).div(:id, "commandTable").table().exist?
        tabl,row_cout = get_table("commandTable")
      else
        tabl,row_cout = get_table
      end
      for i in 0...row_cout
        if tabl.row(:index, i)[$text_col].text == label
          if tabl.row(:index, i)[$text_col + 1].span().exist? # checkbox in a span
            field = tabl.row(:index, i)[$text_col + 1].span().element.to_subtype # to_subtype - focus to this specific elements
          else # other fields is not in a span
            field = tabl.row(:index, i)[$text_col + 1].element.to_subtype # to_subtype - focus to this specific elements
          end
        end # end if
      end # end for
    end # end case
    Watir::Wait.until{field.exist?}
    #   Watir::Wait.until{field.enabled?} undefined method 'enabled? for watir::testarea'
    field
  end

  #get the status in status table by cell's label and return element value
  #browser.get_status('System status')
  def get_status(label)
    tabl,rows_cout = get_table("statusTable")
    for i in 0...rows_cout
      if tabl.row(:index, i)[$text_col].text == label
        status_element = tabl.row(:index, i)[$text_col + 1].text
        break
      end
    end
    return status_element
  end


  # get the firmware update page elements
  def get_file_upload(label)
    ele = $brow.form(:name, 'firmwareHttpForm').button(:value, label)
    if ele.exist?
      ele
    else
      $brow.form(:name, 'firmwareHttpForm').file_field(:name, 'Firmware File Upload')
    end
  end

  # get the buttons by name (Edit, Save, Cancel,Enable, Restart, Reset to Factory Defaults, Web)
  # eg. browser.get_buttons('Edit')
  #     => Edit button object
  def get_buttons(name)
    if name == 'Cancel'
      name = name.downcase
      if $brow.frame(:id, $detframe).button(:id, "#{name}Button").exist?
        #General cancel button
        btn = $brow.frame(:id, $detframe).button(:id, "#{name}Button")
      else
        #Support page
        btn = $brow.frame(:id, $detframe).button(:id, "enableComms#{name}")
      end
    else
      #General button
      btn = $brow.frame(:id, $detframe).button(:value, name)
    end
    
    Watir::Wait.until{btn.exist?}  #for firefox
    Watir::Wait.until{btn.enabled?}  #btn.wait_until_present # wait_until_present = visible + enable ? No, firefox fail
    btn
  end

  # get the links by name (Unity Configuration, System...)
  # eg. browser.get_links('Configuration')
  #     => Configuration link object
  def get_links(name)
    link = $brow.frame(:id, $navframe).link(:text, name)
    link.wait_until_present
    link
  end

  # get the tabs by name (Unity,Device)
  # eg. browser.get_tabs('Unity')
  #     => Unity tab object
  def get_tabs(name)
    tab = $brow.frame(:id, $tabframe).link(:text, name)
    tab.wait_until_present
    tab
  end
  
  # get the table of a page - statusTable, eventTable, commandTable and settingTable
  # eg. browser.get_table
  # => Settings table of current page and row number
  # eg. browser.get_table('statusTable') and row number
  # => Status table of current page (see - Configuration-> Emerson Management Protocol -> MSTP page)
  def get_table(type = 'settingTable')
    tabl = $brow.frame(:id, $detframe).div(:id, type).table()
    Watir::Wait.until{tabl.exist?} # for firefox
    row_cout = tabl.rows.size
    return tabl, row_cout
  end

  # return a hash which contains all the field elements of current page. Will be used in SNMP user add.
  # eg. browser.get_muti_elements
  # => {"System Name"=> textbox object, "Contact Information"=> textbox object, ..."Display Temperature Units"=> select list object}
  def get_muti_elements
    mutielems = Hash.new
    tabl,rows = get_table
    for i in 1...rows
      key = tabl.row(:index, i)[$text_col].text
      if tabl.row(:index, i)[$text_col + 1].span().exist? # checkbox in a span
        value = tabl.row(:index, i)[$text_col + 1].span().element.to_subtype # to_subtype - focus to this specific elements
      else # other fields is not in a span
        value = tabl.row(:index, i)[$text_col + 1].element.to_subtype # to_subtype - focus to this specific elements
      end
      mutielems[key] = value
      $accuelems[key] = ['field',value.type.to_s,value.id.to_s]
    end
    mutielems
  end

  # get the tip text of a item by its label
  # browser.get_tip('System Name')
  #  => 'End user assigned name for the system'
  def get_tip(label) # label comes from excel
    ['settingTable','statusTable','commandTable'].each{|nam|
      if $brow.frame(:id, $detframe).div(:id, "#{nam}").table().exist?
        tabl,rows_cout = get_table("#{nam}")
        for i in 0...rows_cout
          if tabl.row(:index, i)[1].text == label
            tip_img = tabl.row(:index, i)[0].element.to_subtype
            tip_img.fire_event 'onmouseover'
            tip_text = $brow.frame(:id, $detframe).div(:id, 'tt').div(:id, 'ttcont').text.to_s
            tip_img.fire_event 'onmouseout'
            return tip_text
          end
        end
      end
      break if tip_text
    }
  end

  #get the invalid text behind the element when input invalid characters
  #take the element label as the input
  def get_invalid_text(label)
    invlid_text = ''
    tabl,row_cout = get_table
    for i in 0...row_cout
      if tabl.row(:index, i)[$text_col].text == label
        invlid_text = tabl.row(:index, i)[$text_col + 3].text # to_subtype - focus to this specific elements
        break
      end # end if
    end # end for
    invlid_text
  end

  # delete the cache of a browser
  def delete_cache
    #TODO - a script in Tool folder does this
  end

  # Login for Charlie
  # username, password come from excel
  def charlie_login(user, pswd) 
    sleep 1 #This sleep is critical, may need to be adjusted
    autoit = WIN32OLE.new('AutoITX3.Control')
    if @brow_type == 'firefox'
      autoit.WinWait('Authentication Required')
    end
    #If username element has a value, clear up it.
    if$brow.frame(:id, $detframe).text_field(:id, 'inputUsr').value != ""
      $brow.frame(:id, $detframe).text_field(:id, 'inputUsr').set("")
    end
#    puts user
    autoit.Send(user)
    autoit.Send('{TAB}')
    autoit.Send(pswd)
    autoit.Send('{ENTER}')
    puts " Login completed"
  end

  # General login, include Charlie and Delta.
  # username, password come from excel
  def login(user, pswd)
    if $brow.alert.exists?
      login_text = $brow.alert.text
      # Delta's Login
      if login_text =~ /Do you want to login now?/
        popup('OK')
        $brow.form(:name, 'loginForm').text_field(:id, 'username').set(user)
        $brow.form(:name, 'loginForm').text_field(:id, 'password').set(pswd)
        $brow.form(:name, 'loginForm').submit
      else #Charlie's login: login windows is a alert on ie and firefox
        charlie_login(user, pswd)
      end
    else #Charlie's login: login windows is not a alert on Chrome
      charlie_login(user, pswd)
    end
  end

  # check if a 'restart required' popup will be appear
  # based on the text in device status panel
  # return true if will be appear, false if not.
  def popup_status
    status = $brow.frame(:id, $devframe).table().when_present.to_a.to_s
    return status.include?('Normal Operation')
  end

  # Wait, handle the alert and confirm popup and return the content text
  # eg.
  # save.click
  # browser.alert_popup
  # => click 'OK' button and return the content text
  #
  # reset.click
  # browser.alert_popup('Cancel')
  # => click 'Cancel' button and return the content text
  def popup(btn = 'OK')
    Watir::Wait.until{$brow.alert.exists?} # 30 seconds in maxmum
    popup_text = $brow.alert.text
    if btn == 'OK'
      $brow.alert.ok
    elsif btn == 'Cancel'
      $brow.alert.close
    else
      puts "#{btn} button not found in popup window"
    end
    return popup_text
  end


  # Refresh page by this method
  def page_refresh
    autoit = WIN32OLE.new('AutoITX3.Control')
    autoit.Send('{F5}')
  end

  #For firmware update page to get the description
  #browser.get_description('Update Firmware')
  # browser.get_tip('System Name')
  #  return => 'Click this button to initiate the firmware update.'
  def get_description(label)
    tabl = $brow.div(:id, 'DetailPanelAreaFwUpdate').table()
    rows_cout = tabl.rows.size
    for i in 0...rows_cout
      if tabl.row(:index, i)[0].text == label
        desc = tabl.row(:index, i)[1].text
        break
      end
    end
    return desc
  end
end

# --------------------test scripts-----------------

#  browser = Browser.new('chrome')
#  browser.open_browser('10.146.64.60')
#  #puts browser.version
#  #puts browser.language
#  browser.get_element('Unity').click
#  #browser.get_element('Configuration').click
#  browser.get_element('Protocols').click
#  browser.get_buttons('Edit').click
#  browser.login('Liebert', 'Liebert')
#browser.get_element('SNMP Protocol').set()
#browser.get_element('BACnet Protocol').set()
#puts "-------------------"
##  browser.get_buttons('Cancel').click
#brow = browser.browser_entry
#brow.frame(:id, 'detailArea').div(:id, 'settingTable').table().row(:index,3)[2].span().element.to_subtype.set()
#puts "-------------------"
#brow.frame(:id, 'detailArea').button(:value, 'Cancel').click()
##brow.frame(:id, 'detailArea').button(:id, 'cancelButton').click()
#puts "+++++++++++++++++++"
#browser.popup
#puts "+++++++++++++++++++popup"
#browser.close_browser

=begin
  puts Time.now.strftime("%m-%d_%H-%M-%S")
  browser = Browser.new('chrome')
  browser.open_browser('10.146.64.66')
  #puts browser.version
  #puts browser.language
  browser.get_element('Unity').click
  browser.get_element('Configuration').click
  browser.get_element('LIFE (TM').click
  browser.get_element('Advanced').click
  #browser.get_element('Support').click
  browser.get_buttons('Edit').click
  browser.login('Liebert', 'Liebert')
  browser.get_element('Service Mode').click
 # ss=  browser.get_description('Update Firmware')
  #p ss
  #  browser.get_element('Domain Name Suffix List')
  #  #browser.firmware_refresh
  #  puts "refresh sucess"

 
#  browser.get_element('Support').click
#  ss= browser.get_tip('Reset Card to Factory Defaults')
#  p ss
  #  #browser.get_element('Modbus').click
  #  #a = browser.get_table('statusTable')   #############
  #  #p a
  # p browser.get_tip('Restart Card')
#  browser.get_status('Agent Date and Time')

  #  browser.get_element('SNMP').click
  #  browser.get_element('SNMPv3 Users Setting (20)').click
  #  i = 1
  #  while i<=20
  #    puts "on SNMPv3 Users Setting [#{i}]"
  #    browser.get_element("SNMPv3 Users Setting [#{i}]").click
  #    puts browser.get_tip('SNMPv3 User Enable')
  #    puts browser.get_tip('SNMPv3 Username')
  #    puts browser.get_tip('SNMPv3 Access Type')
  #    puts browser.get_tip('SNMPv3 Authentication')
  #    puts browser.get_tip('SNMPv3 Authentication Secret')
  #    puts browser.get_tip('SNMPv3 Privacy')
  #    puts browser.get_tip('SNMPv3 Privacy Secret')
  #    puts browser.get_tip('SNMPv3 Trap Target Addresses')
  #    puts browser.get_tip('SNMPv3 Trap Port')
  #    i += 1
  #  end
  #  # see the elements repository
#  puts $accuelems
#  browser.close_browser
#  puts Time.now.strftime("%m-%d_%H-%M-%S")
=end