# The class for keyword defination
# There are three types for all keyword: "T", "F", "C".
# -- "T" is used for all generic component: click, select_combobox, set_textbox, set_checkbox,set_filefield, read_result.
# -- "F" is used for user-defined interface: login_check, popup_restart, popup_reset, add_snmp.
# -- "C" is used for command: waitsave.
#
class Keyword

  def initialize(browser)
    @browser = browser
    @login_flag = false
  end

  attr_writer :login_flag

  # action for the component of button/link
  # element: component object for button/link
  # argument: just for the reflect for "T" type action, it is unuseful in here.
  # result: result for the process, it is true as a default value for button/link.
  def click(element, argument, result)
    element.click
    result[0] = true
    result
  end

  # action for the component of dropdown box
  # element: component object for dropdown list.
  # argument: the item which you want to select.
  # result: result for the process, it is true as a default value here.
  def select_combobox(element, argument, result)
    element.select(argument[0])
    result[0] = true
    result
  end

  # action for the component of textbox
  # element: component object for textbox.
  # argument: the value which you want to set.
  # result: result for the process, it is true as a default value here.
  def set_textbox(element, argument, result)
    arg = argument[0]
    arg = "" if argument[0] == nil || argument[0] == ""
    element.set(arg)
    result[0] = true
    result
  end

  # action for the component of checkbox
  # element: component object for checkbox.
  # argument: the operation which you want to choose.
  # result: result for the process, it is true as a default value here.
  def set_checkbox(element, argument, result)
    method = argument[0]
    element.method("#{method}").call
    result[0] = true
    result
  end

  # action for the component of file field
  # element: component object for filefield.
  # argument: the file path which you want to choose.
  # result: result for the process, it is true as a default value here.
  def set_filefield(element, argument, result)
    element.set(argument[0])
    result[0] = true
    result
  end

  # read result form a web component
  # element: component object which you want to read its value.
  # argument: just for the reflect for "T" type action, it is unuseful in here.
  # result: element's value.
  def read_result(element, argument, result)
    ele_type = element.type
    case ele_type.to_s
    when "textarea","text"
      argu_actual = element.value
    when "checkbox"
      argu_actual = "set" if element.checked?
      argu_actual = "clear" unless element.checked?
    when "select-one"
      index = element.value
      argu_actual = element.option(:value, index).text
    end
    result[0] = argu_actual.to_s
    result
  end

  # read tip form a web component
  # element: component object which you want to read its value.
  # argument: just for the reflect for "T" type action, it is unuseful in here.
  # result: element's value.
  def get_tip(element, argument, result)
    result[0] = @browser.get_tip(element)
    result
  end

  # read status value form status table
  # element: component object which you want to read its value.
  # argument: just for the reflect for "T" type action, it is unuseful in here.
  # result: element's value.
  def get_status(element, argument, result)
    result[0] = @browser.get_status(element)
    result
  end

  # get error massage when the input value is invalid
  # element: component object which you want to read its value.
  # argument: just for the reflect for "T" type action, it is unuseful in here.
  # result: element's value.
  def error_message(element, argument, result)
    result[0] = @browser.get_invalid_text(element)
    result
  end

  # get description when update firmware
  # element: component object which you want to read its value.
  # argument: just for the reflect for "T" type action, it is unuseful in here.
  # result: description.
  def get_description(element, argument, result)
    result[0] = @browser.get_description(element)
    result
  end

  # Check the action of Login is needed or not
  # user: unity's username
  # pswd: unity's password
  # result: result for the process, it is true as a default value here.
  def login_check(user, pswd, result)
    result[0] = false
    thread = Thread.new do
      @browser.login(user,pswd)
      result[0] = true
    end
    thread.join # wait the thread finished
    waitsave(1)
#    # Click 'Edit' button
#    edit_ele = @browser.get_element("Edit")
#    edit_ele.click
    result
  end

  # Deal with the popup of restart
  # argument: which button you want to click.
  # result: record the alert text from popup
  def popup_check(argument, result)
    is_popup = @browser.popup_status
    result[0] = "Notice!\r\n\r\nRestart of the system is required for new settings to take effect. Once all configuration changes have been made, please restart the system by selecting \"Support\" on the navigation menu."
    if is_popup
      argument[0] = "OK"
      result = popup_reset(argument, result)
    else
      #if there is no restart popup, sleep 3 seconds to wait the new value update.
      waitsave(3)
    end
    #puts result
    result
  end

  # Deal with the popup of reset
  # argument: which button you want to click.
  # result: record the alert text from popup
  def popup_reset(argument, result)
    btn = argument[0]
    alert_text = @browser.popup(btn)
    result[0] = alert_text
    return result
  end

  # The action to add one item configuration for snmp(SNMPv1 Access, SNMPv1 Trap, SNMPv3)
  # user_info: card's username and password.
  # argument:  the values which you want to configure.
  # result: all elements'value for one webpage
  def write_snmp_group(user_info, argument,result)
    user = user_info[0]
    pswd = user_info[1]
    # Click 'Edit' button
    edit_ele = @browser.get_element("Edit")
    edit_ele.click

    # Check login?
    if !@login_flag
      result[0] = @browser.login(user, pswd)
      @login_flag = true
    end
    elements = @browser.get_muti_elements

    # Configure each element by using the given arguments
    config_elements(elements, argument)

    # Click 'Save' button
    save_ele = @browser.get_element("Save")
    save_ele.click

    #check the restart popup
    restart_result = [nil]
    popup_check(nil, restart_result)

    #wait the information saved
    #wait_result = [nil]
    waitsave(3)

    #Read the saved information
    result = read_snmp_group(user_info, argument,result)

  end

  # Config the elements for one webpage
  # elements: all elements' object in one webpage
  # argument: the values which you want to configure.
  def config_elements(elements, argument)
    # Configure each element by using the given arguments
    index = 0
    elements.values.each do |ele|
      arg = argument[index]
      case ele.type.to_s
      when "checkbox"
        ele.method("#{arg}").call
      when "textarea","text"
        ele.set(arg)
      when "select-one"
        ele.select(arg)
      end
      index += 1
    end
  end

  # Read the elements' result form one webpage
  # use_info and arguments: just for the reflect for "F" type action, it is unuseful in here.
  # result: all elements'value for one webpage
  def read_snmp_group(use_info, arguments,result)
    #Read the saved information
    result_index = 0
    config_elements = @browser.get_muti_elements
    config_elements.values.each do |ele|
      ele_type = ele.type
      case ele_type.to_s
      when "textarea", 'text'
        argu_actual = ele.value
      when "checkbox"
        argu_actual = "set" if ele.checked?
        argu_actual = "clear" unless ele.checked?
      when "select-one"
        index = ele.value
        argu_actual = ele.option(:value, index).text
      end
      result[result_index] = argu_actual
      result_index += 1
    end
    result
  end

  # The command of watisave
  # result: result for the process, it is true as a default value here.
  def waitsave(interval, result=[nil])
    interval = 3 if interval == nil || interval == ""
    sleep interval.to_i
    result[0] = true
    result
  end

  #reboot the card
  #user_info: card's username and password.
  #result: result for the process.
  def reboot(user_info, result)
    #TODO: navigate to support page.
    unity_ele = @browser.get_element("Unity")
    support_ele = @browser.get_element("Support")
    unity_ele.click
    support_ele.click

    user = user_info[0]
    pswd = user_info[1]
    enable_ele = @browser.get_element("Enable")
    enable_ele.click

    # Check login?
    if !@login_flag
      result[0] = @browser.login(user, pswd)
      @login_flag = true
    end

    # Click 'Reboot' button
    save_ele = @browser.get_element("Restart")
    save_ele.click

    # Reset popup appear, Click 'OK'
    argument = ['OK']
    popup_reset(argument, result)
    return result
  end

  # Refresh the webpage
  # argument: just for the reflect for "C" type action, it is unuseful in here.
  # result: result for the process, it is true as a default value here.
  def refresh_web(argument, result=[nil])
    @browser.page_refresh
    result[0] = true
    result
  end

  def get_card_info
    card_info = Array.new

    #support page
    support_ele = @browser.get_element("Support")
    support_ele.click

    #get card information from this page
    temp = ['Agent Date and Time', 'Agent Model','Agent App Firmware Version','Agent App Firmware Label',\
        'Agent Boot Firmware Version','Agent Boot Firmware Label','Agent Serial Number',\
        'Agent Manufacture Date','Agent Hardware Version','GDD Version','FDM Version']
    temp.each{|item|
      item_status = @browser.get_status(item)
      card_info << item_status
    }
    card_info
  end

end

