require 'watir-webdriver'
require 'win32ole'
Watir.always_locate = false

$text_col = 1
$navframe = 'navigationFrame'
$detframe = 'detailArea'
$mutielems = Hash.new
$headcolor = 10
$bodycolor = 37

# recursively call this method to scan each page and determine the father-son relationship of each link
def scan_elements(tabl,fatherlink='None')
  i = 0
  temp = ""
  tempflink = ""
  row_cout = tabl.rows.size
  while i < row_cout
    j = 1
    while tabl.row(:index,i)[j].exist?
      if tabl.row(:index,i)[j].table().exist?
        if tabl.row(:index,i)[j].span().id.to_s =~/Child/
          temptemp = fatherlink
          fatherlink = tempflink
          tempflink = scan_elements(tabl.row(:index,i)[j].table(),fatherlink)
          fatherlink = temptemp
        else
          tempflink = scan_elements(tabl.row(:index,i)[j].table(),fatherlink)
        end  
      else
        temp = tabl.row(:index,i)[j].text.to_s
        if temp!="" #it is a page
          puts "scanning page - #{temp}"
          $brow.frame(:id, $navframe).link(:text, temp).click
          temph = Hash.new # for the elements temp itself
          temph[temp] = 'link'
          $mutielems[temp] = temph.merge(get_muti_buttons.merge(get_muti_elements))
          if fatherlink!='None'
            $mutielems[fatherlink][temp] = 'link' 
          end
          return temp
        end
        j  = j + 1
      end
      i = i + 1
    end
  end
end

# get all the elements in the table of current page, except buttons. buttons need to get by another method
def get_muti_elements
  elem_set = Hash.new
  tabl,rows = get_table
  for i in 1...rows
    key = tabl.row(:index, i)[$text_col].text
    if tabl.row(:index, i)[$text_col + 1].span().exist? # checkbox in a span
      value = tabl.row(:index, i)[$text_col + 1].span().element.to_subtype.type # to_subtype - focus to this specific elements
    else # other fields is not in a span
      value = tabl.row(:index, i)[$text_col + 1].element.to_subtype.type # to_subtype - focus to this specific elements
    end
    elem_set[key] = value
  end
  elem_set
end

# get the buttons in the current page
def get_muti_buttons
  butn_set = Hash.new
  butn_label = ['Edit','Save','Cancel','Enable','Restart','Reset to Factory Defaults','Web','Update']
  for i in 0...butn_label.length
    if $brow.frame(:id, $detframe).button(:value, butn_label[i]).exist?
      key = butn_label[i]
      value = $brow.frame(:id, $detframe).button(:value, butn_label[i]).type
      butn_set[key] = value
    end
  end
  butn_set
end

# get the specific table in the current page
def get_table(type = 'settingTable')
  #tabl = $brow.frame(:id, $detframe).div(:id, type).table()
  tabl = $brow.frame(:id, $detframe).div(:class, 'detailPanelArea').div(:id, type).table()
  if tabl.exist?
    row_cout = tabl.rows.size
  else
    row_cout = 0
  end
  return tabl, row_cout
end

def open_xls(path)
  ss = WIN32OLE::new('excel.Application')
  wb = ss.Workbooks.Open(path)
  ws = wb.Worksheets(1)
  ss.visible = true
  xls = [ss,wb,ws]
end

#  - close an existing workbook
def close_xls(s_s) # sheet number
  s_s[1].close
  # s_s.quit # session will remain active until the ruby script has completed
end

#  - save an existing workbook as another file name
def save_as_xls(s_s,save_as)
  sleep 1
  s_s[1].SaveAs(save_as)
end

# - time stamp in 'month-day_hour-minute-second' format
def t_stamp
  Time.now.strftime("%m-%d_%H-%M-%S")
end

def timeStamp(vari)
  ext = /\.\w*$/.match(vari).to_s # match extension from end of the string
  if ext
    vari.chomp(ext)+'_'+t_stamp+ext
  else
    vari+'_'+t_stamp
  end
end

#calc the excel header A,B,C,...AA,AB,AC...
def title_calc(a)
  b = a-65
  c = b/26
  m = b%26
  if c > 0
    e = (c + 64).chr
    d = (m + 65).chr
    e+d
  else
    d = (m + 65).chr
  end
end

$brow = Watir::Browser.new :chrome
$brow.goto('10.146.96.103')
sleep 1
$brow.frame(:id, 'tabArea').link(:id, 'tab4').click
tabl = $brow.frame(:id, $navframe).div(:id, 'navigationTree').table()
scan_elements(tabl)
#$mutielems = {"Configuration"=>{"System"=>"link", "User Settings"=>"link", "Network Settings"=>"link", "Web Server Settings"=>"link", "LIFE"=>"link", "Emerson Management Protocol"=>"link", "Messaging Settings"=>"link"}, "System"=>{"Edit"=>"button", "Save"=>"button", "Cancel"=>"button", "System Name"=>"textarea", "Contact Information"=>"textarea", "System Location"=>"textarea", "System Description"=>"textarea", "Display Temperature Units"=>"select-one", "Time Service Settings"=>"link"}, "Time Service Settings"=>{"Edit"=>"button", "Save"=>"button", "Cancel"=>"button", "External Time Source"=>"select-one", "NTP Time Server"=>"textarea", "NTP Time Sync Rate"=>"select-one", "Time Zone"=>"select-one", "Enable Auto-Sync To Managed Device"=>"checkbox", "Managed Device Auto-Sync Rate"=>"select-one"}, "User Settings"=>{"Edit"=>"button", "Save"=>"button", "Cancel"=>"button", "Administrator Username"=>"textarea", "Administrator Password"=>"text", "Reenter Administrator Password"=>"text", "General User Username"=>"textarea", "General User Password"=>"text", "Reenter General User Password"=>"text"}, "Network Settings"=>{"Edit"=>"button", "Save"=>"button", "Cancel"=>"button", "Speed Duplex"=>"select-one", "Domain Name Suffix List"=>"textarea", "IPv4 Settings"=>"link", "IPv6 Settings"=>"link", "Domain Name Server (DNS) Test"=>"link"}, "IPv4 Settings"=>{"Edit"=>"button", "Save"=>"button", "Cancel"=>"button", "Enable IPv4"=>"checkbox", "Boot Mode"=>"select-one", "Card IP Address"=>"text", "Subnet Mask"=>"text", "Default Gateway"=>"text", "DNS Server Address Source"=>"select-one", "IPv4 Primary DNS Server"=>"textarea", "IPv4 Secondary DNS Server"=>"textarea"}, "IPv6 Settings"=>{"Edit"=>"button", "Save"=>"button", "Cancel"=>"button", "Enable IPv6"=>"checkbox", "Boot Mode"=>"select-one", "Card IP Address"=>"textarea", "Prefix Length"=>"text", "Default Gateway"=>"textarea", "DNS Server Address Source"=>"select-one", "IPv6 Primary DNS Server"=>"textarea", "IPv6 Secondary DNS Server"=>"textarea"}, "Domain Name Server (DNS) Test"=>{"Edit"=>"button", "Save"=>"button", "Cancel"=>"button", "Type of Query"=>"select-one", "Query Value"=>"textarea"}, "Web Server Settings"=>{"Edit"=>"button", "Save"=>"button", "Cancel"=>"button", "Web Server Protocol"=>"select-one", "HTTP Port"=>"text", "HTTPS Port"=>"text", "Password Protected Site"=>"checkbox"}, "LIFE"=>{"Edit"=>"button", "Save"=>"button", "Cancel"=>"button", "LIFE Protocol"=>"checkbox", "System Serial Number"=>"textarea", "Site Equipment Tag Number"=>"textarea", "Site Identifier"=>"textarea", "Answer Incoming Call"=>"checkbox", "Next Call Date and Time"=>"textarea", "Call Interval Days"=>"text", "Call Interval Hours"=>"text", "Call Interval Minutes"=>"text", "Call Trials Number"=>"text", "LIFE Over IP"=>"link", "Advanced"=>"link"}, "LIFE Over IP"=>{"Edit"=>"button", "Save"=>"button", "Cancel"=>"button", "Proxy Enable"=>"checkbox", "Proxy Authentication"=>"checkbox", "Proxy IP Address"=>"textarea", "Proxy IP Port Number"=>"text", "Proxy User Name"=>"textarea", "Proxy User Password"=>"textarea", "LIFE Gate IP Address"=>"textarea", "LIFE Gate IP Port Number"=>"text"}, "Advanced"=>{"Edit"=>"button", "Save"=>"button", "Cancel"=>"button", "Enable"=>"button", "Service Mode"=>"checkbox"}, "Emerson Management Protocol"=>{"Managed Device"=>"link", "MSTP"=>"link", "Ethernet"=>"link", "Internal"=>"link"}, "Managed Device"=>{"Edit"=>"button", "Save"=>"button", "Cancel"=>"button", "LAN Type"=>"select-one", "Node ID (MSTP Only)"=>"text", "IP Address (IP Only)"=>"textarea"}, "MSTP"=>{"Edit"=>"button", "Save"=>"button", "Cancel"=>"button", "Data Rate"=>"select-one", "Max Master Address"=>"select-one", "Network Number"=>"text", "Node ID"=>"text"}, "Ethernet"=>{"Edit"=>"button", "Save"=>"button", "Cancel"=>"button", "Emerson Management Protocol IP Port Number"=>"text", "Network Number"=>"text"}, "Internal"=>{}, "Messaging Settings"=>{"Edit"=>"button", "Save"=>"button", "Cancel"=>"button", "Email"=>"checkbox", "SMS"=>"checkbox", "Email Settings"=>"link", "SMS Settings"=>"link", "Messaging Test"=>"link"}, "Email Settings"=>{"Edit"=>"button", "Save"=>"button", "Cancel"=>"button", "Email From Address"=>"textarea", "Email To Address"=>"textarea", "Email Subject Type"=>"select-one", "Custom Subject Text"=>"textarea", "SMTP Server Address"=>"textarea", "SMTP Server Port"=>"text", "Include IP Address in Message"=>"checkbox", "Include Event Description in Message"=>"checkbox", "Include Name in Message"=>"checkbox", "Include Contact in Message"=>"checkbox", "Include Location in Message"=>"checkbox", "Include Description in Message"=>"checkbox", "Include Web Link in Message"=>"checkbox", "Enable Event Consolidation"=>"checkbox", "Consolidation Time Limit"=>"text", "Consolidation Event Limit"=>"text"}, "SMS Settings"=>{"Edit"=>"button", "Save"=>"button", "Cancel"=>"button", "SMS From Address"=>"textarea", "SMS To Address"=>"textarea", "SMS Subject Type"=>"select-one", "Custom Subject Text"=>"textarea", "SMTP Server Address"=>"textarea", "SMTP Server Port"=>"text", "Include IP Address in Message"=>"checkbox", "Include Event Description in Message"=>"checkbox", "Include Name in Message"=>"checkbox", "Include Contact in Message"=>"checkbox", "Include Location in Message"=>"checkbox", "Include Description in Message"=>"checkbox", "Include Web Link in Message"=>"checkbox", "Enable Event Consolidation"=>"checkbox", "Consolidation Time Limit"=>"text", "Consolidation Event Limit"=>"text"}, "Messaging Test"=>{"Edit"=>"button", "Save"=>"button", "Cancel"=>"button", "Send Test Message"=>"select-one"}, "Protocols"=>{"Edit"=>"button", "Save"=>"button", "Cancel"=>"button", "SNMP Protocol"=>"checkbox", "BACnet Protocol"=>"checkbox", "Modbus Protocol"=>"checkbox", "Modbus"=>"link", "Modbus-TCP"=>"link", "Modbus-RS485"=>"link", "BACnet"=>"link", "SNMP"=>"link"}, "Modbus"=>{"Edit"=>"button", "Save"=>"button", "Cancel"=>"button", "Managed Device Write Access"=>"select-one", "Modbus Interface"=>"select-one"}, "Modbus-TCP"=>{"Edit"=>"button", "Save"=>"button", "Cancel"=>"button", "Limit Network Access Type"=>"select-one", "Port"=>"text", "Maximum Client Connection Count"=>"text", "Trusted IP List (5)"=>"link"}, "Trusted IP List (5)"=>{"Cancel"=>"button", "Trusted IP List [1]"=>"link", "Trusted IP List [2]"=>"link", "Trusted IP List [3]"=>"link", "Trusted IP List [4]"=>"link", "Trusted IP List [5]"=>"link"}, "Trusted IP List [1]"=>{"Edit"=>"button", "Save"=>"button", "Cancel"=>"button", "IP Address"=>"text"}, "Trusted IP List [2]"=>{"Edit"=>"button", "Save"=>"button", "Cancel"=>"button", "IP Address"=>"text"}, "Trusted IP List [3]"=>{"Edit"=>"button", "Save"=>"button", "Cancel"=>"button", "IP Address"=>"text"}, "Trusted IP List [4]"=>{"Edit"=>"button", "Save"=>"button", "Cancel"=>"button", "IP Address"=>"text"}, "Trusted IP List [5]"=>{"Edit"=>"button", "Save"=>"button", "Cancel"=>"button", "IP Address"=>"text"}, "Modbus-RS485"=>{"Edit"=>"button", "Save"=>"button", "Cancel"=>"button", "Node ID"=>"text", "Baud Rate"=>"select-one"}, "BACnet"=>{"Edit"=>"button", "Save"=>"button", "Cancel"=>"button", "Managed Device Write Access"=>"select-one", "BACnet/IP Port Number"=>"text", "Device Object Instance Number"=>"text", "Device Object Name"=>"textarea", "APDU Timeout"=>"text", "APDU Retries"=>"text", "Register as Foreign Device"=>"checkbox", "IP Address of BBMD"=>"text", "Foreign Device Time-to-Live"=>"text"}, "SNMP"=>{"Edit"=>"button", "Save"=>"button", "Cancel"=>"button", "SNMPv1/v2c Enable"=>"checkbox", "SNMPv3 Enable"=>"checkbox", "Authentication Traps"=>"checkbox", "Heartbeat Trap Interval"=>"select-one", "RFC-1628 MIB"=>"checkbox", "RFC-1628 MIB Traps"=>"checkbox", "Liebert Global Products (LGP) MIB"=>"checkbox", "LGP MIB Traps"=>"checkbox", "LGP MIB System Notify Trap"=>"checkbox", "SNMPv3 Users Setting (20)"=>"link", "SNMPv1/v2c Trap Settings (20)"=>"link", "SNMPv1/v2c Access Settings (20)"=>"link"}, "SNMPv3 Users Setting (20)"=>{"Cancel"=>"button", "SNMPv3 Users Setting [1]"=>"link", "SNMPv3 Users Setting [2]"=>"link", "SNMPv3 Users Setting [3]"=>"link", "SNMPv3 Users Setting [4]"=>"link", "SNMPv3 Users Setting [5]"=>"link", "SNMPv3 Users Setting [6]"=>"link", "SNMPv3 Users Setting [7]"=>"link", "SNMPv3 Users Setting [8]"=>"link", "SNMPv3 Users Setting [9]"=>"link", "SNMPv3 Users Setting [10]"=>"link", "SNMPv3 Users Setting [11]"=>"link", "SNMPv3 Users Setting [12]"=>"link", "SNMPv3 Users Setting [13]"=>"link", "SNMPv3 Users Setting [14]"=>"link", "SNMPv3 Users Setting [15]"=>"link", "SNMPv3 Users Setting [16]"=>"link", "SNMPv3 Users Setting [17]"=>"link", "SNMPv3 Users Setting [18]"=>"link", "SNMPv3 Users Setting [19]"=>"link", "SNMPv3 Users Setting [20]"=>"link"}, "SNMPv3 Users Setting [1]"=>{"Edit"=>"button", "Save"=>"button", "Cancel"=>"button", "SNMPv3 User Enable"=>"checkbox", "SNMPv3 Username"=>"textarea", "SNMPv3 Access Type"=>"select-one", "SNMPv3 Authentication"=>"select-one", "SNMPv3 Authentication Secret"=>"textarea", "SNMPv3 Privacy"=>"select-one", "SNMPv3 Privacy Secret"=>"textarea", "SNMPv3 Trap Target Addresses"=>"textarea", "SNMPv3 Trap Port"=>"text"}, "SNMPv3 Users Setting [2]"=>{"Edit"=>"button", "Save"=>"button", "Cancel"=>"button", "SNMPv3 User Enable"=>"checkbox", "SNMPv3 Username"=>"textarea", "SNMPv3 Access Type"=>"select-one", "SNMPv3 Authentication"=>"select-one", "SNMPv3 Authentication Secret"=>"textarea", "SNMPv3 Privacy"=>"select-one", "SNMPv3 Privacy Secret"=>"textarea", "SNMPv3 Trap Target Addresses"=>"textarea", "SNMPv3 Trap Port"=>"text"}, "SNMPv3 Users Setting [3]"=>{"Edit"=>"button", "Save"=>"button", "Cancel"=>"button", "SNMPv3 User Enable"=>"checkbox", "SNMPv3 Username"=>"textarea", "SNMPv3 Access Type"=>"select-one", "SNMPv3 Authentication"=>"select-one", "SNMPv3 Authentication Secret"=>"textarea", "SNMPv3 Privacy"=>"select-one", "SNMPv3 Privacy Secret"=>"textarea", "SNMPv3 Trap Target Addresses"=>"textarea", "SNMPv3 Trap Port"=>"text"}, "SNMPv3 Users Setting [4]"=>{"Edit"=>"button", "Save"=>"button", "Cancel"=>"button", "SNMPv3 User Enable"=>"checkbox", "SNMPv3 Username"=>"textarea", "SNMPv3 Access Type"=>"select-one", "SNMPv3 Authentication"=>"select-one", "SNMPv3 Authentication Secret"=>"textarea", "SNMPv3 Privacy"=>"select-one", "SNMPv3 Privacy Secret"=>"textarea", "SNMPv3 Trap Target Addresses"=>"textarea", "SNMPv3 Trap Port"=>"text"}, "SNMPv3 Users Setting [5]"=>{"Edit"=>"button", "Save"=>"button", "Cancel"=>"button", "SNMPv3 User Enable"=>"checkbox", "SNMPv3 Username"=>"textarea", "SNMPv3 Access Type"=>"select-one", "SNMPv3 Authentication"=>"select-one", "SNMPv3 Authentication Secret"=>"textarea", "SNMPv3 Privacy"=>"select-one", "SNMPv3 Privacy Secret"=>"textarea", "SNMPv3 Trap Target Addresses"=>"textarea", "SNMPv3 Trap Port"=>"text"}, "SNMPv3 Users Setting [6]"=>{"Edit"=>"button", "Save"=>"button", "Cancel"=>"button", "SNMPv3 User Enable"=>"checkbox", "SNMPv3 Username"=>"textarea", "SNMPv3 Access Type"=>"select-one", "SNMPv3 Authentication"=>"select-one", "SNMPv3 Authentication Secret"=>"textarea", "SNMPv3 Privacy"=>"select-one", "SNMPv3 Privacy Secret"=>"textarea", "SNMPv3 Trap Target Addresses"=>"textarea", "SNMPv3 Trap Port"=>"text"}, "SNMPv3 Users Setting [7]"=>{"Edit"=>"button", "Save"=>"button", "Cancel"=>"button", "SNMPv3 User Enable"=>"checkbox", "SNMPv3 Username"=>"textarea", "SNMPv3 Access Type"=>"select-one", "SNMPv3 Authentication"=>"select-one", "SNMPv3 Authentication Secret"=>"textarea", "SNMPv3 Privacy"=>"select-one", "SNMPv3 Privacy Secret"=>"textarea", "SNMPv3 Trap Target Addresses"=>"textarea", "SNMPv3 Trap Port"=>"text"}, "SNMPv3 Users Setting [8]"=>{"Edit"=>"button", "Save"=>"button", "Cancel"=>"button", "SNMPv3 User Enable"=>"checkbox", "SNMPv3 Username"=>"textarea", "SNMPv3 Access Type"=>"select-one", "SNMPv3 Authentication"=>"select-one", "SNMPv3 Authentication Secret"=>"textarea", "SNMPv3 Privacy"=>"select-one", "SNMPv3 Privacy Secret"=>"textarea", "SNMPv3 Trap Target Addresses"=>"textarea", "SNMPv3 Trap Port"=>"text"}, "SNMPv3 Users Setting [9]"=>{"Edit"=>"button", "Save"=>"button", "Cancel"=>"button", "SNMPv3 User Enable"=>"checkbox", "SNMPv3 Username"=>"textarea", "SNMPv3 Access Type"=>"select-one", "SNMPv3 Authentication"=>"select-one", "SNMPv3 Authentication Secret"=>"textarea", "SNMPv3 Privacy"=>"select-one", "SNMPv3 Privacy Secret"=>"textarea", "SNMPv3 Trap Target Addresses"=>"textarea", "SNMPv3 Trap Port"=>"text"}, "SNMPv3 Users Setting [10]"=>{"Edit"=>"button", "Save"=>"button", "Cancel"=>"button", "SNMPv3 User Enable"=>"checkbox", "SNMPv3 Username"=>"textarea", "SNMPv3 Access Type"=>"select-one", "SNMPv3 Authentication"=>"select-one", "SNMPv3 Authentication Secret"=>"textarea", "SNMPv3 Privacy"=>"select-one", "SNMPv3 Privacy Secret"=>"textarea", "SNMPv3 Trap Target Addresses"=>"textarea", "SNMPv3 Trap Port"=>"text"}, "SNMPv3 Users Setting [11]"=>{"Edit"=>"button", "Save"=>"button", "Cancel"=>"button", "SNMPv3 User Enable"=>"checkbox", "SNMPv3 Username"=>"textarea", "SNMPv3 Access Type"=>"select-one", "SNMPv3 Authentication"=>"select-one", "SNMPv3 Authentication Secret"=>"textarea", "SNMPv3 Privacy"=>"select-one", "SNMPv3 Privacy Secret"=>"textarea", "SNMPv3 Trap Target Addresses"=>"textarea", "SNMPv3 Trap Port"=>"text"}, "SNMPv3 Users Setting [12]"=>{"Edit"=>"button", "Save"=>"button", "Cancel"=>"button", "SNMPv3 User Enable"=>"checkbox", "SNMPv3 Username"=>"textarea", "SNMPv3 Access Type"=>"select-one", "SNMPv3 Authentication"=>"select-one", "SNMPv3 Authentication Secret"=>"textarea", "SNMPv3 Privacy"=>"select-one", "SNMPv3 Privacy Secret"=>"textarea", "SNMPv3 Trap Target Addresses"=>"textarea", "SNMPv3 Trap Port"=>"text"}, "SNMPv3 Users Setting [13]"=>{"Edit"=>"button", "Save"=>"button", "Cancel"=>"button", "SNMPv3 User Enable"=>"checkbox", "SNMPv3 Username"=>"textarea", "SNMPv3 Access Type"=>"select-one", "SNMPv3 Authentication"=>"select-one", "SNMPv3 Authentication Secret"=>"textarea", "SNMPv3 Privacy"=>"select-one", "SNMPv3 Privacy Secret"=>"textarea", "SNMPv3 Trap Target Addresses"=>"textarea", "SNMPv3 Trap Port"=>"text"}, "SNMPv3 Users Setting [14]"=>{"Edit"=>"button", "Save"=>"button", "Cancel"=>"button", "SNMPv3 User Enable"=>"checkbox", "SNMPv3 Username"=>"textarea", "SNMPv3 Access Type"=>"select-one", "SNMPv3 Authentication"=>"select-one", "SNMPv3 Authentication Secret"=>"textarea", "SNMPv3 Privacy"=>"select-one", "SNMPv3 Privacy Secret"=>"textarea", "SNMPv3 Trap Target Addresses"=>"textarea", "SNMPv3 Trap Port"=>"text"}, "SNMPv3 Users Setting [15]"=>{"Edit"=>"button", "Save"=>"button", "Cancel"=>"button", "SNMPv3 User Enable"=>"checkbox", "SNMPv3 Username"=>"textarea", "SNMPv3 Access Type"=>"select-one", "SNMPv3 Authentication"=>"select-one", "SNMPv3 Authentication Secret"=>"textarea", "SNMPv3 Privacy"=>"select-one", "SNMPv3 Privacy Secret"=>"textarea", "SNMPv3 Trap Target Addresses"=>"textarea", "SNMPv3 Trap Port"=>"text"}, "SNMPv3 Users Setting [16]"=>{"Edit"=>"button", "Save"=>"button", "Cancel"=>"button", "SNMPv3 User Enable"=>"checkbox", "SNMPv3 Username"=>"textarea", "SNMPv3 Access Type"=>"select-one", "SNMPv3 Authentication"=>"select-one", "SNMPv3 Authentication Secret"=>"textarea", "SNMPv3 Privacy"=>"select-one", "SNMPv3 Privacy Secret"=>"textarea", "SNMPv3 Trap Target Addresses"=>"textarea", "SNMPv3 Trap Port"=>"text"}, "SNMPv3 Users Setting [17]"=>{"Edit"=>"button", "Save"=>"button", "Cancel"=>"button", "SNMPv3 User Enable"=>"checkbox", "SNMPv3 Username"=>"textarea", "SNMPv3 Access Type"=>"select-one", "SNMPv3 Authentication"=>"select-one", "SNMPv3 Authentication Secret"=>"textarea", "SNMPv3 Privacy"=>"select-one", "SNMPv3 Privacy Secret"=>"textarea", "SNMPv3 Trap Target Addresses"=>"textarea", "SNMPv3 Trap Port"=>"text"}, "SNMPv3 Users Setting [18]"=>{"Edit"=>"button", "Save"=>"button", "Cancel"=>"button", "SNMPv3 User Enable"=>"checkbox", "SNMPv3 Username"=>"textarea", "SNMPv3 Access Type"=>"select-one", "SNMPv3 Authentication"=>"select-one", "SNMPv3 Authentication Secret"=>"textarea", "SNMPv3 Privacy"=>"select-one", "SNMPv3 Privacy Secret"=>"textarea", "SNMPv3 Trap Target Addresses"=>"textarea", "SNMPv3 Trap Port"=>"text"}, "SNMPv3 Users Setting [19]"=>{"Edit"=>"button", "Save"=>"button", "Cancel"=>"button", "SNMPv3 User Enable"=>"checkbox", "SNMPv3 Username"=>"textarea", "SNMPv3 Access Type"=>"select-one", "SNMPv3 Authentication"=>"select-one", "SNMPv3 Authentication Secret"=>"textarea", "SNMPv3 Privacy"=>"select-one", "SNMPv3 Privacy Secret"=>"textarea", "SNMPv3 Trap Target Addresses"=>"textarea", "SNMPv3 Trap Port"=>"text"}, "SNMPv3 Users Setting [20]"=>{"Edit"=>"button", "Save"=>"button", "Cancel"=>"button", "SNMPv3 User Enable"=>"checkbox", "SNMPv3 Username"=>"textarea", "SNMPv3 Access Type"=>"select-one", "SNMPv3 Authentication"=>"select-one", "SNMPv3 Authentication Secret"=>"textarea", "SNMPv3 Privacy"=>"select-one", "SNMPv3 Privacy Secret"=>"textarea", "SNMPv3 Trap Target Addresses"=>"textarea", "SNMPv3 Trap Port"=>"text"}, "SNMPv1/v2c Trap Settings (20)"=>{"Cancel"=>"button", "SNMPv1/v2c Trap Settings [1]"=>"link", "SNMPv1/v2c Trap Settings [2]"=>"link", "SNMPv1/v2c Trap Settings [3]"=>"link", "SNMPv1/v2c Trap Settings [4]"=>"link", "SNMPv1/v2c Trap Settings [5]"=>"link", "SNMPv1/v2c Trap Settings [6]"=>"link", "SNMPv1/v2c Trap Settings [7]"=>"link", "SNMPv1/v2c Trap Settings [8]"=>"link", "SNMPv1/v2c Trap Settings [9]"=>"link", "SNMPv1/v2c Trap Settings [10]"=>"link", "SNMPv1/v2c Trap Settings [11]"=>"link", "SNMPv1/v2c Trap Settings [12]"=>"link", "SNMPv1/v2c Trap Settings [13]"=>"link", "SNMPv1/v2c Trap Settings [14]"=>"link", "SNMPv1/v2c Trap Settings [15]"=>"link", "SNMPv1/v2c Trap Settings [16]"=>"link", "SNMPv1/v2c Trap Settings [17]"=>"link", "SNMPv1/v2c Trap Settings [18]"=>"link", "SNMPv1/v2c Trap Settings [19]"=>"link", "SNMPv1/v2c Trap Settings [20]"=>"link"}, "SNMPv1/v2c Trap Settings [1]"=>{"Edit"=>"button", "Save"=>"button", "Cancel"=>"button", "SNMP Trap Target Address"=>"textarea", "SNMP Trap Port"=>"text", "SNMP Trap Community String"=>"textarea"}, "SNMPv1/v2c Trap Settings [2]"=>{"Edit"=>"button", "Save"=>"button", "Cancel"=>"button", "SNMP Trap Target Address"=>"textarea", "SNMP Trap Port"=>"text", "SNMP Trap Community String"=>"textarea"}, "SNMPv1/v2c Trap Settings [3]"=>{"Edit"=>"button", "Save"=>"button", "Cancel"=>"button", "SNMP Trap Target Address"=>"textarea", "SNMP Trap Port"=>"text", "SNMP Trap Community String"=>"textarea"}, "SNMPv1/v2c Trap Settings [4]"=>{"Edit"=>"button", "Save"=>"button", "Cancel"=>"button", "SNMP Trap Target Address"=>"textarea", "SNMP Trap Port"=>"text", "SNMP Trap Community String"=>"textarea"}, "SNMPv1/v2c Trap Settings [5]"=>{"Edit"=>"button", "Save"=>"button", "Cancel"=>"button", "SNMP Trap Target Address"=>"textarea", "SNMP Trap Port"=>"text", "SNMP Trap Community String"=>"textarea"}, "SNMPv1/v2c Trap Settings [6]"=>{"Edit"=>"button", "Save"=>"button", "Cancel"=>"button", "SNMP Trap Target Address"=>"textarea", "SNMP Trap Port"=>"text", "SNMP Trap Community String"=>"textarea"}, "SNMPv1/v2c Trap Settings [7]"=>{"Edit"=>"button", "Save"=>"button", "Cancel"=>"button", "SNMP Trap Target Address"=>"textarea", "SNMP Trap Port"=>"text", "SNMP Trap Community String"=>"textarea"}, "SNMPv1/v2c Trap Settings [8]"=>{"Edit"=>"button", "Save"=>"button", "Cancel"=>"button", "SNMP Trap Target Address"=>"textarea", "SNMP Trap Port"=>"text", "SNMP Trap Community String"=>"textarea"}, "SNMPv1/v2c Trap Settings [9]"=>{"Edit"=>"button", "Save"=>"button", "Cancel"=>"button", "SNMP Trap Target Address"=>"textarea", "SNMP Trap Port"=>"text", "SNMP Trap Community String"=>"textarea"}, "SNMPv1/v2c Trap Settings [10]"=>{"Edit"=>"button", "Save"=>"button", "Cancel"=>"button", "SNMP Trap Target Address"=>"textarea", "SNMP Trap Port"=>"text", "SNMP Trap Community String"=>"textarea"}, "SNMPv1/v2c Trap Settings [11]"=>{"Edit"=>"button", "Save"=>"button", "Cancel"=>"button", "SNMP Trap Target Address"=>"textarea", "SNMP Trap Port"=>"text", "SNMP Trap Community String"=>"textarea"}, "SNMPv1/v2c Trap Settings [12]"=>{"Edit"=>"button", "Save"=>"button", "Cancel"=>"button", "SNMP Trap Target Address"=>"textarea", "SNMP Trap Port"=>"text", "SNMP Trap Community String"=>"textarea"}, "SNMPv1/v2c Trap Settings [13]"=>{"Edit"=>"button", "Save"=>"button", "Cancel"=>"button", "SNMP Trap Target Address"=>"textarea", "SNMP Trap Port"=>"text", "SNMP Trap Community String"=>"textarea"}, "SNMPv1/v2c Trap Settings [14]"=>{"Edit"=>"button", "Save"=>"button", "Cancel"=>"button", "SNMP Trap Target Address"=>"textarea", "SNMP Trap Port"=>"text", "SNMP Trap Community String"=>"textarea"}, "SNMPv1/v2c Trap Settings [15]"=>{"Edit"=>"button", "Save"=>"button", "Cancel"=>"button", "SNMP Trap Target Address"=>"textarea", "SNMP Trap Port"=>"text", "SNMP Trap Community String"=>"textarea"}, "SNMPv1/v2c Trap Settings [16]"=>{"Edit"=>"button", "Save"=>"button", "Cancel"=>"button", "SNMP Trap Target Address"=>"textarea", "SNMP Trap Port"=>"text", "SNMP Trap Community String"=>"textarea"}, "SNMPv1/v2c Trap Settings [17]"=>{"Edit"=>"button", "Save"=>"button", "Cancel"=>"button", "SNMP Trap Target Address"=>"textarea", "SNMP Trap Port"=>"text", "SNMP Trap Community String"=>"textarea"}, "SNMPv1/v2c Trap Settings [18]"=>{"Edit"=>"button", "Save"=>"button", "Cancel"=>"button", "SNMP Trap Target Address"=>"textarea", "SNMP Trap Port"=>"text", "SNMP Trap Community String"=>"textarea"}, "SNMPv1/v2c Trap Settings [19]"=>{"Edit"=>"button", "Save"=>"button", "Cancel"=>"button", "SNMP Trap Target Address"=>"textarea", "SNMP Trap Port"=>"text", "SNMP Trap Community String"=>"textarea"}, "SNMPv1/v2c Trap Settings [20]"=>{"Edit"=>"button", "Save"=>"button", "Cancel"=>"button", "SNMP Trap Target Address"=>"textarea", "SNMP Trap Port"=>"text", "SNMP Trap Community String"=>"textarea"}, "SNMPv1/v2c Access Settings (20)"=>{"Cancel"=>"button", "SNMPv1/v2c Access Settings [1]"=>"link", "SNMPv1/v2c Access Settings [2]"=>"link", "SNMPv1/v2c Access Settings [3]"=>"link", "SNMPv1/v2c Access Settings [4]"=>"link", "SNMPv1/v2c Access Settings [5]"=>"link", "SNMPv1/v2c Access Settings [6]"=>"link", "SNMPv1/v2c Access Settings [7]"=>"link", "SNMPv1/v2c Access Settings [8]"=>"link", "SNMPv1/v2c Access Settings [9]"=>"link", "SNMPv1/v2c Access Settings [10]"=>"link", "SNMPv1/v2c Access Settings [11]"=>"link", "SNMPv1/v2c Access Settings [12]"=>"link", "SNMPv1/v2c Access Settings [13]"=>"link", "SNMPv1/v2c Access Settings [14]"=>"link", "SNMPv1/v2c Access Settings [15]"=>"link", "SNMPv1/v2c Access Settings [16]"=>"link", "SNMPv1/v2c Access Settings [17]"=>"link", "SNMPv1/v2c Access Settings [18]"=>"link", "SNMPv1/v2c Access Settings [19]"=>"link", "SNMPv1/v2c Access Settings [20]"=>"link"}, "SNMPv1/v2c Access Settings [1]"=>{"Edit"=>"button", "Save"=>"button", "Cancel"=>"button", "SNMP Access Address"=>"textarea", "SNMP Access Type"=>"select-one", "SNMP Access Community String"=>"textarea"}, "SNMPv1/v2c Access Settings [2]"=>{"Edit"=>"button", "Save"=>"button", "Cancel"=>"button", "SNMP Access Address"=>"textarea", "SNMP Access Type"=>"select-one", "SNMP Access Community String"=>"textarea"}, "SNMPv1/v2c Access Settings [3]"=>{"Edit"=>"button", "Save"=>"button", "Cancel"=>"button", "SNMP Access Address"=>"textarea", "SNMP Access Type"=>"select-one", "SNMP Access Community String"=>"textarea"}, "SNMPv1/v2c Access Settings [4]"=>{"Edit"=>"button", "Save"=>"button", "Cancel"=>"button", "SNMP Access Address"=>"textarea", "SNMP Access Type"=>"select-one", "SNMP Access Community String"=>"textarea"}, "SNMPv1/v2c Access Settings [5]"=>{"Edit"=>"button", "Save"=>"button", "Cancel"=>"button", "SNMP Access Address"=>"textarea", "SNMP Access Type"=>"select-one", "SNMP Access Community String"=>"textarea"}, "SNMPv1/v2c Access Settings [6]"=>{"Edit"=>"button", "Save"=>"button", "Cancel"=>"button", "SNMP Access Address"=>"textarea", "SNMP Access Type"=>"select-one", "SNMP Access Community String"=>"textarea"}, "SNMPv1/v2c Access Settings [7]"=>{"Edit"=>"button", "Save"=>"button", "Cancel"=>"button", "SNMP Access Address"=>"textarea", "SNMP Access Type"=>"select-one", "SNMP Access Community String"=>"textarea"}, "SNMPv1/v2c Access Settings [8]"=>{"Edit"=>"button", "Save"=>"button", "Cancel"=>"button", "SNMP Access Address"=>"textarea", "SNMP Access Type"=>"select-one", "SNMP Access Community String"=>"textarea"}, "SNMPv1/v2c Access Settings [9]"=>{"Edit"=>"button", "Save"=>"button", "Cancel"=>"button", "SNMP Access Address"=>"textarea", "SNMP Access Type"=>"select-one", "SNMP Access Community String"=>"textarea"}, "SNMPv1/v2c Access Settings [10]"=>{"Edit"=>"button", "Save"=>"button", "Cancel"=>"button", "SNMP Access Address"=>"textarea", "SNMP Access Type"=>"select-one", "SNMP Access Community String"=>"textarea"}, "SNMPv1/v2c Access Settings [11]"=>{"Edit"=>"button", "Save"=>"button", "Cancel"=>"button", "SNMP Access Address"=>"textarea", "SNMP Access Type"=>"select-one", "SNMP Access Community String"=>"textarea"}, "SNMPv1/v2c Access Settings [12]"=>{"Edit"=>"button", "Save"=>"button", "Cancel"=>"button", "SNMP Access Address"=>"textarea", "SNMP Access Type"=>"select-one", "SNMP Access Community String"=>"textarea"}, "SNMPv1/v2c Access Settings [13]"=>{"Edit"=>"button", "Save"=>"button", "Cancel"=>"button", "SNMP Access Address"=>"textarea", "SNMP Access Type"=>"select-one", "SNMP Access Community String"=>"textarea"}, "SNMPv1/v2c Access Settings [14]"=>{"Edit"=>"button", "Save"=>"button", "Cancel"=>"button", "SNMP Access Address"=>"textarea", "SNMP Access Type"=>"select-one", "SNMP Access Community String"=>"textarea"}, "SNMPv1/v2c Access Settings [15]"=>{"Edit"=>"button", "Save"=>"button", "Cancel"=>"button", "SNMP Access Address"=>"textarea", "SNMP Access Type"=>"select-one", "SNMP Access Community String"=>"textarea"}, "SNMPv1/v2c Access Settings [16]"=>{"Edit"=>"button", "Save"=>"button", "Cancel"=>"button", "SNMP Access Address"=>"textarea", "SNMP Access Type"=>"select-one", "SNMP Access Community String"=>"textarea"}, "SNMPv1/v2c Access Settings [17]"=>{"Edit"=>"button", "Save"=>"button", "Cancel"=>"button", "SNMP Access Address"=>"textarea", "SNMP Access Type"=>"select-one", "SNMP Access Community String"=>"textarea"}, "SNMPv1/v2c Access Settings [18]"=>{"Edit"=>"button", "Save"=>"button", "Cancel"=>"button", "SNMP Access Address"=>"textarea", "SNMP Access Type"=>"select-one", "SNMP Access Community String"=>"textarea"}, "SNMPv1/v2c Access Settings [19]"=>{"Edit"=>"button", "Save"=>"button", "Cancel"=>"button", "SNMP Access Address"=>"textarea", "SNMP Access Type"=>"select-one", "SNMP Access Community String"=>"textarea"}, "SNMPv1/v2c Access Settings [20]"=>{"Edit"=>"button", "Save"=>"button", "Cancel"=>"button", "SNMP Access Address"=>"textarea", "SNMP Access Type"=>"select-one", "SNMP Access Community String"=>"textarea"}, "Status"=>{}, "Support"=>{"Cancel"=>"button", "Enable"=>"button", "Restart"=>"button", "Reset to Factory Defaults"=>"button", "Web"=>"button", "Active Networking"=>"link"}, "Active Networking"=>{}}
puts $mutielems

path = (File.dirname(__FILE__)+'/unity_case.xlsx').gsub('/','\\')
puts path
xl = open_xls(path)
newss = timeStamp(path)
wb = xl[1]
ws = wb.Worksheets(2)
a = 65  # ASCII A
ws.Range("A1").value = 'Pages'
ws.Range("A1").Interior.ColorIndex = $headcolor
i = 2
$mutielems.each_pair do |k,v|
  puts k, v
  ws.Range("A#{i}").value = k
ws.Range("A#{i}").Interior.ColorIndex = $bodycolor
  i = i + 1
  a = a + 1
  letter = title_calc(a)
  j = 2
  ws.Range("#{letter}1").value = k
  ws.Range("#{letter}1").Interior.ColorIndex = $headcolor
  v.each_key do |key|
    ws.Range("#{letter}#{j}").value = key
    ws.Range("#{letter}#{j}").Interior.ColorIndex = $bodycolor
    j = j + 1
  end
end

ws = wb.Worksheets(3)
ws.Range("A1").value = 'Component'
ws.Range("A1").Interior.ColorIndex = $headcolor
ws.Range("B1").value = 'Type'
ws.Range("B1").Interior.ColorIndex = $headcolor
i = 2
particular_set = Hash.new # delete the duplicate componets
$mutielems.each_value do |value|
  particular_set = particular_set.merge(value)
end
particular_set.each_pair do |k,v|
  ws.Range("A#{i}").value = k
  ws.Range("A#{i}").Interior.ColorIndex = $bodycolor
  case v
  when 'link'
    ws.Range("B#{i}").value = 'Link'
  when 'text'
    ws.Range("B#{i}").value = 'TextBox'
  when 'textarea'
    ws.Range("B#{i}").value = 'TextBox'
  when 'select-one'
    ws.Range("B#{i}").value = 'ComboBox'
  when 'checkbox'
    ws.Range("B#{i}").value = 'CheckBox'
  when 'button'
    ws.Range("B#{i}").value = 'Button'
  end
  ws.Range("B#{i}").Interior.ColorIndex = $bodycolor
  i = i + 1
end

save_as_xls(xl,newss)
close_xls(xl)
puts "done"
