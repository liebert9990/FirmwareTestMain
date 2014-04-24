require 'win32ole'
require "rexml/document"
include REXML

class Report

  def initialize
  end

  def creat_launch_xml(path)
    @launch_xml = WIN32OLE::new("MSXML2.DOMDocument")
    @launch_xml.async = false
    @launch_xml.loadXML('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><?xml-stylesheet type="text/xsl" href="TestReportGenerator.xslt"?><Launch></Launch>')
    @path = path

    #save xml file
    @launch_xml.save(@path)

  end

  def creat_xml_node(arg=nil)

    #creat testcase node
    testcase_node  =  @launch_xml.createElement("Testcase")
    @launch_xml.documentElement.appendChild(testcase_node)

    #Add testcase arrtibute name and other node
    node = ['name','ElapsedTime','SpreadsheetLink','TestcaseLogLink','Result']
    number = 0
    node.each{|x|
      if x == "name"
        attr = @launch_xml.createAttribute("name")
        testcase_node.setAttributeNode(attr)  #Add testcase arrtibute name
        attr.value = arg[number]
      else
        node = @launch_xml.createElement(x) #creat node
        testcase_node.appendChild(node)
        arg[number] = '' if arg[number] == nil
        node.text = arg[number]
      end

      number += 1

    }

    #save xml file
    @launch_xml.save(@path)
    return @launch_xml
  end

  #Create all xml nodes
  #arg is a planar array, it includes: node's name and node's index
  #An example: arg = [[0,'node1'],[1,'node2'],[3,'node3'],[4,'node4'],[5,'node5'],[6,'node6']]
  def create_xml_nodes(arg=nil)
    #Add testcase arrtibute name and other node
    nodes = ['name','index','ElapsedTime','SpreadsheetLink','TestcaseLogLink','Result']
    arg.each{|node_name|
      #creat testcase node
      testcase_node  =  @launch_xml.createElement("Testcase")
      @launch_xml.documentElement.appendChild(testcase_node)
      nodes.each{|node|
        case node
        when "name"
          attr = @launch_xml.createAttribute("name")
          testcase_node.setAttributeNode(attr)  #Add testcase arrtibute name
          attr.value = node_name[1]
        when "index"
          attr = @launch_xml.createAttribute("index")
          testcase_node.setAttributeNode(attr)  #Add testcase arrtibute index
          attr.value = node_name[0]
        when "Result"
          node = @launch_xml.createElement(node) #creat 'Result' node
          testcase_node.appendChild(node)
          node.text = "Not Running"
        when "ElapsedTime"
          node = @launch_xml.createElement(node) #creat 'ElapsedTime' node
          testcase_node.appendChild(node)
          node.text = "0"
        else
          node = @launch_xml.createElement(node) #creat node
          testcase_node.appendChild(node)
        end
      }
    }

    #save xml file
    @launch_xml.save(@path)
    return @launch_xml
  end

  #Edit an XML node
  #arg is an array, it includes: node's index, elapsed time,spreadsheet link,log file link, result
  def edit_xml_node(arg=nil)
    #select the needed node which index is given.
    node = @launch_xml.selectSingleNode("/Launch/Testcase[@index=#{arg[0]}]")
    if node.haschildnodes then
      for i in (0..node.childNodes.length-1)
        node.childnodes.item(i).text = arg[i+1]
        #puts node.childnodes.item(i).nodename
        #puts node.childnodes.item(i).text
      end
    end
    #save xml file
    @launch_xml.save(@path)
    return @launch_xml
  end

  #create a object for log report file.
  def create_log(path)
    log_file = File.new(path, "w+")
    return log_file
  end

  #write the log info into log report
  #log_file: object of log file.
  #arg: string of log information
  def write_log(log_file, arg=nil)
    if arg != nil
      log_file.puts(arg)
    end
  end

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
    
  #write the card information into xml report
  def write_card_info(start_time, arg)
    temp = ['name','datetime','Agent_Date_Time',\
        'Agent_Model_Unity','Agent_App_Firmware_Version','Agent_App_Firmware_Label',\
        'Agent_Boot_Firmware_Version','Agent_Boot_Firmware_Label','Agent_Serial_Number',\
        'Agent_Manufacture_Date','Agent_Hardware_Version','GDD_Version','FDM_Version','Browser_Type']

    attrb = []

    if arg == nil
      puts("Card information save failed.")
    else
      #design Launch attribute
      temp.each{|x|attrb << @launch_xml.createAttribute(x)}

      #set Launch attribute value and set attribute to Launch node
      dirnode = @launch_xml.documentElement

      number = 0
      attrb.each{|y|
        if  y.nodeName == "name"
          y.value = "launch.xml"
        elsif y.nodeName == "datetime"
          y.value = start_time
        else
          arg[number] = '' if arg[number] == nil
          y.value = arg[number]
          number += 1
        end

        #set attribute to Launch node
        dirnode.setAttributeNode(y)
      }
    end
    #save xml file
    @launch_xml.save(@path)
  end

end
#a = Report.new
#a.creat_launch_xml([2,55,4,5,6,8,7,9,10,11,12,13],"c:\\test.xml")
#a.create_xml_nodes([[0,'case1'],[1,'y'],[3,'u'],[4,'i'],[5,'o'],[6,'e']])
#a.edit_xml_node([0,'666','aaa','bbb','Pass'])
#a.edit_xml_node([5,'555','aaa','bbb','Pass'])
#a.edit_xml_node([4,'444','aaa','bbb','Pass'])
#a.creat_xml_node(['case2',2,3,4])
#a.create_log("c:\\test.txt")
#a.write_log("write log infomation--------------")
#a.write_log("write first line--------------")
#a.write_log("write second line--------------")
#a.write_log("write third line--------------")

#report = Report.new
#arg = ["710.040 System/710.040.10 Layout and Defaults/710.040.10.110.xls","Pass"]
#report.write_testlog(arg)
