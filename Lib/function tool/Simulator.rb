require 'win32ole'
require 'socket'

class Simulator

  attr_writer :sim_path
  attr_writer :sim_type
  attr_writer :sim_ip
  attr_writer :sim_port
  attr_writer :sim_GUI_path

  def initialize(sim_ip = "",
      sim_path     = 'c:\v4_sim\NXr\NXr_v866_2.19_08-16-13_09-36-56\Nxr.exe',
      sim_gui_path = 'c:\v4_sim', sim_type = 'NXr',sim_port = 47809)

    @sim_path     = sim_path
    @sim_type     = sim_type
    @sim_ip       = sim_ip
    @sim_port     = sim_port
    @sim_gui_path = sim_gui_path
  end


  #Open and select simulator
  def open_start

    @simulator = WIN32OLE.new('AutoItX3.Control')

    Dir.chdir(@sim_gui_path)
    Dir.pwd
  
    @simulator.Run("V4SingleSimulatorGUI.exe")
    @simulator.WinWaitActive("Emerson (Copyright 2012)")
    @simulator.ControlClick("Emerson (Copyright 2012)","Remote Control on Port:","WindowsForms10.BUTTON.app.0.bb8560_r15_ad11")
    @simulator.Send("{F1}")
    @simulator.WinWaitActive("Open")
    @simulator.ControlSetText("Open", "","Edit1",@sim_path)
    @simulator.ControlClick("Open","&Open","Button1")
    @simulator.WinWaitActive("SelectIPAddress")
    @simulator.ControlClick("SelectIPAddress","Select","WindowsForms10.BUTTON.app.0.bb8560_r15_ad12")
 
    #Start simulator
    @simulator.Send("{F5}")

  end

  #Stop Simulator
  def stop
    udpstop = UDPSocket.new
    udpstop.bind(@sim_ip, 47123) # source port
    udpstop.send("stop", 0, @sim_ip, @sim_port)   # alarm active   # alarm active
    udpstop.close
  end
  
  #Read udp socket
  def read(gdd_id = [], index = 0)
    udpread = UDPSocket.new
    re = []

    udpread.bind(@sim_ip, 47123) # source port
    
    gdd_id.each_index do |ind|
      temp = []
      temp =  udpread.send("#{gdd_id[ind].to_s},#{index.to_s}", 0, @sim_ip, @sim_port)
      puts "Read udp packed, gddid is #{gdd_id[ind]}"
      sleep 2
      re << ind+temp
    end
    udpread.close
    return re
  end


  #Send udp socket
  def write(gdd_id = [],value = [],index = 0 )
  
    udpsock = UDPSocket.new

    udpsock.bind(@sim_ip, 47123) # source port
    gdd_id.each_index do |ind|
      puts "#{gdd_id[ind].to_s},#{index.to_s},#{value[ind].to_s}"
      udpsock.send("#{gdd_id[ind].to_s},#{index.to_s},#{value[ind].to_s}", 0,@sim_ip,@sim_port)   # alarm active
      puts "Write udp packed, gddid is #{gdd_id[ind]}, value is #{value[ind].to_s}"
      sleep 2
    end
    udpsock.close
  end

 
end
s = Simulator.new("10.163.230.113")
#s.open_start
sleep 1
p s.read(["4295"])
sleep 1
s.read(["4295"],['3'])
s.read(["4233"],['11'])
s.read(["4135"],['19'])
s.read(["4162"],['27'])
p s.read(["4162"])

#s.stop
sleep 5
puts "stop"





