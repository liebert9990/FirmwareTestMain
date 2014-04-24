require 'win32ole'

autoit = WIN32OLE.new('AutoItX3.Control') 

#Please put your simulator in C:/v4_sim
p Dir.chdir("c:/v4_sim")
p Dir.pwd

#Open and select simulator

p 'done'
  autoit.Run("V4SingleSimulatorGUI.exe")
   p "1"
  autoit.WinWaitActive("Emerson (Copyright 2012)")
  p "2"
  autoit.Send("{F1}")
   p "31"
   autoit.WinWaitActive("Open")
   p "3"
  autoit.ControlSetText("Open", "","Edit1","c:\\v4_sim\\NXr\\NXr_v866_2.19_08-16-13_09-36-56\\Nxr.exe")
  
   p "4"
  autoit.Send("{TAB}{TAB}")
   p "5"
  autoit.ControlClick("Open","&Open","Button1")
   p "6"
   autoit.WinWaitActive("SelectIPAddress")
   p "61"
  
  autoit.Send("{TAB}")
   p "7"
    
  autoit.ControlClick("SelectIPAddress","Select","WindowsForms10.BUTTON.app.0.bb8560_r15_ad12")
   p "8"
   
   #Start simulator
   autoit.Send("{F5}")
   
   
   
   #Stop Simulator
   
     autoit.WinWaitActive("NXr @ 10.163.230.113")
   p "9"
    autoit.ControlClick("NXr @ 10.163.230.113","Quit","WindowsForms10.BUTTON.app.0.bb8560_r15_ad12")
  p "10"
