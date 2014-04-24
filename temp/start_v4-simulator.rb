# To change this template, choose Tools | Templates
# and open the template in the editor.

require 'win32ole'
require 'socket'
p'start'


# this example assumes that the the simulator is at "C:\v4_sim"
# change this directory path to much your environment
Dir.chdir("c:\v4_sim\NXr\NXr_v866_2.19_08-16-13_09-36-56\Nxr.exe")
p "working directory changed to - #{Dir.pwd}"

system "V4SingleSimulatorGUI.exe start.txt"

p 'done'