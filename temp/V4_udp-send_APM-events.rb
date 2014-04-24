=begin
The script will toggle some of the APM events


** Download the latest Velocity V4 simulator from the Twiki

1) Launch the Velocity V4 simulator
2) Enable "Remote Control On Port" check box in the simulator (lower right side)
3) Select the APM (NXr) device in the simulator and Start the simulator
4) Start a web card and point it to the velocity simulator
5) Observe the alarms on the APM web page (Main Page, Active Events or System Events)
6) Start this script and watch the Alarms come and go on the APM System Events web page
7) You must wait for the script to complete before trying to launch it again



Fundamental example:  
sock = UDPSocket.new
sock.bind("", 47123) # host port
sock.send("4295,0,19", 0, "127.0.0.1", 47809)
 
Where:
4295 = Velocity ID
0 = Multi-module index
19 = Active Alarm (0 = Clear Alarm)
0 = Socket
127.0.0.1 = host IP (local host in this example.  Can actually be any IP where simulator running
47809 = destination port

=end


require 'socket'

gdd_id = ["4295",     # Rectifier Failure
          "4233",     # Inverter Failure
          "4135",     # Bypass Not Available
          "4162"      # Battery Low
          ]

s = Time.now
puts "please wait a moment......."

sock = UDPSocket.new
host  = "127.0.0.1"
port = 47809 # destination port
sock.bind("", 47123) # source port

#set and clear some APM events
gdd_id.each do |id|
 sock.send("#{id},0,19", 0, host, port)   # alarm active
 sleep 2
 sock.send("#{id},0,0", 0, host, port)    # alarm clear
 sleep 2
end

sock.close

f = Time.now
print "script was completed in #{f - s} sec"




