
=begin
Methods below can be used to stop the Velocity simulator

The script or method that starts the simulator will remain running/active
until the simulator is stopped or killed.


=end





  # - This method can be used at the beginning of a watir script to determine
  # - if another instance of ie is running.  Typically if multiple instances of
  # - ie are running watir scripts will hang on popups.
  #TODO - is this method using tasklist sufficient or should it use wmi??
  def is_process_running?(image_name)
    puts "Looking for instances of #{image_name}"
    command = 'tasklist /FI "IMAGENAME eq ' + "#{image_name}"""
    command_output = `#{command}`
    command_output.each_line do |line|
      if line =~ /^#{image_name}/
        return true
      end
    end
    return false
  end

  # - Returns an array of process ids (pids) for a given image name using the
  # - windows tasklist command
  def get_process_pids(image_name)
    pid_array = Array.new
    command = 'tasklist /FI "IMAGENAME eq ' + "#{image_name}"""
    command_output = `#{command}`
    command_output.each_line do |line|
      if line =~ /^#{image_name}/
        pid_array << line.split(/ +/)[1]
      end
    end
    return pid_array
  end

  # - Kills processes matched in the supplied array of pids using the windows
  # - taskkill command
  def kill_processes(pid_array)
    command = 'taskkill '
    pid_array.each do |pid|
      command = command + "/pid #{pid} "
    end
    puts "process killed" if `#{command}`
  end

# process that we want to kill
image_name = "V4SingleSimulatorGUI.exe"

# check to see if the process is running
puts is_process_running?(image_name)

# get the process id
puts ids = get_process_pids(image_name)

# kill the process
kill_processes(ids)