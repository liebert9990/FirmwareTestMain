

# Show directory of 'file_type' files to choose from
def select_file_from_list(file_dir,file_type)
  Dir.chdir(file_dir)

  puts"Files Available:"
  
  # initialize file array and file counter  
  file_list = [nil]            
  file_counter = 1             

  # display file list, including the absolute path in console
  Dir.glob('*.' + file_type).each do |f| 
    puts "   #{file_counter}" + ' - ' + f
    file_list.push(File.expand_path(f)) 
    file_counter += 1
  end

  # Select the test file to display - show its absolute path
  # 'file_list' is an array of available files
  # 'file_number' is the element for file selected
  print "\nPlease select file number then press <enter>: "
  file_number = gets.to_i
  puts "Executing: #{file_number} - " + File.basename(file_list[file_number]) + "\n\n"
  return file_list[file_number]  
end


# Show list of script folders to choose from
def directory
  
  dir = Dir.open(".")         # open the current working directory

  # initialize folder array and folder counter
  folder_list = [nil]
  folder_counter = 1

  dir.each do |d|
    if d =~/^710/             # show only "710.*" folders (with test cases).
      folder_list.push(File.expand_path(d))
      puts "   #{folder_counter}" + ' - ' + d
      folder_counter += 1
    end
  end
  dir.close

  # Select the test file to display - show its absolute path
  print "\nPlease select folder number then press <enter>: "
  folder_number = gets.to_i
  print "\nDisplaying Scripts for folder # #{folder_number} - "
  puts folder_list[folder_number]
  return folder_list[folder_number]   # 'file_list' is an array of available files
                                      # 'file_number' is the array element for file selected
end


# Change directory to location of script folders
Dir.chdir(File.dirname(__FILE__)+("./Keywordscripts"))  

# select the folder
selected_dir = directory
file_type = "xls"

# select the script
select_file_from_list(selected_dir,file_type)

