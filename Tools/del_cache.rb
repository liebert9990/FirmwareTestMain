#!/bin/ruby
# set TEMPIF=%USERPROFILE%\Local Settings\Temporary Internet Files
# %TEMPIF%\Content.IE5\Index.DAT

#TODO This file needs to be cleaned up
require 'Win32API'
require 'fileutils'

# HashMethods = a Hash that can be accessed with methods
#   e.g.    h = MethodHash.new
#           h['street'] = 'Broadway'
#           h.street    = 'Broadway'
#           puts h.street  ===> Broadway

require 'delegate'

class MethodHash < SimpleDelegator
  def initialize h = {}
    super h
  end
    
  def method_missing(method_name, *args)
    name = method_name.to_s
    if name.ends_with?('=')
      self[ name.chop ] = args[0]
    else
      self[ name ]
    end
  end
end


class String
  def ends_with?(substr)
    len = substr.length
    self.reverse() [0 .. len-1].reverse == substr
  end
    
  def starts_with?(substr)
    len = substr.length
    self[0 .. len-1] == substr
  end
    
  alias start_with?  starts_with?
  alias begin_with?  starts_with?
  alias begins_with? starts_with?
  alias end_with?    ends_with?

  # String each() operator reads line-by-line
  # These functions return characters
  def each_char
    self.each_byte{|x| yield x.chr }
  end
  def collect_char
    r = []
    self.each_byte{|x| r << x.chr }
    r
  end
end


def get_first_info(api)
  sizenum = [0,0].pack('L*')
  buf =  [1024,0].pack('L*').ljust(1024)
  r = api.FindFirstUrlCacheEntry.Call(nil,nil,sizenum)
  n = sizenum.unpack('L')[0]
  info = sizenum.ljust(n)
  [info,sizenum]
end

def get_next_info(api, handle)
  sizenum = [0,0].pack('L*')
  buf =  [1024,0].pack('L*').ljust(1024)
  r = api.FindNextUrlCacheEntry.Call(handle,nil,sizenum)
  n = sizenum.unpack('L')[0]
  info = sizenum.ljust(n)
  [info,sizenum]
end

#
# Win32 API used in this file is listed here.
# Each system call can be instatiated like this
#    DeleteUrlCacheEntry = Win32API.new("wininet", "DeleteUrlCacheEntry", ['P'], 'V')
# Instead, the functions are defined dynamically from this list:
#
FUNCS = {
  'FindFirstUrlCacheEntry' => ['ppp','n'],
  'FindNextUrlCacheEntry'  => ['npp','n'],
  'DeleteUrlCacheEntry'    => ['n',  'n'],
  'FindCloseUrlCache'      => ['n',  'n']
}

#
# get_api returns a hash with Win32 API system calls
# Usage:
#     api = get_api('Kernel32', {'GetLastError'=>['V', 'N'],...})
#
def get_api(library, function_hash)
  f = MethodHash.new
  function_hash.each{|funcname,types|
    in_types = types[0].collect_char{|x| x}
    out_type = types[1]
    f[funcname] = Win32API.new(library, funcname, in_types, out_type)
  }
  f
end


def getLastError
  f = Win32API.new('Kernel32', 'GetLastError', ['V'], 'N')
  f.call
end

def get_username
  #mswin32 used for 32bits, mingw32 used for 64 bits.
  if RUBY_PLATFORM =~ /mswin32|mingw32/
    query = `reg query HKEY_LOCAL_MACHINE\\SYSTEM\\ControlSet001\\Control\\ComputerName\\ActiveComputerName /v ComputerName`
    /ComputerName\s+REG_SZ\s+(\S+)/ =~ query
    puts "computer name is: #{$1}"

    #just used for windows XP
    query = `reg query HKEY_CURRENT_USER\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer /v "Logon User Name"`
    /Logon\ User\ Name\s+REG_SZ\s+(\S+)/ =~ query
    puts "user name is: #{$1}"
    # unix, cygwin, mac
  end
end


def delete_cache
  @window_os_flag = false
  @window_os_flag = true if RUBY_PLATFORM =~ /mswin32|mingw32/
  #@mac_os_flag = RUBY_PLATFORM.downcase.include?("darwin")
  #@browser_name = page.get_eval("navigator.userAgent")

  if  @window_os_flag.to_s == "true"
    if @browser_name.to_s =~ /IE/
      puts "Test is running in window platform with browser IE"
      puts Dir.glob('/Users/'+ENV['USERNAME']+'/AppData/Local/Microsoft/Windows/Temporary Internet Files/*')
      FileUtils.rm_r Dir.glob('/Users/'+ENV['USERNAME']+'/AppData/Local/Microsoft/Windows/Temporary Internet Files/*'),:force => true
      puts Dir.glob('/Users/'+ENV['USERNAME']+'/AppData/Local/Microsoft/Windows/Temporary Internet Files/*')
    end

    if @browser_name.to_s =~ /Chrome/
      puts ENV['USERNAME']
      puts "Test is running in window platform with browser google chrome"
      puts Dir.glob('/Users/'+ENV['USERNAME']+'/AppData/Local/Google/Chrome/User Data/Default/*')
      FileUtils.rm_r Dir.glob('/Users/'+ENV['USERNAME']+'/AppData/Local/Google/Chrome/User Data/Default/*'),:force => true
    end

    if @browser_name.to_s =~ /Firefox/
      puts "Test is running in window platform with browser Firefox"
      puts Dir.glob('/Users/'+ENV['USERNAME']+'/AppData/Local/Mozilla/Firefox/Profiles/*/Cache/*')
      FileUtils.rm_r Dir.glob('/Users/'+ENV['USERNAME']+'/AppData/Local/Mozilla/Firefox/Profiles/*/Cache/*'),:force => true
    end
  end
end

if __FILE__ == $0
  #    n = main
  #    print "Deleted #{n} items\n"
  #  get_username
  @browser_name = "IE"
  delete_cache
end




# It's not the best idea to delete the temporary files at the filesytem level
#like that.  it may cause errors in the IE instance. so don't think it's the best way to 
#call this method before we run launch script each time. 

# Instead, Take
#a look at the Win32API functions that you can use for this:
#<http://msdn.microsoft.com/en-us/library/aa384026%28VS.85%29.aspx>.
#
#You can access these functions in Ruby using the Win32API library or
#the dl/win32 library, both of which are part of the Ruby stdlib.

#I tried with sh %{ruby lib/del_cache.rb} and its working fine.

#I did in a different way. I need to call that in my script so that when
#I call, the complete del_cache.rb should run. So I used the below and
#its working fine.

require 'rake'
#sh %{ruby lib/del_cache.rb}