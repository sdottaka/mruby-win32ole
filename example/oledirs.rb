#
# You need WSH(Windows Scripting Host) to run this script.
#

require 'win32ole' unless Module.const_defined?(:WIN32OLE)

def listup(items)
#  items.each do |i|
  for i in items
    puts i.name
  end
end

fs = WIN32OLE.new("Scripting.FileSystemObject")

folder = fs.GetFolder(".")

puts "--- folder of #{folder.path} ---"
listup(folder.SubFolders)

puts "--- files of #{folder.path} ---"
listup(folder.Files)

