
def defined?(o)
  begin
    o.class
    true
  rescue
    false
  end
end

STDERR = Kernel unless Module.const_defined?(:STDERR)

$: << "."

re = WIN32OLE.new("VBScript.RegExp")
re.pattern = "^test_.*\.rb"
for i in WIN32OLE.new("Scripting.FileSystemObject").GetFolder(".").files
  if re.Test(i.name)
    puts i.name
    require i.name 
  end
end

