if Object.const_defined?(:MTest) && Kernel.respond_to?(:require) && Module.const_defined?(:Regexp)
  STDERR = Kernel unless Module.const_defined?(:STDERR)
  Test = MTest

  fso = WIN32OLE.new("Scripting.FileSystemObject")
  wsh = WIN32OLE.new("WScript.Shell")
  re  = WIN32OLE.new("VBScript.RegExp")

  old_wd = wsh.currentdirectory
  wsh.currentdirectory = $:[0] + "/../../mrbgems/mruby-win32ole/test/ruby-test"
  $: << "."

  require "test_folderitem2_invokeverb.rb"
  require "test_nil2vtempty.rb"
  require "test_ole_methods.rb"
  require "test_propertyputref.rb"
  require "test_thread.rb"
  require "test_win32ole.rb"
  require "test_win32ole_event.rb"
  require "test_win32ole_method.rb"
  require "test_win32ole_param.rb"
#  require "test_win32ole_record.rb"
  require "test_win32ole_type.rb"
  require "test_win32ole_typelib.rb"
  require "test_win32ole_variable.rb"
  require "test_win32ole_variant.rb"
  require "test_win32ole_variant_m.rb"
  require "test_win32ole_variant_outarg.rb"

  re.pattern = "^test_.*\.rb"
  for i in fso.GetFolder(".").files
    if re.Test(i.name)
      puts i.name
#      require i.name
    end
  end

  if $ok_test
    MTest::Unit.new.mrbtest
  else
    1.times {
      MTest::Unit.new.run
    }
  end

  wsh.currentdirectory = old_wd
else
  $asserts << "skipped mruby-win32ole test"  if $asserts
end

