require 'win32ole' unless Module.const_defined?(:WIN32OLE)
url = 'http://www.ruby-lang.org/'
ie = WIN32OLE.new('InternetExplorer.Application')
ie.visible = true
ie.gohome
print "Now navigate Ruby home page... Please enter."
gets if Kernel.respond_to?(:gets)
ie.navigate(url)
print "Now quit Internet Explorer... Please enter."
gets if Kernel.respond_to?(:gets)
ie.Quit()
