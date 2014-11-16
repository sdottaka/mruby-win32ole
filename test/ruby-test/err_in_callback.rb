require 'win32ole' unless Module.const_defined?(:WIN32OLE) unless Module.const_defined?(:WIN32OLE)
db = WIN32OLE.new('ADODB.Connection')
db.connectionString = "Driver={Microsoft Text Driver (*.txt; *.csv)};DefaultDir=.;"
ev = WIN32OLE_EVENT.new(db)
ev.on_event('WillConnect') {|*args|
  foo
}
db.open
WIN32OLE_EVENT.message_loop
