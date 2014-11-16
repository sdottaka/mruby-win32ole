MRuby::Gem::Specification.new('mruby-win32ole') do |spec|
  spec.licenses = ['Artistic License', 'GPL']
  spec.authors  = ['Takashi Sawanaka (mruby porter)', 'Masaki.Suketa (original author)', 'perl and ruby developers (see win32ole.c)']
  if ENV['OS'] == 'Windows_NT'
    spec.linker.libraries << ['ole32', 'oleaut32', 'advapi32', 'user32', 'uuid']
  else
    raise "mruby-win32ole does not support on non-Windows OSs."
  end
end
