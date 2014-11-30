MRuby::Gem::Specification.new('mruby-win32ole') do |spec|
  spec.licenses = ['Artistic License', 'GPL']
  spec.authors  = ['Takashi Sawanaka (mruby porter)', 'Masaki.Suketa (original author)', 'perl and ruby developers (see win32ole.c)']
  spec.add_dependency 'mruby-time'
  
  if ENV['OS'] == 'Windows_NT'
    spec.linker.library_paths << '/lib/w32api' unless cc.command =~ /^cl(\.exe)?$/
    spec.linker.libraries << ['ole32', 'oleaut32', 'advapi32', 'user32', 'uuid']
  else
    raise "mruby-win32ole does not support on non-Windows OSs."
  end
end
