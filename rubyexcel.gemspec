Gem::Specification.new do |s|
  s.name        = 'rubyexcel'
  s.version     = '0.3.8'
  s.summary     = 'Spreadsheets in Ruby'
  s.description = "A tabular data structure in Ruby, with header-based helper methods for analysis and editing, and some of Excel's API style. Can output as 2D Array, HTML Table, CSV, TSV, or an Excel WIN32OLE Object"
  s.authors     = ['Joel Pearson']
  s.files       =  Dir.glob( 'lib/**/*.rb' ) + Dir.glob( '*.md' )
  s.homepage    = 'https://github.com/VirtuosoJoel'
  s.email       = 'VirtuosoJoel@gmail.com'
  s.license     = 'MIT'
end
