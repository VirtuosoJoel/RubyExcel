RubyExcel
=========

An attempt to create Excel-like Workbooks in Ruby.

####Still under construction! Bugs are inevitable.

Example
-------

```ruby

require 'RubyExcel'
wb = RubyExcel::Workbook.new

s = wb.load RubyExcel.sample_data
puts s

s.rows(3,4).each { |r| puts r['A'] }

puts s.range 'B2:C4'

s.range( s.cell(1,1), s.cell(2,2) ).delete
puts s

s.column(3).delete
s.column('D').each { |el| puts el }

s2 = wb.add 'NewSheet'
s2.load RubyExcel.sample_data.transpose
rng = s2.range 'A1:B3'
rng.value = rng.map { |cell| cell + 'a' }
	
```