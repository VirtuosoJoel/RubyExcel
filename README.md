RubyExcel
=========

An attempt to create Excel-like Workbooks in Ruby.

####Still under construction! Bugs are inevitable.

Example
-------
####Loading the class with example data:
```ruby

require 'RubyExcel'
wb = RubyExcel::Workbook.new

s = wb.load RubyExcel.sample_data
puts s

```

####Reference a cell's value:
```ruby
s['A7']
s.cell(7,1).value
s.range('A7').value
s.row(7)['A']
s.row(7)[1]
s.column('A')[7]
s.column('A')['7']

```
####Reference a group of cells:

```ruby
s['A1:B3'] #=> Value
s.range( 'A1:B3' ) #=> Element
s.range( s.cell( 1, 1 ), s.cell( 3, 2 ) ) #=> Element
s.row( 1 ) #=> Row
s.column( 'A' ) #=> Column
s.column( 1 ) #=> Column

```
####Advanced Interactions:
```ruby

puts s.column('D').map &:to_s

s2 = wb.add 'NewSheet'
s2.load RubyExcel.sample_data.transpose
rng = s2.range 'A1:B3'
rng.value = rng.map { |cell| cell + 'a' }

```

####Open and populate an Excel Workbook using win32ole
```ruby
RubyExcel::Workbook.new.load( RubyExcel.sample_data ).workbook.to_excel
	
```