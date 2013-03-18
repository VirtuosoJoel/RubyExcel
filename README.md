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

s3 = wb.add
s3.load RubyExcel.sample_data
s3.header_rows = 1
s3.filter! 'B1', &/B[247]/
puts s3

```

####Open and populate an Excel Workbook using win32ole
```ruby
RubyExcel::Workbook.new.load( RubyExcel.sample_data ).workbook.to_excel
```

####Todo List:

- Some way to support "+=" and "-=" with each class?

- get the row number from a header (or other address type?) and a lookup value: =MATCH()

- get the address of a value: =FIND()

- unique the rows by a header

- add upcase and strip options for the data

- add tools to handle date conversion

- add the ability to summarise a column

- add a sumif and a countif

- add something to the excel dump which takes a range and puts outer borders on it, plus optional inner borders.

- add the ability to loop across a column or row while appending items. Maybe by referencing a section outside the existing range?

- add the ability to import (recursively?) a nested hash into something like this:

{ Type1: { SubType1: 1, SubType2: 2, SubType3: 3 }, Type2: { SubType1: 4, SubType2: 5, SubType3: 6 } }
<table>
<tr>
<td>Type1<td>SubType1<td>1
<tr><td><td>SubType2<td>2
<tr><td><td>SubType3<td>3
<tr><td>Type2<td>SubType1<td>4
<tr><td><td>SubType2<td>5
<tr><td><td>SubType3<td>6
