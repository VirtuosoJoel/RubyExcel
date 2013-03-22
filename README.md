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

#Or:
s = RubyExcel.sample_sheet

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

s = RubyExcel.sample_sheet

s.column('D').each_cell { |c| puts "#{ c.address }: #{ c.value }" }

s.range( 'A1:B3' ).map! { |val| val + 'a' }

s.filter! 'C1', &/C[247]/

```

####Open and populate an Excel Workbook using win32ole
```ruby
RubyExcel.sample_sheet.parent.to_excel
```

####Todo List:

- add the option to unique the rows by a header

- add upcase and strip options for the data

- add tools to handle date conversion

- add the ability to summarise a column

- add something to the excel tools which takes an excel sheet and a range, and puts outer borders on it, plus optional inner borders.

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
