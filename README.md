RubyExcel
=========

####Still under construction! Bugs are inevitable.

This gem is designed as a way to conveniently edit table data before outputting it to Excel (XLSX) or TSV format (which Excel can interpret).
It attempts to take as much as possible from Excel's API while providing some of the best bits of Ruby ( e.g. Enumerators, Blocks, Regexp )
As this works directly on the data, editing is faster than using Excel's API.

This was written out of the frustration of editing tabular data using Ruby's multidimensional arrays, without affecting headers and while maintaining code readability.

Key design features taken from Excel:
- 1-based indexing.
- referencing objects as if in Excel's API ( Workbook, Sheet, Row, Column, Cell, Range ).
- Referencing Columns via their Headers for convenience and enhanced code readability.
- Useful data-handling functions ( e.g. Filter, Match, Sumif ).

Typical usage:
1) Extract a HTML Table into 2D Array ( normally with Nokogiri )
2) Organise and interpret data with RubyExcel
3) Output results into a file.

Examples
-------

####Getting started with example data

This is an example of the expected layout of the sheet data (2D Array)
data = [
        [ 'Part',  'Ref1', 'Ref2', 'Qty', 'Cost' ],
        [ 'Type1', 'QT1',  '231',  1,     35.15  ], 
        [ 'Type2', 'QT3',  '123',  1,     40     ], 
        [ 'Type3', 'XT1',  '321',  3,     0.1    ], 
        [ 'Type1', 'XY2',  '132',  1,     30.00  ], 
        [ 'Type4', 'XT3',  '312',  2,     3      ], 
        [ 'Type2', 'QY2',  '213',  1,     99.99  ], 
        [ 'Type1', 'QT4',  '123',  2,     104    ]
       ]
The number of header rows defaults to 1
```ruby
require 'rubyexcel'

wb = RubyExcel::Workbook.new
s = wb.add( 'Sheet1' )
s.load( data )

#Or:

wb = RubyExcel::Workbook.new
s = wb.add( 'Sheet1' )
s.load( RubyExcel.sample_data )

#Or:

wb = RubyExcel::Workbook.new
s = wb.load( RubyExcel.sample_data )

#Or:

s = RubyExcel.sample_sheet
wb = s.parent
```

####Reference a cell's value

```ruby
s['A7']
s.cell(7,1).value
s.range('A7').value
s.row(7)['A']
s.row(7)[1]
s.column('A')[7]
s.column('A')['7']
```

####Reference a group of cells

```ruby
s['A1:B3'] #=> Value
s.range( 'A1:B3' ) #=> Element
s.range( 'A1', 'B3' ) #=> Element
s.range( s.cell( 1, 1 ), s.cell( 3, 2 ) ) #=> Element
s.row( 1 ) #=> Row
s.column( 'A' ) #=> Column
s.column( 1 ) #=> Column
```

####Detailed Interactions

Workbook
```ruby
#Create a workbook
wb = RubyExcel::Workbook.new

#Add sheets to the workbook
sheet1, sheet2 = wb.add('Sheet1'), wb.add

#Delete all sheets from a workbook
wb.clear_all

#Delete a specific sheet
wb.delete( 1 )
wb.delete( 'Sheet1' )
wb.delete( sheet1 )
wb.delete( /sheet1/i )

#Shortcut to create a sheet with a default name and fill it with data
wb.load( data )

#Select a sheet
wb.sheets(1) #=> RubyExcel::Sheet
wb.sheets('Sheet1') #=> RubyExcel::Sheet

#Iterate through all sheets
wb.sheets #=> Enumerator
wb.each #=> Enumerator

#Sort the sheets
wb.sort! { |x,y| x.name <=> y.name }
wb.sort_by! &:name
```

Sheet
```
#Create a sheet
s = wb.add #Name defaults to 'Sheet' + total number of sheets
s = wb.add( 'Sheet1' )

#Access the sheet name
s.name #=> 'Sheet1'
s.name = 'Sheet1'

#Access the parent workbook
s.workbook
s.parent

#Access the headers
s.header_rows #=> 1
s.headers #=> 1
s.headers = 1
s.header_rows = 1

#Specify the number of header rows when loading data
s.load( data, 1 )

#Append data (at the bottom of the sheet)
s << data
s << s
s += data
s += s

#Remove identical rows in another data set (skipping any headers)
s -= data
s -= s

#Select a column by its header
s.column_by_header( 'Part' )
s.ch( 'Part' )
#=> Column

#Iterate through rows or columns
s.rows { |r| puts r } #All rows
s.rows( 2 ) { |r| puts r } #From the 2nd to the last row
s.rows( 1, 3 ) { |r| puts r } #Rows 1 to 3
s.columns { |c| puts c } #All columns
s.columns( 'B' ) { |c| puts c } #From the 2nd to the last column
s.columns( 2 ) { |c| puts c } #From the 2nd to the last column
s.columns( 'B', 'D' ) { |c| puts c } #Columns 2 to 4
s.columns( 2, 4 ) { |c| puts c } #Columns 2 to 4

#Remove all empty rows & columns
s.compact!

#Delete the current sheet from the workbook
s.delete

#Delete rows or columns "if( condition )" (iterates in reverse to preserve references during loop)
s.delete_rows_if { |r| r.empty? }
s.delete_columns_if { |c| c.empty? }

#Filter the data given a column and a block to test values against.
#Note: Returns a copy of the sheet when used without "!".
#Note: This gem carries a Regexp to_proc method for Regex shorthand (shown below).
s.filter!( 'Part' ) { |value| value =~ /Type[13]/ }
s.filter!( 'Part', &/Type[13]/ )

#Filter the data to a specific set of columns by their headers.
#Note: Returns a copy of the sheet when used without "!".
s.get_columns!( 'Cost', 'Part', 'Qty' )
s.gc!( 'Cost', 'Part', 'Qty' )

#Insert blank rows or columns ( before, number to insert )
s.insert_rows( 2, 2 ) #Inserts 2 empty rows before row 2
s.insert_columns( 'B', 1 ) #Inserts 2 empty columns before column 2
s.insert_columns( 2, 1 ) #Inserts 2 empty columns before column 2

#Find the first row which matches a value within a column (selected by header)
s.match( 'Qty' ) { |value| value == 1 } #=> 2
s.match( 'Part', &/Type2/ ) #=> 3

#Find the current end of the data range
s.maxrow #=> 8
s.rows.count #=> 8
s.maxcol #=> 5
s.columns.count #=> 5

#Reverse the data by rows or columns (ignores headers)
s.reverse_rows!
s.reverse_columns!

#Sort the rows by criteria (ignores headers)
s.sort! { |r1,r2| r1['A'] <=> r2['A'] }
s.sort_by! { |r| r['A'] }

#Sum all elements in a column by criteria in another column (selected by header)
#Parameters: Header to pass to the block, Header to sum, Block.
s.sumif( 'Part', 'Cost' ) { |part| part == 'Type1' } #=> 169.15
s.sumif( 'Part', 'Cost', &/Type1/ ) #=> 169.15

#Remove all rows with duplicate values in the given column (selected by header)
s.uniq! 'Part'
```

Row / Column
```ruby
#Reference a Row or Column
row = s.row(2)
col = s.column('B')

#Append a value
#Note: Only extends the data boundaries when at the first row or column.
#This allows looping through an entire row or column to append single values without worrying about using the correct index.
s.row(1) << 'New'
s.rows(2) { |r| r << 'Column' }
s.column(1) << 'New'
s.columns(2) { |c| c << 'Row' }

#Delete the data referenced by self.
row.delete
col.delete

#Find the address of a cell matching a block
row.find { |value| value == 'QT1' }
row.find &/QT1/
col.find { |value| value == 'QT1' }
col.find &/QT1/

#Summarise the current row or column into a Hash.
s.column(1).summarise
#=> {"Type1"=>3, "Type2"=>2, "Type3"=>1, "Type4"=>1}

#Loop through all values
row.each { |val| puts val }
col.each { |val| puts val }

#Loop through all values without including headers
row.each_without_headers { |val| puts val }
row.each_wh { |val| puts val }

#Loop through each cell
row.each_cell { |ce| puts "#{ ce.address }: #{ ce.value }" }
col.each_cell { |ce| puts "#{ ce.address }: #{ ce.value }" }

#Overwrite each value based on its current value
row.map! { |val| val.to_s + 'a' }
col.map! { |val| val.to_s + 'a' }

#Get the value of a column in the current row from its header
row.value_by_header( 'Part' ) #=> 'Type1'
row.val( 'Part' ) #=> 'Type1'
```

Cell / Range (Elements)
```ruby
#Reference a Cell or Range
cell = s.cell( 2, 2 )
range = s.range('B2:C3')

#Get the address and indices of the Element (Indices return that of the first cell for multi-cell Ranges)
cell.address
cell.row
cell.column
range.address
range.row
range.column

#Get and set the value(s)
cell.value #=> "QT1"
cell.value = 'QT1'
range.value #=> [["QT1", "231"], ["QT3", "123"]]
range.value = "a"
range.value #=> [["a", "a"], ["a", "a"]]
range.value = [["QT1", "231"], ["QT3", "123"]]
range.value #=> [["QT1", "231"], ["QT3", "123"]]

#Loop through a range
range.each { |val| puts val }

#Loop through each cell within a range
range.each_cell { |ce| puts "#{ ce.address }: #{ ce.value }" }

```

####Address Tools (Included in Sheet, Section, and Element)
```ruby
#Get the column index from an address string
s.address_to_col_index( 'A2' ) #=> 1

#Translate an address to indices
s.address_to_indices( 'A2' ) #=> [ 2, 1 ]

#Translate letter(s) to a column index
s.col_index( 'A' ) #=> 1

#Translate a number to column letter(s)
s.col_letter( 1 ) #=> "A"

#Extract the column letter(s) or row number from an address
s.column_id( 'A2' ) #=> "A"
s.row_id( 'A2' ) #=> 2

#Expand a Range address
s.expand( 'A1:B2' ) #=> [["A1", "B1"], ["A2","B2"]]
s.expand( 'A1' ) #=> [["A1"]]

#Translate indices to an address
s.indices_to_address( 2, 1 ) #=> "A2"

#Offset an address by rows and columns
s.offset( 'A2', 1, 2 ) #=> "C3"
s.offset( 'A2', 2, 0 ) #=> "A4"
s.offset( 'A2', -1, 0 ) #=> "A1"

```

####Excel Tools for output convenience ( requires win32ole and Excel 2007 or later )
```ruby
#Sample RubyExcel::Workbook to work with
rubywb = RubyExcel.sample_sheet.parent

#Get a new Excel instance
excel = rubywb.get_excel

#Get a new Excel Workbook
excelwb = rubywb.get_workbook( excel )
excelwb = rubywb.get_workbook

#Drop data into an Excel Sheet
rubywb.dump_to_sheet( rubywb.sheets(1).to_a )
rubywb.dump_to_sheet( rubywb.sheets(1).to_a, excelwb.sheets(1) )

#Autofit and left-align a WIN32OLE Excel Sheet
rubywb.make_sheet_pretty( excelwb.sheets(1) )

#Output the RubyExcel::Workbook into a new Excel Workbook
rubywb.to_excel

#Output the RubyExcel::Sheet into a new Excel Workbook
rubywb.sheets(1).to_excel

#Output the RubyExcel::Workbook into an Excel Workbook and save the file
rubywb.save_excel
rubywb.save_excel( 'Output.xlsx' )
rubywb.save_excel( 'c:/example/Output.xlsx' )

```

###Comparison of operations with and without RubyExcel gem

Without RubyExcel (one way to to it):
```ruby
#Filter to only 'Part' of 'Type1' and 'Type3' while keeping the header row
idx = data[0].index( 'Part' )
data = [ data[0] ] + data[1..-1].select { |row| row[ idx ] =~ /Type[13]/ }

#Keep only the columns 'Cost' and 'Ref2' in that order
max_size = data.max_by(&:length).length #Standardise the row size to transpose into columns
data.map! { |row| row.length == max_size ? row : row + Array.new( max_size - row.length, nil) }
headers = [ 'Cost', 'Ref2' ]
data = data.transpose.select { |header,_| headers.index(header) }.sort_by { |header,_| headers.index(header) }.transpose

#Get the combined 'Cost' of every 'Part' of 'Type1' and 'Type3'
find_idx, sum_idx = data[0].index('Part'), data[0].index('Cost')
data[1..-1].inject(0) { |sum, row| row[find_idx] =~ /Type[13]/ ? sum + row[sum_idx] : sum }

#Write the data to a TSV file
output = data.map { |row| row.map { |el| "#{el}".strip.gsub( /\s/, ' ' ) }.join "\t" }.join $/
File.write( 'output.txt', output )

#Drop the data into an Excel sheet ( using Excel and win32ole )
excel = WIN32OLE::new( 'excel.application' )
excel.visible = true
wb = excel.workbooks.add
sheet = wb.sheets(1)
sheet.range( sheet.cells( 1, 1 ), sheet.cells( data.length, data[0].length ) ).value = data
wb.saveas( Dir.pwd.gsub('/','\\') + '\\Output.xlsx' )
```

With RubyExcel:
```ruby
#Filter to only 'Part' of 'Type1' and 'Type3' while keeping the header row
s.filter!( 'Part', &/Type[13]/ )

#Keep only the columns 'Cost' and 'Ref2' in that order
s.get_columns!( 'Cost', 'Ref2' )

#Get the combined 'Cost' of every 'Part' of 'Type1' and 'Type3'
s.sumif( 'Part', 'Cost', &/Type[13]/ )

#Write the data to a TSV file
File.write( 'output.txt', s.to_s )

#Write the data to an XLSX file ( requires Excel and win32ole )
s.parent.save_excel( 'Output.xlsx' )
```

##Todo List:

- add something to the excel tools which takes an excel sheet and a range, and puts outer borders on it, plus optional inner borders.

- add an option to split the data whilst retaining the headers in each output, like partition

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
