RubyExcel
=========

Designed for Ruby on Windows with MS Excel

Introduction
------------

A Data-analysis tool for Ruby, with an Excel-style API.

You can find the gem [here](https://rubygems.org/gems/rubyexcel "Rubygems").

Main documentation is [here](http://rubydoc.info/gems/rubyexcel "Rubydoc")

For any requests, comments, etc. I keep an eye on [this forum](http://www.ruby-forum.com/forum/ruby "Ruby Mailing List").
If you put "RubyExcel" in the subject title I should see it.

Please feel free to log any enhancement requests or bugs [here](https://github.com/VirtuosoJoel/RubyExcel/issues "Bug Tracker").

Details
-----

This gem was made to simplify the steps between data extraction and the final output.
You can drop the data into it, reorganise and edit it, and then output into Excel or your preferred file format.
The methods provided for Excel interaction will return the relevant Excel object, allowing you to get as detailed as you like with the output.

Key design features taken from Excel:

* 1-based indexing.
* Referencing objects like Excel's API ( Workbook, Sheet, Row, Column, Cell, Range ).
* Useful data-handling functions ( e.g. Filter, Match, Sumif, Vlookup ).

Typical usage:

1. Extract a HTML Table or CSV File into 2D Array ( normally with Nokogiri / Mechanize ).
2. Organise and interpret data with RubyExcel.
3. Output results into a file.

About
-----

This gem is designed as a way to conveniently edit table data before outputting it to Excel (XLSX) or TSV format (which Excel can interpret).
It attempts to take as much as possible from Excel's API while providing some of the best bits of Ruby ( e.g. Enumerators, Blocks, Regexp ).
An important feature is allowing reference to Columns via their Headers for convenience and enhanced code readability.
As this works directly on the data, processing is faster than using Excel itself.

This was written out of the frustration of editing tabular data using Ruby's multidimensional arrays,
without affecting headers and while maintaining code readability.
Its API is designed to simplify moving code across from VBA into Ruby format when processing spreadsheet data.
The combination of Ruby, WIN32OLE Excel, and analysing table data is probably quite rare; but I thought I'd share what I came up with.

Examples
========

Expected Data Layout (2D Array)
--------

```ruby
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
```
The number of header rows defaults to 1

Loading the data into a Sheet
--------

```ruby
require 'rubyexcel'

wb = RubyExcel::Workbook.new
s = wb.add( 'Sheet1' )
s.load( data )

Or:

wb = RubyExcel::Workbook.new
s = wb.add( 'Sheet1' )
s.load( RubyExcel.sample_data )

Or:

wb = RubyExcel::Workbook.new
s = wb.load( RubyExcel.sample_data )

Or:

s = RubyExcel.sample_sheet
wb = s.parent
```

Using the Mechanize gem to get data
--------

This example is for context, there are many potential data sources

```ruby
s = RubyExcel::Workbook.new.load( CSV.parse( Mechanize.new.get('http://example.com/myfile.csv').content ) )
```

Reference a cell's value
--------

```ruby
s['A7']
s.A7
s.cell(7,1).value
s.range('A7').value
s.row(7)['A']
s.row(7)[1]
s.column('A')[7]
s.column('A')['7']
```

Reference a group of cells
--------

```ruby
s['A1:B3'] #=> Array
s.range( 'A1:B3' ) #=> Range
s.range( 'A:A' ) #=> Range (Column)
s.range( '1:2' ) #=> Range (Rows)
s.range( 'A1', 'B3' ) #=> Range
s.range( s.cell( 1, 1 ), s.cell( 3, 2 ) ) #=> Range
s.row( 1 ) #=> Row
s.column( 'A' ) #=> Column
s.column( 1 ) #=> Column
```

Using headers to reference the data
--------

Here we're looking for the "Part" in row 7

```ruby
s.row(7).value_by_header( 'Part' ) #=> "Type2"
s.row(7).val( 'Part' ) #=> "Type2"

s.column_by_header( 'Part' )[7] #=> "Type2"
s.ch( 'Part' )[7] #=> "Type2"

s.row(7).cell_by_header( 'Part' ) #=> Cell A7
s.row(7).cell_h( 'Part' ) #=> Cell A7

s.row(7).getref('Part') #=> "A"
```

Common Operations
--------

```ruby
#Some data to play with
s = RubyExcel.sample_sheet

#Have a look at the data
puts s

#Append a Column by adding a header
s << 'Number'

#Iterate through the rest of the rows while appending data
x = 1
s.rows(2) { |row| row << x; x+=1 }

#Filter to specific part numbers
s.filter!( 'Part', &/Type[1-3]/ )

#Sort by Part Number
s.sort_by!( 'Part' )

#Add the Number to the Cost in each row.
s.rows(2) { |row| row.cell_h('Cost').value += row.cell_h('Number').value }

#Split the data into multiple sheets by part number
wb = s.split( 'Part' )

#Output a sheet as a TSV file
File.write( 'Output.txt', wb.sheets(1).to_s )

#Output a sheet as an HTML page
File.write( 'Output.htm', wb.sheets(2).to_html )

#Open a sheet in an Excel Workbook
wb.sheets( 'Type3' ).to_excel

#If you feel the need to delete everything between a set of rows:
s.rows(2, 4).reverse_each &:delete

```

Workbook
--------

```ruby
#Create a workbook
wb = RubyExcel::Workbook.new
wb = RubyExcel::Workbook.new( 'My Workbook' )

#Get and set the name
wb.name = 'My Workbook'
wb.name #=> "My Workbook"

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

#Output the workbook as a series of HTML tables
wb.to_html
```

Sheet
--------

```ruby
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

#Filter on multiple criteria
#You can add as many arguments as you like. The order is: Header, Method(Symbol), Argument
#Note: Returns a copy of the sheet when used without "!".
#Note: Sheet#filter is simpler to use, this is for the more in-depth stuff.
_
#Filter to Part 'Type1' and 'Type3' where Qty is greater than 1
s.advanced_filter!( 'Part', :=~, /Type[13]/, 'Qty', :>, 1 )
_
#Filter to Part 'Type1' where Ref1 includes 'X'
s.advanced_filter!( 'Part', :==, 'Type1', 'Ref1', :include?, 'X' )

#Average all elements in a column by criteria in another column (selected by header)
#Parameters: Header to pass to the block, Header to average, Block.
#Note: Accepts Column objects in place of headers.
s.averageif( 'Part', 'Cost' ) { |part| part == 'Type1' } #=> 56.38333333333333
s.averageif( 'Part', 'Cost', &/Type1/ ) #=> 56.38333333333333

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

#Count is included via Enumerable and accepts a block which will yield Row or Column.
s.rows.count { |r| r.val('Part') == 'Type1' } #=> 3
s.rows.count { |r| r.val('Part') == 'Type1' && r.val('Qty') < 2 } #=> 2

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
#Note: Can now accept a Column object in place of a header.
s.match( 'Qty' ) { |value| value == 1 } #=> 2
s.match( 'Part', &/Type2/ ) #=> 3

#Find the current end of the data range
s.maxrow #=> 8
s.rows.count #=> 8
s.maxcol #=> 5
s.columns.count #=> 5

#Partition the sheet into two, given a header and a block (like Filter)
#Note: this keeps the headers intact in both output sheets
type_1_and_3, other = s.partition( 'Part' ) { |value| value =~ /Type[13]/ }
type_1_and_3, other = s.partition( 'Part', &/Type[13]/ )

#Reverse the data by rows or columns (ignores headers)
s.reverse_rows!
s.reverse_columns!

#Sort the rows by header(s) (ignores header rows)
s.sort_by!( 'Part' )
s.sort_by!( 'Qty', 'Part' )

#Split a Sheet into a Workbook of Sheets by a column (selected by header)
wb = s.split( 'Part' )
#=> <Workbook: [Sheet:Type1, Sheet:Type2, Sheet:Type3, Sheet:Type4]>

#Sum all elements in a column by criteria in another column (selected by header)
#Parameters: Header to pass to the block, Header to sum, Block.
#Note: Now also accepts Column objects in place of headers.
s.sumif( 'Part', 'Cost' ) { |part| part == 'Type1' } #=> 169.15
s.sumif( 'Part', 'Cost', &/Type1/ ) #=> 169.15

#Summarise a column by header into a Hash.
s.summarise( 'Part' )
#=> {"Type1"=>3, "Type2"=>2, "Type3"=>1, "Type4"=>1}

#Convert the data into various formats:
s.to_a #=> 2D Array
s.to_excel #=> WIN32OLE Excel Workbook (Contains only the current sheet)
s.to_html  #=> String (HTML table)
s.to_s #=> String (TSV)

#Remove all rows with duplicate values in the given column (selected by header or Column object)
s.uniq! 'Part'

#Find a value in one column by searching another one (selected by headers or Column objects)
s.vlookup( 'Part', 'Ref1', &/Type4/ ) #=> "XT3"
```

Row / Column (Section)
--------

```ruby
#Reference a Row or Column
row = s.row(2)
col = s.column('B')

=begin
Append a value
Note: Only extends the data boundaries when at the first row or column.
This allows looping through an entire row or column to append single values,
without worrying about using the correct index.
=end
s.row(1) << 'New'
s.rows(2) { |r| r << 'Column' }
s.column(1) << 'New'
s.columns(2) { |c| c << 'Row' }

#Access a cell by column header (Row only)
s.row(2).cell_by_header( 'Part' ) #=> Cell A2
s.row(2).cell_h( 'Cost' ) #=> Cell E2

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
col.each_without_headers { |val| puts val }
col.each_wh { |val| puts val }

#Loop through each cell
row.each_cell { |ce| puts "#{ ce.address }: #{ ce.value }" }
col.each_cell { |ce| puts "#{ ce.address }: #{ ce.value }" }

#Loop through each cell without including headers
col.each_cell_without_headers { |ce| puts "#{ ce.address }: #{ ce.value }" }
col.each_cell_wh { |ce| puts "#{ ce.address }: #{ ce.value }" }

#Get the letter of a column by its header ( helps build an address when combined with row.idx )
row.getref( 'Part' ) #=> "A"

#Overwrite each value based on its current value
row.map! { |val| val.to_s + 'a' }
col.map! { |val| val.to_s + 'a' }

#Get the value of a cell in the current row by its header
row.value_by_header( 'Part' ) #=> 'Type1'
row.val( 'Part' ) #=> 'Type1'

#Set the value of a cell in the current row by its header
row.set_value_by_header( 'Part', 'Type5' )
row.set_val( 'Part', 'Type5' )
```

Cell / Range (Element)
--------

```ruby
#Reference a Cell or Range
#If you select a single-Cell Range you get a Cell
cell = s.cell( 2, 2 ) #=> Cell
s.range('B2') #=> Cell
range = s.range('B2:C3') #=> Range

#Get the address and indices of the Element (Indices return that of the first Cell in a Range)
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

Address Tools (Included in Sheet, Section, and Element)
--------

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

#Here's a trick you can try if you're going to use the Address tools a lot:
include RubyExcel::Address
#Now you can use all the address methods in the current scope:
offset( 'C3', 1, -2 ) #=> A4

```

Importing a Hash
--------

```ruby
#Import a nested Hash (useful if you're summarising data before handing it to RubyExcel)

#Here's an example Hash (built into the gem as RubyExcel.sample_hash)
 h = {
      Part1: {
        Type1: {
          SubType1: 1, SubType2: 2, SubType3: 3
        },
        Type2: {
          SubType1: 4, SubType2: 5, SubType3: 6
        }
      },
      Part2: {
        Type1: {
          SubType1: 1, SubType2: 2, SubType3: 3
        },
        Type2: {
          SubType1: 4, SubType2: 5, SubType3: 6
        }
      }
    }

#Import the Hash to a Sheet
s.load( h )
#Or append the Hash to a Sheet
s << h

#Convert the symbols to strings (Not essential, but Excel can't handle Symbols in output)
s.rows { |r| r.map! { |v| v.is_a?(Symbol) ? v.to_s : v } }

#Have a look at the results
require 'pp'
pp s.to_a
[["Part1", "Type1", "SubType1", 1],
 ["Part1", "Type1", "SubType2", 2],
 ["Part1", "Type1", "SubType3", 3],
 ["Part1", "Type2", "SubType1", 4],
 ["Part1", "Type2", "SubType2", 5],
 ["Part1", "Type2", "SubType3", 6],
 ["Part2", "Type1", "SubType1", 1],
 ["Part2", "Type1", "SubType2", 2],
 ["Part2", "Type1", "SubType3", 3],
 ["Part2", "Type2", "SubType1", 4],
 ["Part2", "Type2", "SubType2", 5],
 ["Part2", "Type2", "SubType3", 6]]
 
```

Excel Tools ( requires win32ole and Excel )
--------

Make sure all your data types are compatible with Excel first (see Workbook#to_safe_format!)

RubyExcel should only attempt to require win32ole if one of these methods is called.
This should make it compatible as a data structure on non-windows systems.

```ruby

#Sample RubyExcel::Workbook to work with
rubywb = RubyExcel.sample_sheet.parent

#Get a new Excel instance
excel = rubywb.get_excel

#Get a new Excel Workbook
excelwb = rubywb.get_workbook( excel )
excelwb = rubywb.get_workbook

#Adds a leading single quote to any Strings with a leading Equals sign.
#This should prevent Excel from throwing Exceptions about invalid data such as "==".
rubywb.disable_formulas!

#Drop data into an Excel Sheet
#Note: The optional 2nd argument lets you give it an Excel sheet to use
rubywb.dump_to_sheet( rubywb.sheets(1).to_a )
rubywb.dump_to_sheet( rubywb.sheets(1).to_a, excelwb.sheets(1) )

#Autofit, left-align, and border a WIN32OLE Excel Sheet
rubywb.make_sheet_pretty( excelwb.sheets(1) )

#Output the RubyExcel::Workbook into a new Excel Workbook
rubywb.to_excel

#Output the RubyExcel::Sheet into a new Excel Workbook
rubywb.sheets(1).to_excel

#Output the RubyExcel::Workbook into an Excel Workbook and save the file
#Note: The default directory is "Documents" or "My Documents" to support Ocra + InnoSetup installs.
#The Workbook name is used as the filename unless an argument is passed.
#Note: There is an optional second argument which if set to true doesn't make Excel visible.
  #This is a useful accelerator when running as an automated process.
  #If you set the process to be invisible, don't forget to close Excel after you're finished with it!
rubywb.save_excel
rubywb.save_excel( 'Output' )
rubywb.save_excel( 'c:/example/Output.xlsx' )
rubywb.save_excel( 'Output', true ) # ( Keeps Excel invisible if creating a new instance )

#Convert all internal data to Strings and disable leading equals signs
#This should ensure compatibility with Excel for most types of data.
#Note: This will disable any formulas currently in the Workbook.
#Note: This modifies the data inside the Workbook. For a non-destructive version, leave off the exclamation mark.
rubywb.to_safe_format!

#Add borders to a given Excel Range
#1st Argument: WIN32OLE Range
#2nd Argument (default 1), weight of borders (0 to 4)
#3rd Argument (default false), include inner borders
RubyExcel.borders( excelwb.sheets(1).usedrange ) #Give used range outer borders
RubyExcel.borders( excelwb.sheets(1).usedrange, 2, true ) #Give used range inner and outer borders, medium weight
RubyExcel.borders( excelwb.sheets(1).usedrange, 0, false ) #Clear outer borders from used range

#You can even enter formula strings and Excel will evaluate them in the output.
s = rubywb.sheets(1)
s.row(1) << 'Formula'
s.rows(2) { |row| row << "=SUM(D#{ row.idx }:E#{ row.idx })" }
s.to_excel

```

Comparison of operations with and without RubyExcel gem
--------

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
#Excel should automatically target the "Documents" / "My Documents" folder when using a new instance.
#Excel should use the default file extension for the version it's using.
wb.saveas( 'Output' )
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

#Write the data to an Excel file ( requires Excel and win32ole )
#This aims at "Documents" or "My Documents" by default and uses the Workbook's "name" attribute.
s.parent.save_excel
```

Todo List
=========

- Find bugs and extirpate them.

- Optimise slow operations


-----------

Copyright (c) 2013, Joel Pearson and Contributors. All Rights Reserved.

This project is licenced under the [MIT License](LICENSE.md).