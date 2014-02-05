#
# Namespace for all RubyExcel Classes and Modules
#

module RubyExcel

  #
  # The front-end class for data manipulation and output.
  #

  class Sheet
    include Address

    # The Data underlying the Sheet
    attr_reader :data
    
    # The name of the Sheet
    attr_accessor :name
    
    # The number of rows treated as headers
    attr_accessor :header_rows
    
    # The Workbook parent of this Sheet
    attr_accessor :workbook

    alias parent workbook; alias parent= workbook=
    alias headers header_rows; alias headers= header_rows=    
    
    #
    # Creates a RubyExcel::Sheet instance
    #
    # @param [String] name the name of the Sheet
    # @param [RubyExcel::Workbook] workbook the Workbook which holds this Sheet
    #
    
    def initialize( name, workbook )
      @workbook = workbook
      @name = name
      @header_rows = 1
      @data = Data.new( self, [[]] )
    end
    
    #
    # Read a value by address
    #
    # @example
    #   sheet['A1']
    #   #=> "Part"
    #
    # @example
    #   sheet['A1:B2']
    #   #=> [["Part", "Ref1"], ["Type1", "QT1"]]
    #
    # @param [String] addr the address to access
    #

    def[]( addr )
      range( addr ).value
    end
    
    #
    # Write a value by address
    #
    # @example
    #   sheet['A1'] = "Bart"
    #   sheet['A1']
    #   #=> "Bart"
    #
    # @param (see #[])
    # @param [Object] val the value to write into the data
    #

    def []=( addr, val )
      range( addr ).value = val
    end

    #
    # Add data with the Sheet
    #
    # @param [Array<Array>, Hash<Hash>, RubyExcel::Sheet] other the data to add
    # @return [RubyExcel::Sheet] returns a new Sheet
    # @note When adding another Sheet it won't import the headers unless this Sheet is empty.
    #
 
    def +( other )
      dup << other
    end
    
    #
    # Subtract data from the Sheet
    #
    # @param [Array<Array>, RubyExcel::Sheet] other the data to subtract
    # @return [RubyExcel::Sheet] returns a new Sheet
    #
    
    def -( other )
      case other
      when Array ; Workbook.new.load( data.all - other )
      when Sheet ; Workbook.new.load( data.all - other.data.no_headers )
      else       ; fail ArgumentError, "Unsupported class: #{ other.class }"
      end
    end
    
    #
    # Append an object to the Sheet
    #
    # @param [Object] other the object to append
    # @return [self]
    # @note When adding another Sheet it won't import the headers unless this Sheet is empty.
    # @note Anything other than an an Array, Hash, Row, Column or Sheet will be appended to the first row
    #
    
    def <<( other )
      data << other
      self
    end
    
    # @deprecated Please use {#filter!} instead
    # @overload advanced_filter!( header, comparison_operator, search_criteria, ... )
    #   Filter on multiple criteria
    # @param [String] header a header to search under
    # @param [Symbol] comparison_operator the operator to compare with
    # @param [Object] search_criteria the value to filter by
    # @raise [ArgumentError] 'Number of arguments must be a multiple of 3'
    # @raise [ArgumentError] 'Operator must be a symbol'
    # @example Filter to 'Part': 'Type1' and 'Type3', with Qty greater than 1
    #   s.advanced_filter!( 'Part', :=~, /Type[13]/, 'Qty', :>, 1 )
    # @example Filter to 'Part': 'Type1', with 'Ref1' containing 'X'
    #   s.advanced_filter!( 'Part', :==, 'Type1', 'Ref1', :include?, 'X' )
    #
    
    def advanced_filter!( *args )
      warn "[DEPRECATION] `advanced_filter!` is deprecated.  Please use `filter!` instead."
      data.advanced_filter!( *args ); self
    end
    
    #
    # Average the values in a Column by searching another Column
    #
    # @param [String] find_header the header of the Column to yield to the block
    # @param [String] avg_header the header of the Column to average
    # @yield yields the find_header column values to the block
    #
    
    def averageif( find_header, avg_header )
      return to_enum( :sumif ) unless block_given?
      find_col, avg_col  = ch( find_header ), ch( avg_header )
      sum = find_col.each_cell_wh.inject([0,0]) do |sum,ce|
        if yield( ce.value )
          sum[0] += avg_col[ ce.row ]
          sum[1] += 1
          sum 
        else
          sum
        end
      end
      sum.first.to_f / sum.last
    end
    
    #
    # Access an Cell by indices.
    #
    # @param [Fixnum] row the row index
    # @param [Fixnum] col the column index
    # @return [RubyExcel::Cell]
    # @note Indexing is 1-based like Excel VBA
    #
    
    def cell( row, col )
      Cell.new( self, indices_to_address( row, col ) )
    end
    alias cells cell
    
    #
    # Delete all data and headers from Sheet
    # 
    
    def clear_all
      data.delete_all
      self
    end
    alias delete_all clear_all
    
    #
    # Access a Column (Section) by its reference.
    #
    # @param [String, Fixnum] index the Column reference
    # @return [RubyExcel::Column]
    # @note Index 'A' and 1 both select the 1st Column
    #
    
    def column( index )
      Column.new( self, col_letter( index ) )
    end
    
    #
    # Access a Column (Section) by its header.
    #
    # @param [String] header the Column header
    # @return [RubyExcel::Column]
    #
    
    def column_by_header( header )
      header.is_a?( Column ) ? header : Column.new( self, data.colref_by_header( header ) )
    end
    alias ch column_by_header
    
    #
    # Yields each Column to the block
    #
    # @param [String, Fixnum] start_column the Column to start looping from
    # @param [String, Fixnum] end_column the Column to end the loop at
    # @note Iterates to the last Column in the Sheet unless given a second argument.
    #
    
    def columns( start_column = 'A', end_column = data.cols )
      return to_enum( :columns, start_column, end_column ) unless block_given?
      ( col_letter( start_column )..col_letter( end_column ) ).each { |idx| yield column( idx ) }
      self
    end
    
    #
    # Removes empty Columns and Rows
    #
    
    def compact!
      data.compact!; self
    end
    
    #
    # Removes Sheet from the parent Workbook
    #
    
    def delete
      workbook.delete self
    end
    
    #
    # Deletes each Row where the block is true
    #
    
    def delete_rows_if
      return to_enum( :delete_rows_if ) unless block_given?
      rows.reverse_each { |r| r.delete if yield r }; self
    end
    
    #
    # Deletes each Column where the block is true
    #
    
    def delete_columns_if
      return to_enum( :delete_columns_if ) unless block_given?
      columns.reverse_each { |c| c.delete if yield c }; self
    end
    
    #
    # Return a copy of self
    #
    # @return [RubyExcel::Sheet]
    #
    
    def dup
      s = Sheet.new( name, workbook )
      d = data
      unless d.nil?
        d = d.dup
        s.load( d.all, header_rows )
        d.sheet = s
      end
      s
    end
    
    #
    # Check whether the Sheet contains data (not counting headers)
    #
    # @return [Boolean] if there is any data
    #

    def empty?
      data.empty?
    end
    
    #
    # Export data to a specific WIN32OLE Excel Sheet
    #
    # @param win32ole_sheet the Sheet to export to
    # @return WIN32OLE Sheet
    #
    
    def export( win32ole_sheet )
      parent.dump_to_sheet( to_a, win32ole_sheet )
    end
    
    #
    # Removes all Rows (omitting headers) where the block is falsey
    #
    # @param [String, Array] headers splat of the headers for the Columns to filter by
    # @yield [Array] the values at the intersections of Column and Row
    # @return [self]
    #
    
    def filter!( *headers, &block )
      return to_enum( :filter!, headers ) unless block_given?
      data.filter!( *headers, &block ); self
    end
    
    #
    # Select and re-order Columns by a list of headers
    #
    # @param [Array<String>] headers the ordered list of headers to keep
    # @note This method can accept either a list of arguments or an Array
    # @note Invalid headers will be skipped
    #

    def get_columns!( *headers )
      data.get_columns!( *headers ); self
    end
    alias gc! get_columns!
    
    # @overload insert_columns( before, number=1 )
    #   Insert blank Columns into the data
    #
    #   @param [String, Fixnum] before the Column reference to insert before.
    #   @param [Fixnum] number the number of new Columns to insert
    #
    
    def insert_columns( *args )
      data.insert_columns( *args ); self
    end
    
    # @overload insert_rows( before, number=1 )
    #   Insert blank Rows into the data
    #
    #   @param [Fixnum] before the Row index to insert before.
    #   @param [Fixnum] number the number of new Rows to insert
    #
    
    def insert_rows( *args )
      data.insert_rows( *args ); self
    end
    
    #
    # View the object for debugging
    #
    
    def inspect
      "#{ self.class }:0x#{ '%x' % (object_id << 1) }: #{ name }"
    end
    
    #
    # The last Column in the Sheet
    #
    # @return [RubyExcel::Column]
    #
    
    def last_column
      column( maxcol )
    end
    alias last_col  last_column
    
    #
    # The last Row in the Sheet
    #
    # @return [RubyExcel::Row]
    #
    
    def last_row
      row( maxrow )
    end   
    
    #
    # Populate the Sheet with data (overwrite)
    #
    # @param [Array<Array>, Hash<Hash>] input_data the data to fill the Sheet with
    # @param header_rows [Fixnum] the number of Rows to be treated as headers
    #
    
    def load( input_data, header_rows=1 )
      input_data = _convert_hash(input_data) if input_data.is_a?(Hash)
      input_data.is_a?(Array) or fail ArgumentError, 'Input must be an Array or Hash'
      @header_rows = header_rows
      @data = Data.new( self, input_data ); self
    end
    
    #
    # Find the row number by looking up a value in a Column
    #
    # @param [String] header the header of the Column to pass to the block
    # @yield yields each value in the Column to the block
    # @return [Fixnum, nil] the row number of the first match or nil if nothing is found
    #
    
    def match( header, &block )
      row_id( column_by_header( header ).find( &block ) ) rescue nil
    end
    
    #
    # The highest currently used row number
    #
    
    def maxrow
      data.rows
    end
    alias length maxrow
    
    #
    # The highest currently used column number
    #
    
    def maxcol
      data.cols
    end
    alias maxcolumn maxcol
    alias width maxcol
    
    #
    # Allow shorthand range references and non-bang versions of bang methods.
    #
    
    def method_missing(m, *args, &block)
      method_name = m.to_s
      
      if method_name[-1] != '!' && respond_to?( method_name + '!' )
      
        dup.send( method_name + '!', *args, &block )
        
      elsif method_name =~ /\A[A-Z]{1,3}\d+=?\z/i
      
        method_name.upcase!
        if method_name[-1] == '='
          range( method_name.chop ).value = ( args.length == 1 ? args.first : args )
        else
          range( method_name ).value
        end
        
      else
        super
      end
    end
    
    #
    # Allow for certain method_missing calls
    #
    
    def respond_to?( m, include_private = false )
    
      if m[-1] != '!' && respond_to?( m.to_s + '!' )
        true
      elsif m.to_s.upcase.strip =~ /\A[A-Z]{1,3}\d+=?\z/
        true
      else
        super
      end
      
    end
    
    #
    # Split the Sheet into two Sheets by evaluating each value in a column
    #
    # @param [String] header the header of the Column which contains the yield value
    # @yield [value] yields the value of each row under the given header
    # @return [Array<RubyExcel::Sheet, RubyExcel::Sheet>] Two Sheets: true and false. Headers included.
    #
    
    def partition( header, &block )
      data.partition( header, &block ).map { |d| dup.load( d ) }
    end
    
    #
    # Access a Range by address.
    #
    # @param [String, Cell, Range] first_cell the first Cell or Address in the Range
    # @param [String, Cell, Range] last_cell the last Cell or Address in the Range
    # @return [RubyExcel::Range]
    # @note These are all valid arguments:
    #   ('A1') 
    #   ('A1:B2') 
    #   ('A:A')
    #   ('1:1')
    #   ('A1', 'B2') 
    #   (cell1) 
    #   (cell1, cell2) 
    #
    
    def range( first_cell, last_cell=nil )
      addr = to_range_address( first_cell, last_cell )
      addr.include?(':') ? Range.new( self, addr ) : Cell.new( self, addr )
    end
    
    #
    # Reverse the Sheet Columns
    #
    
    def reverse_columns!
      data.reverse_columns!; self
    end
    
    #
    # Reverse the Sheet Rows (without affecting the headers)
    #
    
    def reverse_rows!
      data.reverse_rows!; self
    end
    alias reverse! reverse_rows!
    
    #
    # Create a Row from an index
    #
    # @param [Fixnum] index the Row index
    # @return [RubyExcel::Row]
    #

    def row( index )
      Row.new( self, index )
    end
    
    #
    # Yields each Row to the block
    #
    # @param [Fixnum] start_row the Row to start looping from
    # @param [Fixnum] end_row the Row to end the loop at
    # @note Iterates to the last Row in the Sheet unless given a second argument.
    #
    
    def rows( start_row = 1, end_row = data.rows )
      return to_enum(:rows, start_row, end_row) unless block_given?
      ( start_row..end_row ).each { |idx| yield row( idx ) }; self
    end
    alias each rows
    
    #
    # Save the RubyExcel::Sheet as an Excel Workbook
    #
    # @param [String] filename the filename to save as
    # @param [Boolean] invisible leave Excel invisible if creating a new instance
    # @return [WIN32OLE::Workbook] the Workbook, saved as filename.
    #
    
    def save_excel( filename = nil, invisible = false )      
      workbook.dup.clear_all.add( self.dup ).workbook.save_excel( filename, invisible )
    end
    
    #
    # Sort the data by a column, selected by header(s)
    #
    # @param [String, Array<String>] headers the header(s) to sort the Sheet by
    #
    
    def sort_by!( *headers )
      raise ArgumentError, 'Sheet#sort_by! does not support blocks.' if block_given?
      idx_array = headers.flatten.map { |header| data.index_by_header( header ) - 1 }
      sort_method = lambda { |array| idx_array.map { |idx| array[idx] } }
      data.sort_by!( &sort_method )
      self
    rescue ArgumentError => err
      raise( NoMethodError, 'Item not comparable in "' + headers.flatten.map(&:to_s).join(', ') + '"' ) if err.message == 'comparison of Array with Array failed'
      raise err
    end
    
    #
    # Break the Sheet into a Workbook with multiple Sheets, split by the values under a header.
    #
    # @param [String] header the header to split by
    # @return [RubyExcel::Workbook] a new workbook containing the split Sheets (each with headers)
    #
    
    def split( header )
      wb = Workbook.new
      ch( header ).each_wh.to_a.uniq.each { |name| wb.add( name ).load( data.headers ) }
      rows( header_rows+1 ) do |row|
        wb.sheets( row.val( header ) ) << row
      end
      wb
    end
    
    #
    # Sum the values in a Column by searching another Column
    #
    # @param [String] find_header the header of the Column to yield to the block
    # @param [String] sum_header the header of the Column to sum
    # @yield yields the find_header column values to the block
    #
    
    def sumif( find_header, sum_header )
      return to_enum( :sumif ) unless block_given?
      find_col, sum_col  = ch( find_header ), ch( sum_header )
      find_col.each_cell.inject(0) { |sum,ce| yield( ce.value ) && ce.row > header_rows ? sum + sum_col[ ce.row ] : sum }
    end
    
    #
    # Return a Hash containing the Column values and the number of times each appears.
    #
    # @param [String] header the header of the Column to summarise
    # @return [Hash]
    #
    
    def summarise( header )
      ch( header ).summarise
    end
    alias summarize summarise
    
    #
    # Overwrite the sheet with the Summary of a Column
    # 
    # @param [String] header the header of the Column to summarise
    #
    
    def summarise!( header )
      load( summarise( header ).to_a.unshift [ header, 'Count' ] )
    end
    alias summarize! summarise!
    
    #
    # The Sheet as a 2D Array
    #
    
    def to_a
      data.all
    end
    
    #
    # The Sheet as a CSV String
    #
    
    def to_csv
      CSV.generate { |csv| to_a.each { |r| csv << r } }
    end
    
    #
    # The Sheet as a WIN32OLE Excel Workbook
    # @note This requires Windows and MS Excel
    #
    
    def to_excel
      workbook.dup.clear_all.add( self.dup ).workbook.to_excel
    end
    
    #
    # The Sheet as a String containing an HTML Table
    #
    
    def to_html
      %Q|<table border=1>\n<caption>#@name</caption>\n| + data.map { |row| '<tr>' + row.map { |v| '<td>' + CGI.escapeHTML(v.to_s) }.join }.join("\n") + "\n</table>"
    end
    
    #
    # The Sheet as a Tab Seperated Value String (Strips extra whitespace)
    #
    
    def to_s
      data.map { |ar| ar.map { |v| v.to_s.gsub(/\t|\n|\r/,' ') }.join "\t" }.join( $/ )
    end
    
    # {Sheet#to_safe_format!}
    
    def to_safe_format
      dup.to_safe_format!
    end
    
    #
    # Standardise the data for safe export to Excel.
    #   Set each cell contents to a string and remove leading equals signs.
    #
    
    def to_safe_format!
      rows { |r| r.map! { |v|
        if v.is_a?( String )
          v[0] == '=' ? v.sub( /\A=/,"'=" ) : v
        else
          v.to_s
        end
      } }; self
    end
    
    #
    # the Sheet as a TSV String
    #
    
    def to_tsv
      CSV.generate( :col_sep => "\t" ) { |csv| to_a.each { |r| csv << r } }
    end
    
    #
    # Remove any Rows with duplicate values within a Column
    #
    # @param [String] header the header of the Column to check for duplicates
    #
    
    def uniq!( header )
      data.uniq!( header ); self
    end
    alias unique! uniq!
    
    #
    # Select the used Range in the Sheet
    #
    # @return [Range] the Sheet's contents in Range
    #
    
    def usedrange
      raise NoMethodError, 'Sheet is empty' if empty?
      Range.new( self, 'A1:' + indices_to_address( maxrow, maxcol ) )
    end
    
    #
    # Find a value within a Column by searching another Column
    # 
    # @param [String] find_header the header of the Column to search
    # @param [String] return_header the header of the return value Column 
    # @yield the first matching value
    #
    
    def vlookup( find_header, return_header, &block )
      find_col, return_col  = ch( find_header ), ch( return_header )
      return_col[ row_id( find_col.find( &block ) ) ] rescue nil
    end
    
  end # Sheet
  
end # RubyExcel