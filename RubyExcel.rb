require_relative 'rubyexcel/rubyexcel_components.rb'
require_relative 'rubyexcel/excel_tools.rb'
require 'cgi'

#
# Ruby's standard Regexp class.
#   Regexp#to_proc is a bit of "syntactic sugar" which allows shorthand Regexp blocks
#
# @example
#   sheet.filter!( 'Part', &/Type[13]/ )
#

class Regexp
  def to_proc
    proc { |s| self =~ s.to_s }
  end
end

#
# Namespace for all RubyExcel Classes and Modules
#

module RubyExcel

  #
  # A Workbook which can hold multiple Sheets
  #

  class Workbook
    include Enumerable

    #
    # Creates a RubyExcel::Workbook instance.
    #
    
    def initialize
      @sheets = []
    end
    
    #
    # Appends an object to the Workbook
    #
    # @param [RubyExcel::Workbook, RubyExcel::Sheet, Array<Array>] other the object to append to the Workbook
    #
    
    def <<( other )
      case other
      when Workbook ; other.each { |sht| sht.workbook = self; @sheets << sht }
      when Sheet    ; @sheets << other; other.workbook = self
      when Array    ; load( other )
      else          ; fail TypeError, "Unsupported Type: #{ other.class }"
      end
      self
    end
    
    #
    # Adds a Sheet to the Workbook.
    #   If no argument is given, names the Sheet 'Sheet' + total number of Sheets
    #
    # @example
    #   sheet = workbook.add
    #   #=> RubyExcel::Sheet:0x2b3a0b8: Sheet1
    #
    # @param [nil, RubyExcel::Sheet, String] ref the identifier or Sheet to add
    # @return [RubyExcel::Sheet] the Sheet which was added
    
    def add( ref=nil )
      case ref
      when nil    ; s = Sheet.new( 'Sheet' + ( @sheets.count + 1 ).to_s, self )
      when Sheet  ; ( s = ref ).workbook = self
      when String ; s = Sheet.new( ref, self )
      else        ; fail TypeError, "Unsupported Type: #{ ref.class }"
      end
      @sheets << s
      s
    end
    alias add_sheet add
    
    #
    # Removes all Sheets from the Workbook
    #
    
    def clear_all
      @sheets = []; self
    end
    alias delete_all clear_all
    
    #
    # Removes Sheet(s) from the Workbook
    #
    # @param [Fixnum, String, Regexp, RubyExcel::Sheet] ref the reference or object to remove
    #
    
    def delete( ref )
      case ref
      when Fixnum ; @sheets.delete_at( ref - 1 )
      when String ; @sheets.reject! { |s| s.name == ref }
      when Regexp ; @sheets.reject! { |s| s.name =~ ref }
      when Sheet  ; @sheets.reject! { |s| s == ref }
      else        ; fail ArgumentError, 'Unrecognised Argument Type: ' + ref.class.to_s
      end
      self
    end
    
    #
    # Return a copy of self
    #
    # @return [RubyExcel::Workbook]
    #
    
    def dup
      wb = Workbook.new
      self.each { |s| wb.add s.dup }
      wb
    end
    
    #
    # Check whether the workbook has Sheets
    #
    # @return [Boolean] if there are any Sheets in the Workbook
    #
    
    def empty?
      @sheets.empty?
    end
    
    # @overload load( input_data, header_rows=1 )
    #   Shortcut to create a Sheet and fill it with data
    #   @param [Array<Array>, Hash<Hash>] input_data the data to fill the Sheet with
    #   @param Fixnum] header_rows the number of Rows to be treated as headers
    #
    
    def load( *args )
      add.load( *args )
    end

    #
    # Select a Sheet or iterate through them
    #
    # @param [Fixnum, String, nil] ref the reference to select a Sheet by
    # @return [RubyExcel::Sheet] if a search term was given
    # @return [Enumerator] if nil or no argument given
    #
    
    def sheets( ref=nil )
      return to_enum (:each) if ref.nil?
      ref.is_a?( Fixnum ) ? @sheets[ ref - 1 ] : @sheets.find { |s| s.name =~ /^#{ ref }$/i }
    end
    
    # {Workbook#sort!}
    
    def sort( &block )
      dup.sort!( &block )
    end
    
    #
    # Sort Sheets according to a block
    #
    
    def sort!( &block )
      @sheets = @sheets.sort( &block )
    end
    
    # {Workbook#sort_by!}
    
    def sort_by( &block )
      dup.sort_by!( &block )
    end
    
    #
    # Sort Sheets by an attribute given in a block
    #
    
    def sort_by!( &block )
      @sheets = @sheets.sort_by( &block )
    end
    
    #
    # Yields each Sheet.
    #
    
    def each
      return to_enum( :each ) unless block_given?
      @sheets.each { |s| yield s }
    end
    
  end # Workbook
  
  #
  # The front-end class for data manipulation and output.
  #

  class Sheet
    include Address
    include Enumerable

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
      @header_rows = nil
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
    # Append data to the Sheet
    #
    # @param [Array<Array>, Hash<Hash>, RubyExcel::Sheet] other the data to append
    # @return [self]
    #
    
    def <<( other )
      case other
      when Array ; load( data.all + other, header_rows )
      when Hash  ; load( data.all + _convert_hash( other ), header_rows )
      when Sheet ; load( data.all + other.data.no_headers, header_rows )
      else       ; fail ArgumentError, "Unsupported class: #{ other.class }"
      end
      self
    end
    
    # {Sheet#advanced_filter!}
    
    def advanced_filter( *args )
      dup.advanced_filter!( *args )
    end

    # @overload advanced_filter!( header, comparison_operator, search_criteria, ... )
    #   Filter on multiple criteria
    # @example Filter to 'Part': 'Type1' and 'Type3', with Qty greater than 1
    #   s.advanced_filter!( 'Part', :=~, /Type[13]/, 'Qty', :>, 1 )
    # @example Filter to 'Part': 'Type1', with 'Ref1' containing 'X'
    #   s.advanced_filter!( 'Part', :==, 'Type1', 'Ref1', :include?, 'X' )
    #
    # @param [String] header a header to search under
    # @param [Symbol] comparison_operator the operator to compare with
    # @param [Object] search_criteria the value to filter by
    # @raise [ArgumentError] 'Number of arguments must be a multiple of 3'
    # @raise [ArgumentError] 'Operator must be a symbol'
    #
    
    def advanced_filter!( *args )
      data.advanced_filter!( *args ); self
    end
    
    #
    # Access an Element by indices.
    #
    # @param [Fixnum] row the row index
    # @param [Fixnum] col the column index
    # @return [RubyExcel::Element]
    # @note Indexing is 1-based like Excel VBA
    #
    
    def cell( row, col )
      Element.new( self, indices_to_address( row, col ) )
    end
    alias cells cell
    
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
    
    # {Sheet#compact!}
    
    def compact
      dup.compact!
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
    # Check whether the Sheet contains data
    #
    # @return [Boolean] if there is any data
    #

    def empty?
      data.empty?
    end
    
    # {Sheet#filter!}
    
    def filter( header, &block )
      dup.filter!( header, &block )
    end

    #
    # Removes all Rows (omitting headers) where the block is false
    #
    # @param [String] header the header of the Column to pass to the block
    # @yield [Object] the value at the intersection of Column and Row
    # @return [self]
    #
    
    def filter!( header, &block )
      data.filter!( header, &block ); self
    end
    
    # {Sheet#get_columns!}
    
    def get_columns( *headers )
      dup.get_columns!( *headers )
    end
    alias gc get_columns
    
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
      row_id( column_by_header( header ).find( &block ) )
    end
    
    #
    # The highest currently used row number
    #
    
    def maxrow
      data.rows
    end
    
    #
    # The highest currently used column number
    #
    
    def maxcol
      data.cols
    end
    alias maxcolumn maxcol
    
    #
    # Allow shorthand range references
    #
    
    def method_missing(m, *args, &block)
      method_name = m.to_s.upcase.strip
      if method_name =~ /\A[A-Z]{1,3}\d+=?\z/
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
    
    def respond_to?(meth)
      if meth.to_s.upcase.strip =~ /\A[A-Z]{1,3}\d+=?\z/
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
    # Access an Element by address.
    #
    # @param [String, Element] first_cell the first Cell or Address in the Range
    # @param [String, Element] last_cell the last Cell or Address in the Range
    # @return [RubyExcel::Element]
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
      Element.new( self, to_range_address( first_cell, last_cell ) )
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
    
    # {Sheet#sort!}
    
    def sort( &block )
      dup.sort!( &block )
    end
    
    #
    # Sort the data according to a block (avoiding headers)
    #
    
    def sort!( &block )
      data.sort!( &block ); self
    end
    
    # {Sheet#sort_by!}
    
    def sort_by( &block )
      dup.sort_by!( &block )
    end
    
    #
    # Sort the data by the block value (avoiding headers)
    #
    
    def sort_by!( &block )
      data.sort_by!( &block ); self
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
    # Summarise the values of a Column into a Hash
    #
    # @param [String] header the header of the Column to summarise
    # @return [Hash]
    #
    
    def summarise( header )
      ch( header ).summarise
    end
    
    #
    # The Sheet as a 2D Array
    #
    
    def to_a
      data.all
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
      "<table>\n" + data.map { |row| '<tr>' + row.map { |v| '<td>' + CGI.escapeHTML(v.to_s) }.join() + "\n" }.join() + '</table>'
    end
    
    #
    # The Sheet as a Tab Seperated Value String
    #
    
    def to_s
      data.map { |ar| ar.map { |v| v.to_s.gsub(/\t|\n/,' ') }.join "\t" }.join( $/ )
    end
    
    # {Sheet#uniq!}
    
    def uniq( header )
      dup.uniq!( header )
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
    
    private
    
    def _hash_to_a(h)
      h.map { |k,v| v.is_a?(Hash) ? _hash_to_a(v).map { |val| ([ k ] + [ val ]).flatten(1) } : [ k, v ] }.flatten(1)
    end

    def _convert_hash(h)
      _hash_to_a(h).each_slice(2).map { |a1,a2| a1 << a2.last }
    end

  end # Sheet
  
end # RubyExcel