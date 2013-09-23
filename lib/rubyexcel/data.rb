module RubyExcel

require_relative 'address.rb'

  #
  # The class which holds a Sheet's data
  #
  # @note This class is exposed to the API purely for debugging.
  #

  class Data
    include Address
    include Enumerable
  
    #The number of rows in the data
    attr_reader :rows
    
    #The number of columns in the data
    attr_reader :cols
    
    #The parent Sheet
    attr_accessor :sheet
    alias parent sheet
    
    #
    # Creates a RubyExcel::Data instance
    #
    # @param [RubyExcel::Sheet] sheet the parent Sheet
    # @param [Array<Array>] input_data the multidimensional Array which holds the data
    #
    
    def initialize( sheet, input_data )
      ( input_data.kind_of?( Array ) &&  input_data.all? { |el| el.kind_of?( Array ) } ) or fail ArgumentError, 'Input must be Array of Arrays'
      @sheet = sheet
      @data = input_data.dup
      calc_dimensions
    end

    #
    # Append an object to Data
    #
    # @param [Object] other the data to append
    # @return [self]
    #    
    
    def <<( other )
      case other
      when Array
        if multi_array?( other )
          all.all?(&:empty?) ? @data = other : @data += other
        else
          all.all?(&:empty?) ? @data = [ other ] : @data << other
        end
      when Hash   ; @data += _convert_hash( other )
      when Sheet  ; empty? ? @data = other.data.all.dup : @data += other.data.dup.no_headers
      when Row    ; @data << other.to_a.dup
      when Column ; @data.map!.with_index { |row, i| row << other[ i+1 ] }
      else        ; @data[0] << other
      end
      calc_dimensions
      self
    end
    
    # @overload advanced_filter!( header, comparison_operator, search_criteria, ... )
    #   Filter on multiple criteria
    #
    # @example Filter to 'Part': 'Type1' and 'Type3', with 'Qty' greater than 1
    #   s.advanced_filter!( 'Part', :=~, /Type[13]/, 'Qty', :>, 1 )
    #
    # @example Filter to 'Part': 'Type1', with 'Ref1' containing 'X'
    #   s.advanced_filter!( 'Part', :==, 'Type1', 'Ref1', :include?, 'X' )
    #
    #   @param [String] header a header to search under
    #   @param [Symbol] comparison_operator the operator to compare with
    #   @param [Object] search_criteria the value to filter by
    #   @raise [ArgumentError] 'Number of arguments must be a multiple of 3'
    #   @raise [ArgumentError] 'Operator must be a symbol'
    #
    
    def advanced_filter!( *args )
      hrows = sheet.header_rows
      args.length % 3 == 0 or fail ArgumentError, 'Number of arguments must be a multiple of 3'
      1.step( args.length - 2, 3 ) { |i| args[i].is_a?( Symbol ) or fail ArgumentError, 'Operator must be a symbol: ' + args[i].to_s }
      0.step( args.length - 3, 3 ) { |i| index_by_header( args[i] ) }
      
      @data = @data.select.with_index do |row, i|
        if hrows > i
          true
        else
          args.each_slice(3).map do |h, op, crit|
            row[ index_by_header( h ) - 1 ].send( op, crit )
          end.all?
        end
      end
      calc_dimensions
    end
    
    #
    # Returns a copy of the data
    #
    # @return [Array<Array>]
    #
    
    def all
      @data.dup
    end
 
    #
    # Finds a Column reference by a header
    #
    # @param [String] header the header to search for
    # @return [String] the Column reference
    # @raise [NoMethodError] 'No header rows present'
    # @raise [IndexError] header.to_s + ' is not a valid header'
    #
    
    def colref_by_header( header )
      return header.idx if header.is_a?( Column )
      sheet.header_rows > 0 or fail NoMethodError, 'No header rows present'
      @data[ 0..sheet.header_rows-1 ].each { |r| idx = r.index( header ); return col_letter( idx+1 ) if idx }
      fail IndexError, header.to_s + ' is not a valid header'
    end
    
    #
    # Removes empty rows and columns from the data
    #
    
    def compact!
      compact_columns!
      compact_rows!
    end
    
    #
    # Removes empty columns from the data
    #
    
    def compact_columns!
      ensure_shape
      @data = @data.transpose.delete_if { |ar| ar.all? { |el| el.to_s.empty? } || ar.empty? }.transpose
      calc_dimensions
    end
    
    #
    # Removes empty rows from the data
    #
    
    def compact_rows!
      @data.delete_if { |ar| ar.all? { |el| el.to_s.empty? } || ar.empty? }
      calc_dimensions
    end
    
    #
    # Deletes the data referenced by an object
    #
    # @param [RubyExcel::Column, RubyExcel::Element, RubyExcel::Row] object the object to delete
    # @raise [NoMethodError] object.class.to_s + ' is not supported"
    #
    
    def delete( object )
      case object
      when Row
        @data.slice!( object.idx - 1 )
      when Column
        idx = col_index( object.idx ) - 1
        @data.each { |r| r.slice! idx }
      when Element
        addresses = expand( object.address )
        indices = [ address_to_indices( addresses.first.first ), address_to_indices( addresses.last.last ) ].flatten.map { |n| n-1 }
        @data[ indices[0]..indices[2] ].each { |r| r.slice!( indices[1], indices[3] - indices[1] + 1 ) }
        @data.delete_if.with_index { |r,i| r.empty? && i.between?( indices[0], indices[2] ) }
      else
        fail NoMethodError, object.class.to_s + ' is not supported'
      end
      calc_dimensions
    end
    
    #
    # Wipe all data
    #
    
    def delete_all
      @data = [[]]
    end
    
    #
    # Deletes the data referenced by a column id
    #
    
    def delete_column( ref )
      delete( Column.new( sheet, ref ) )
    end
  
    #
    # Deletes the data referenced by a row id
    #
  
    def delete_row( ref )
      delete( Row.new( sheet, ref ) )
    end
    
    #
    # Deletes the data referenced by an address
    #
    
    def delete_range( ref )
      delete( Element.new( sheet, ref ) )
    end
    
    #
    # Return a copy of self
    #
    # @return [RubyExcel::Data]
    #
    
    def dup
      Data.new( sheet, @data.map(&:dup) )
    end
    
    #
    # Check whether the data (without headers) is empty
    #
    # @return [Boolean]
    #
    
    def empty?
      no_headers.empty? rescue true
    end

    #
    # Yields each "Row" as an Array
    #
    
    def each
      return to_enum( :each ) unless block_given?
      @data.each { |ar| yield ar }
    end

    #
    # Removes all Rows (omitting headers) where the block is false
    #
    # @param [String] header the header of the Column to pass to the block
    # @yield [Object] the value at the intersection of Column and Row
    # @return [self]
    #

    def filter!( header )
      hrows = sheet.header_rows
      idx = index_by_header( header )
      @data = @data.select.with_index { |row, i| hrows > i || yield( row[ idx -1 ] ) }
      calc_dimensions
    end
  
    #
    # Select and re-order Columns by a list of headers
    #
    # @param [Array<String>] headers the ordered list of headers to keep
    # @note This method can accept either a list of arguments or an Array
    # @note Invalid headers will be skipped
    #
  
    def get_columns!( *headers )
      headers = headers.flatten
      hrow = sheet.header_rows - 1
      ensure_shape
      @data = @data.transpose.select{ |col| col[0..hrow].any?{ |val| headers.include?( val ) } }
      @data = @data.sort_by{ |col| headers.index( col[0..hrow].select { |val| headers.include?( val ) }.first ) || headers.length }.transpose
      calc_dimensions
    end
    
    #
    # Return the header section of the data
    #
    
    def headers
      return nil if sheet.header_rows.nil? || sheet.header_rows.zero?
       @data[ 0..sheet.header_rows-1 ]
    end
    
    #
    # Find a Column index by header
    #
    # @param [String] header the Column header to search for
    # @return [Fixnum] the index of the given header
    #
    
    def index_by_header( header )
      sheet.header_rows > 0 or fail NoMethodError, 'No header rows present'
      col_index(  colref_by_header( header ) )
    end
    
    #
    # Insert blank Columns into the data
    #
    # @param [String, Fixnum] before the Column reference to insert before.
    # @param [Fixnum] number the number of new Columns to insert
    #
    
    def insert_columns( before, number=1 )
      a = Array.new( number, nil )
      before = col_index( before ) - 1
      @data.map! { |row|  row.insert( before, *a ) }
      calc_dimensions
    end
    
    #
    # Insert blank Rows into the data
    #
    # @param [Fixnum] before the Row index to insert before.
    # @param [Fixnum] number the number of new Rows to insert
    #
    
    def insert_rows( before, number=1 )
      @data = @data.insert( ( col_index( before ) - 1 ), *Array.new( number, [nil] ) )
      calc_dimensions
    end
    
    #
    # Return the data without headers
    #
    
    def no_headers
      return @data unless sheet.header_rows
      @data[ sheet.header_rows..-1 ]
    end
    
    #
    # Split the data into two sections by evaluating each value in a column
    #
    # @param [String] header the header of the Column which contains the yield value
    # @yield [value] yields the value of each row under the given header
    #
    
    def partition( header, &block )
      copy = dup
      idx = index_by_header( header )
      d1, d2 = copy.no_headers.partition { |row| yield row[ idx -1 ] }
      [ copy.headers + d1, copy.headers.map(&:dup) + d2 ] if headers
    end
    
    #
    # Read a value by address
    #
    
    def read( addr )
      row_idx, col_idx = address_to_indices( addr )
      return nil if row_idx > rows
      @data[ row_idx-1 ][ col_idx-1 ]
    end
    alias [] read
    
    #
    # Reverse the data Columns
    #
    
    def reverse_columns!
      ensure_shape
      @data = @data.transpose.reverse.transpose
    end

    #
    # Reverse the data Rows (without affecting the headers)
    #
    
    def reverse_rows!
      @data = skip_headers &:reverse
    end
    
    #
    # Perform an operation on the data without affecting the headers
    #
    # @yield [data] yield the data without the headers
    # @return [Array<Array>] returns the data with the block operation performed on it, and the headers back in place
    #

    def skip_headers
      return to_enum(:skip_headers) unless block_given?
      hr = sheet.header_rows
      if hr > 0
        @data[ 0..hr - 1 ] + yield( @data[ hr..-1 ] )
      else
        yield( @data )
      end 
    end
    
    #
    # Sort the data according to the block
    #
    
    def sort!( &block )
      @data = skip_headers { |d| d.sort( &block ) }; self
    end
    
    #
    # Sort the data according to the block value
    #
    
    def sort_by!( &block )
      @data = skip_headers { |d| d.sort_by( &block ) }; self
    end
    
    #
    # Unique the rows according to the values within a Column, selected by header
    #
    
    def uniq!( header )
      column = col_index( colref_by_header( header ) )
      @data = skip_headers { |d| d.uniq { |row| row[ column - 1 ] } }
      calc_dimensions
    end
    alias unique! uniq!
    
    #
    # Write a value into the data
    #
    # @param [String] addr the address to write the value to
    # @param val the value to write to the address
    #
    
    def write( addr, val )
      row_idx, col_idx = address_to_indices( addr )
      ( row_idx - rows ).times { @data << [] }
      @data[ row_idx-1 ][ col_idx-1 ] = val
      calc_dimensions if row_idx > rows || col_idx > cols
      val
    end
    alias []= write
    
    private
    
    def calc_dimensions
      @rows = ( @data.length rescue 0 )
      @cols = ( @data.max_by { |row| row.length }.length rescue 0 )
      self
    end
    
    def ensure_shape
      calc_dimensions
      @data = @data.map { |ar| ar.length == cols ? ar : ar + Array.new( cols - ar.length, nil) }
    end
    
    def _convert_hash(h)
      _hash_to_a(h).each_slice(2).map { |a1,a2| a1 << a2.last }
    end
    
    def _hash_to_a(h)
      h.map { |k,v| v.is_a?(Hash) ? _hash_to_a(v).map { |val| ([ k ] + [ val ]).flatten(1) } : [ k, v ] }.flatten(1)
    end
    
  end

end