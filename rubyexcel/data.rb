module RubyExcel

  class Data
    attr_reader :rows, :cols
    attr_accessor :sheet
    alias parent sheet
    
    include Address
    
    def initialize( sheet, input_data )
      @sheet = sheet
      ( input_data.kind_of?( Array ) &&  input_data.all? { |el| el.kind_of?( Array ) } ) or fail ArgumentError, 'Input must be Array of Arrays'
      @data = input_data.dup
      calc_dimensions
    end
    
    def all
      @data.dup
    end
    
    def append( multi_array )
      @data
      @data += multi_array
      calc_dimensions
    end
    
    def colref_by_header( header )
      sheet.header_rows > 0 or fail NoMethodError, 'No header rows present'
      @data[ 0..sheet.header_rows-1 ].each do |r|
        if ( idx = r.index( header ) )
          return col_letter( idx+1 ) 
        end
      end
      fail IndexError, "#{ header } is not a valid header"
    end
    
    def compact!
      compact_columns!
      compact_rows!
    end
    
    def compact_columns!
      ensure_shape
      @data = @data.transpose.delete_if { |ar| ar.all? { |el| el.to_s.empty? } || ar.empty? }.transpose
      calc_dimensions
      @data
    end
    
    def compact_rows!
      @data.delete_if { |ar| ar.all? { |el| el.to_s.empty? } || ar.empty? }
      calc_dimensions
      @data
    end
    
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
        fail NoMethodError, "#{ object.class } is not supported"
      end
      self
    end
    
    def delete_column( ref )
      delete( Column.new( sheet, ref ) )
    end
  
    def delete_row( ref )
      delete( Row.new( sheet, ref ) )
    end
    
    def delete_range( ref )
      delete( Element.new( sheet, ref ) )
    end
    
    def dup
      Data.new( sheet, @data.dup )
    end
    
    def empty?
      no_headers.empty?
    end

    def filter!( header )
      hrows = sheet.header_rows
      idx = col_index( hrows > 0 ? colref_by_header( header ) : header )
      @data = @data.select.with_index { |row, i| hrows > i || yield( row[ idx -1 ] ) }
      calc_dimensions
      self
    end
  
    def get_columns!( *headers )
      hrow = sheet.header_rows - 1
      ensure_shape
      @data = @data.transpose.select{ |col| headers.include?( col[hrow] ) }
      ensure_shape
      @data = @data.sort_by{ |col| headers.index( col[hrow] ) || col[hrow] }.transpose
      calc_dimensions
      self
    end
        
    def insert_columns( before, number=1 )
      a = Array.new( number, nil )
      before = col_index( before ) - 1
      @data.map! { |row|  row.insert( before, *a ) }
    end
    
    def insert_rows( before, number=1 )
      @data = @data.insert( ( col_index( before ) - 1 ), *Array.new( number, [nil] ) )
    end
    
    def no_headers
      if sheet.header_cols.zero?
        @data[ sheet.header_rows..-1 ].dup
      else
        @data[ sheet.header_rows..-1 ].map { |row| row[ sheet.header_cols..-1 ] }
      end
    end
    
    def read( addr )
      row_idx, col_idx = address_to_indices( addr )
      @data[ row_idx-1 ][ col_idx-1 ]
    end
    alias [] read
    
    def uniq!( header )
      column = col_index( colref_by_header( header ) )
      @data = @data.uniq { |row| row[ column - 1 ] }
      self
    end
    alias unique! uniq!
    
    def write( addr, val )
      row_idx, col_idx = address_to_indices( addr )
      ( row_idx - rows ).times { @data << [] }
      @data[ row_idx-1 ][ col_idx-1 ] = val
      calc_dimensions
    end
    alias []= write

    include Enumerable
    
    def each
      @data.each { |ar| yield ar }
    end
    
    private
    
    def calc_dimensions
      @rows, @cols = @data.length, @data.max_by(&:length).length
    end
    
    def ensure_shape
      calc_dimensions
      @data.map! { |ar| ar.length == cols ? ar : ar + Array.new( cols - ar.length, nil) }
      @data
    end
    
  end

end