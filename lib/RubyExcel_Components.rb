module RubyExcel

  module Address
  
    def address_to_col_index( address )
      col_index( column_id( address ) )
    end
    
    def address_to_indices( address )
      [ row_id( address ), address_to_col_index( address ) ]
    end
    
    def col_index( letter )
      return letter if letter.is_a? Fixnum
      letter !~ /[^A-Z]/ && [1,2,3].include?( letter.length ) or fail ArgumentError, "Invalid column reference: #{ letter }"
      idx, a = 1, 'A'
      loop { return idx if a == letter; idx+=1; a.next! }
    end
  
    def col_letter( index )
      return index if index.is_a? String
      index > 0 or fail ArgumentError, 'Indexing is 1-based'
      a = 'A' ; return a if index == 1
      (index - 1).times{ a.next! }; a
    end
  
    def column_id( address )
      address[/[A-Z]+/] or fail ArgumentError, "Invalid address: #{ address }"
    end
    
    def expand( address )
      return [[address]] unless address.include? ':'
      start_col, end_col, start_row, end_row = [ address[/^[A-Z]+/], address[/(?<=:)[A-Z]+/] ].sort + [ address.match(/(\d+):/).captures.first, address[/\d+$/] ].sort
      (start_row..end_row).map { |r| (start_col..end_col).map { |c| "#{ c }#{ r }" } }
    end
    
    def indices_to_address( row_idx, column_idx )
      [ row_idx, column_idx ].all? { |a| a.is_a?( Fixnum ) } or fail ArgumentError, 'Input must be Fixnum'
      col_letter( column_idx ) + row_idx.to_s
    end
    
    def multi_array?( obj )
      obj.all? { |el| el.is_a?( Array ) } && obj.is_a?( Array ) rescue false
    end
    
    def offset(address, row, col)
      ( col_letter( address_to_col_index( address ) ) + col ) + ( row_id( address ) + row ).to_s
    end
    
    def to_range_address( obj1, obj2 )
      if obj2
        obj2.is_a?( String ) ? ( obj1 + ':' + obj2 ) : "#{ obj1.address }:#{ obj2.address }"
      else
        obj1.is_a?( String ) ? obj1 : obj1.address
      end
    end
    
    def row_id( address )
      address[/\d+/].to_i
    end
    
  end

  class Data
    attr_reader :rows, :cols
    attr_accessor :sheet
    
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

    def dup
      Data.new( sheet, @data.dup )
    end
    
    def empty?
      no_headers.empty?
    end
    
    def insert_columns( before, number )
      a = Array.new( number, nil )
      before = col_index( before ) - 1
      @data.map! { |row|  row.insert( before, *a ) }
    end
    
    def insert_rows( before, number )
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

  class Section
  
    include Address

    attr_reader :sheet, :idx, :data

    def initialize( sheet )
      @sheet = sheet
      @data = sheet.data
    end
  
    def delete
      data.delete( self )
    end
  
    def inspect
      "#{ self.class }: #{ idx }"
    end
  
    def read( id )
      data[ translate_address( id ) ]
    end
    alias [] read
    
    def to_s
      to_a.join ( self.is_a?( Row ) ? "\t" : "\n" )
    end

    def write( id, val )
      data[ translate_address( id ) ] = val
    end
    alias []= write

    include Enumerable

    def each
      return to_enum(:each) unless block_given?
      each_address { |addr| yield data[ addr ] }
    end
    
    def map!
      each_address { |addr| data[addr] = ( yield data[addr] ) }
    end
    
    def empty?
      all? { |val| val.to_s.empty? }
    end
    
    private
    
    def translate_address( addr )
      case self
      when Row
        col_letter( addr ) + idx.to_s
      when Column
        idx + addr.to_s
      end
    end

  end

  class Row < Section
  
    attr_reader :idx

    def initialize( sheet, idx )
      @idx = idx.to_i
      super( sheet )
    end
  
    private

    def each_address
      ( 'A'..col_letter( data.cols ) ).each { |col_id| yield "#{col_id}#{idx}" }
    end

  end

  class Column < Section
  
    attr_reader :idx
    
    def initialize( sheet, idx )
      @idx = idx
      super( sheet )
    end

    private

    def each_address
      1.upto( data.rows ) { |row_id| yield idx + row_id.to_s }
    end

  end

  class Element

    attr_reader :sheet, :address, :data

    def initialize( sheet, addr )
      fail ArgumentError, "Invalid range: #{ addr }" unless addr =~ /\A[A-Z]+\d+:[A-Z]+\d+\z|\A[A-Z]+\d+\z/
      @sheet = sheet
      @data = sheet.data
      @address = addr
    end
  
    def delete
      data.delete( self )
    end
  
    include Address
  
    def value
      if address.include? ':'
        expand( address ).map { |ar| ar.map { |addr| data[ addr ] } }
      else
        data[ address ]
      end
    end
  
    def value=( val )
      if address.include? ':'
        if multi_array?( val )
          expand( address ).each_with_index { |row,idx| row.each_with_index { |el,i| data[ el ] = val[idx][i] } }
        else
          expand( address ).each { |ar| ar.each { |addr| data[ addr ] = val } }
        end
      else
        data[ address ] = val
      end
      self
    end
  
    def to_s
      value.is_a?( Array ) ? value.map { |ar| ar.join "\t" }.join($/) : value.to_s
    end
    
    def inspect
      "#{ self.class }: #{ address }"
    end
  
    include Enumerable
    
    def each
      expand( address ).flatten.each { |addr| yield data[ addr ] }
    end
  
    def map!
      expand( address ).flatten.each { |addr| data[ addr ] = yield data[ addr ] }
    end
  
  end

end