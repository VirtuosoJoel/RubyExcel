module RubyExcel

  class Section
  
    include Address

    attr_reader :sheet, :idx, :data
    alias parent sheet

    def initialize( sheet )
      @sheet = sheet
      @data = sheet.data
    end
  
    def <<( value )
      case self
      when Row ; lastone = ( col_index( idx ) == 1 ? data.cols + 1 : data.cols )
      else     ; lastone = ( col_index( idx ) == 1 ? data.rows + 1 : data.rows )
      end
      data[ translate_address( lastone ) ] = value
    end
    
    def cell( ref )
      Element.new( sheet, translate_address( ref ) )
    end
  
    def delete
      data.delete( self )
    end
  
    def empty?
      all? { |val| val.to_s.empty? }
    end
  
    def find
      return to_enum( :find ) unless block_given?
      each_cell { |ce| return ce.address if yield ce.value }; nil
    end
  
    def inspect
      "#{ self.class }:0x#{ '%x' % (object_id << 1) }: #{ idx }"
    end
  
    def read( id )
      data[ translate_address( id ) ]
    end
    alias [] read
    
    def summarise
      each_wh.inject( Hash.new(0) ) { |h, v| h[v]+=1; h }
    end
    alias summarize summarise
    
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
    
    def each_without_headers
      return to_enum( :each_without_headers ) unless block_given?
      each_address_without_headers { |addr| yield data[ addr ] }
    end
    alias each_wh each_without_headers
    
    def each_cell
      return to_enum( :each_cell ) unless block_given?
      each_address { |addr| yield Element.new( sheet, addr ) }
    end
    
    def each_cell_without_headers
      return to_enum( :each_cell_without_headers ) unless block_given?
      each_address { |addr| yield Element.new( sheet, addr ) }
    end
    alias each_cell_wh each_cell_without_headers
    
    def map!
      return to_enum( :map! ) unless block_given?
      each_address { |addr| data[addr] = ( yield data[addr] ) }
    end

    private
    
    def translate_address( addr )
      case self
      when Row
        col_letter( addr ) + idx.to_s
      when Column
        addr = addr.to_s unless addr.is_a?( String )
        fail ArgumentError, "Invalid address : #{ addr }" if addr =~ /[^\d]/
        idx + addr
      end
    end

  end

  class Row < Section
  
    attr_reader :idx

    def initialize( sheet, idx )
      @idx = idx.to_i
      super( sheet )
    end
  
    def cell_by_header( header )
      cell( getref( header ) )
    end
    alias cell_h cell_by_header
  
    def getref( header )
      column_id( sheet.row(1).find &/#{header}/ )
    end
    
    def value_by_header( header )
      self[ getref( header ) ]
    end
    alias val value_by_header

    private

    def each_address
      ( 'A'..col_letter( data.cols ) ).each { |col_id| yield "#{col_id}#{idx}" }
    end
    
    def each_address_without_headers
      ( col_letter( sheet.header_cols+1 )..col_letter( data.cols ) ).each { |col_id| yield "#{col_id}#{idx}" }
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
      ( 1..data.rows ).each { |row_id| yield idx + row_id.to_s }
    end

    def each_address_without_headers
      ( sheet.header_rows+1 ).upto( data.rows ) { |row_id| yield idx + row_id.to_s }
    end

  end

end