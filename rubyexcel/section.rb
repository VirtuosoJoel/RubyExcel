module RubyExcel

  #
  # Superclass for Row and Column
  #

  class Section
    include Address
    include Enumerable

    # The Sheet parent of the Section
    attr_reader :sheet
    alias parent sheet

    # The Data underlying the Sheet
    attr_reader :data
    
    #
    # Creates a RubyExcel::Section instance
    #
    # @param [RubyExcel::Sheet] sheet the parent Sheet
    #
    
    def initialize( sheet )
      @sheet = sheet
      @data = sheet.data
    end
    
    #
    # Append a value to the Section. 
    #   This only adds an extra cell if it is the first Row / Column.
    #   This prevents a loop through Rows or Columns from extending diagonally away from the main data.
    #
    # @param [Object] value the object to append
    #
  
    def <<( value )
      case self
      when Row ; lastone = ( col_index( idx ) == 1 ? data.cols + 1 : data.cols )
      else     ; lastone = ( col_index( idx ) == 1 ? data.rows + 1 : data.rows )
      end
      data[ translate_address( lastone ) ] = value
    end
    
    #
    # Access a cell by its index within the Section
    #
    
    def cell( ref )
      Element.new( sheet, translate_address( ref ) )
    end
  
    #
    # Delete the data referenced by self
    #
  
    def delete
      data.delete( self ); self
    end
  
    #
    # Check whether the data in self is empty
    #
  
    def empty?
      all? { |val| val.to_s.empty? }
    end
    
    #
    # Return the address of a given value
    #
    # @yield [Object] yields each cell value to the block
    # @return [String, nil] the address of the value or nil
    #
  
    def find
      return to_enum( :find ) unless block_given?
      each_cell { |ce| return ce.address if yield ce.value }; nil
    end
  
    #
    # View the object for debugging
    #
  
    def inspect
      "#{ self.class }:0x#{ '%x' % (object_id << 1) }: #{ idx }"
    end
  
    #
    # Read a value by address
    #
    # @param [String, Fixnum] id the index or reference of the required value
    #
  
    def read( id )
      data[ translate_address( id ) ]
    end
    alias [] read
    
    #
    # Summarise the values of a Section into a Hash
    #
    # @return [Hash]
    #
    
    def summarise
      each_wh.inject( Hash.new(0) ) { |h, v| h[v]+=1; h }
    end
    alias summarize summarise
    
    #
    # The Section as a seperated value String
    #
    
    def to_s
      to_a.join ( self.is_a?( Row ) ? "\t" : "\n" )
    end

    #
    # Write a value by address
    #
    # @param [String, Fixnum] id the index or reference to write to
    # @param [Object] val the object to place at the address
    #
    
    def write( id, val )
      data[ translate_address( id ) ] = val
    end
    alias []= write

    #
    # Yields each value
    #

    def each
      return to_enum(:each) unless block_given?
      each_address { |addr| yield data[ addr ] }
    end
    
    #
    # Yields each value, skipping headers
    #
    
    def each_without_headers
      return to_enum( :each_without_headers ) unless block_given?
      each_address_without_headers { |addr| yield data[ addr ] }
    end
    alias each_wh each_without_headers
    
    #
    # Yields each cell
    #
    
    def each_cell
      return to_enum( :each_cell ) unless block_given?
      each_address { |addr| yield Element.new( sheet, addr ) }
    end
    
    #
    # Yields each cell, skipping headers
    #
    
    def each_cell_without_headers
      return to_enum( :each_cell_without_headers ) unless block_given?
      each_address { |addr| yield Element.new( sheet, addr ) }
    end
    alias each_cell_wh each_cell_without_headers
    
    #
    # Replaces each value with the result of the block
    #
    
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

  #
  # A Row in the Sheet
  #
  
  class Row < Section
  
    # The Row index
    attr_reader :idx
    
    #
    # Creates a RubyExcel::Row instance
    #
    # @param [RubyExcel::Sheet] sheet the Sheet which holds this Row
    # @param [Fixnum] idx the index of this Row
    #

    def initialize( sheet, idx )
      @idx = idx.to_i
      super( sheet )
    end
    
    #
    # Access a cell by its header
    #
    # @param [String] header the header to search for
    # @return [RubyExcel::Element] the cell
    #
  
    def cell_by_header( header )
      cell( getref( header ) )
    end
    alias cell_h cell_by_header
  
    #
    # Find the Address of a header
    #
    # @param [String] header the header to search for
    # @return [String] the address of the header
    #
  
    def getref( header )
      sheet.header_rows.times do |t|
        res = sheet.row( t + 1 ).find &/^#{header}$/
        return column_id( res ) if res
      end
      fail ArgumentError, 'Invalid header: ' + header.to_s
    end
    
    #
    # Find a value in this Row by its header
    #
    # @param [String]header the header to search for
    # @return [Object] the value at the address
    #
    
    def value_by_header( header )
      self[ getref( header ) ]
    end
    alias val value_by_header

    private

    def each_address
      ( 'A'..col_letter( data.cols ) ).each { |col_id| yield "#{col_id}#{idx}" }
    end
    
    def each_address_without_headers
      ( 'A'..col_letter( data.cols ) ).each { |col_id| yield "#{col_id}#{idx}" }
    end

  end

  #
  # A Column in the Sheet
  #
  
  class Column < Section
  
    # The Row index
    attr_reader :idx
    
    #
    # Creates a RubyExcel::Column instance
    #
    # @param [RubyExcel::Sheet] sheet the Sheet which holds this Column
    # @param [String, Fixnum] idx the index of this Column
    #
    
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