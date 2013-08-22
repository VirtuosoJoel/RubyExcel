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
    # Access a cell by its index within the Section
    #
    
    def cell( ref )
      Cell.new( sheet, translate_address( ref ) )
    end
  
    #
    # Delete the data referenced by self
    #
  
    def delete
      data.delete( self ); self
    end
  
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
      each_address( false ) { |addr| yield data[ addr ] }
    end
    alias each_wh each_without_headers
    
    #
    # Yields each cell
    #
    
    def each_cell
      return to_enum( :each_cell ) unless block_given?
      each_address { |addr| yield Cell.new( sheet, addr ) }
    end
    
    #
    # Yields each cell, skipping headers
    #
    
    def each_cell_without_headers
      return to_enum( :each_cell_without_headers ) unless block_given?
      each_address( false ) { |addr| yield Cell.new( sheet, addr ) }
    end
    alias each_cell_wh each_cell_without_headers
  
    #
    # Check whether the data in self is empty
    #
  
    def empty?
      each_wh.all? { |val| val.to_s.empty? }
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
    # Return the value of the last cell
    #
  
    def last
      last_cell.value
    end
  
    #
    # Return the last cell
    #
    # @return [RubyExcel::Cell]
    #
  
    def last_cell
      Cell.new( sheet, each_address.to_a.last )
    end
  
    #
    # Replaces each value with the result of the block
    #
    
    def map!
      return to_enum( :map! ) unless block_given?
      each_address { |addr| data[addr] = ( yield data[addr] ) }
    end
    
    #
    # Replaces each value with the result of the block, skipping headers
    #
  
    def map_without_headers!
      return to_enum( :map_without_headers! ) unless block_given?
      each_address( false ) { |addr| data[addr] = ( yield data[addr] ) }
    end
    alias map_wh! map_without_headers!
  
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
      to_a.map { |v| v.to_s.gsub(/\t|\n|\r/,' ') }.join ( self.is_a?( Row ) ? "\t" : "\n" )
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

  end

  #
  # A Row in the Sheet
  # @attr_reader [Fixnum] idx the Row index
  # @attr_reader [Fixnum] length the Row length
  #
  
  class Row < Section
  
    # The Row number
    attr_reader :idx
    alias index idx
    
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
    # Append a value to the Row. 
    #
    # @param [Object] value the object to append
    # @note This only adds an extra cell if it is the first Row
    #       This prevents a loop through Rows from extending diagonally away from the main data.
    #
  
    def <<( value )
      data[ translate_address( idx == 1 ? data.cols + 1 : data.cols ) ] = value
    end
    
    #
    # Access a Cell by its header
    #
    # @param [String] header the header to search for
    # @return [RubyExcel::Cell] the cell
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
        res = sheet.row( t + 1 ).find { |v| v == header }
        return column_id( res ) if res
      end
      fail ArgumentError, 'Invalid header: ' + header.to_s
    end
    
    #
    # The number of Columns in the Row
    #
    
    def length
      data.cols
    end
    
    #
    # Find a value in this Row by its header
    #
    # @param [String] header the header to search for
    # @return [Object] the value at the address
    #
    
    def value_by_header( header )
      self[ getref( header ) ]
    end
    alias val value_by_header
    
    #
    # Set a value in this Row by its header
    #
    # @param [String] header the header to search for
    # @param [Object] val the value to write
    # 
    
    def set_value_by_header( header, val )
      self[ getref( header ) ] = val
    end
    alias set_val set_value_by_header

    private

    def each_address( unused=nil )
      return to_enum( :each_address ) unless block_given?
      ( 'A'..col_letter( data.cols ) ).each { |col_id| yield translate_address( col_id ) }
    end

    def translate_address( addr )
      col_letter( addr ) + idx.to_s
    end
    
  end

  #
  # A Column in the Sheet
  #
  # @attr_reader [String] idx the Column index
  # @attr_reader [Fixnum] length the Column length
  #
  
  class Column < Section
  
    # The Column letter
    attr_reader :idx
    alias index idx
    
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
    
    #
    # Append a value to the Column. 
    #
    # @param [Object] value the object to append
    # @note This only adds an extra cell if it is the first Column.
    #       This prevents a loop through Columns from extending diagonally away from the main data.
    #
  
    def <<( value )
      data[ translate_address( idx == 'A' ? data.rows + 1 : data.rows ) ] = value
    end
    
    #
    # The number of Rows in the Column
    #
    
    def length
      data.rows
    end

    private
    
    def each_address( headers=true )
      return to_enum( :each_address ) unless block_given?
      ( headers ? 1 : sheet.header_rows + 1 ).upto( data.rows ) { |row_id| yield translate_address( row_id ) }
    end

    def translate_address( addr )
      addr = addr.to_s unless addr.is_a?( String )
      fail ArgumentError, "Invalid address : #{ addr }" if addr =~ /[^\d]/
      idx + addr
    end
    
  end

end