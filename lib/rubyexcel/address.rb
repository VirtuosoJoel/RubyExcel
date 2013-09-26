module RubyExcel

  #
  #Provides address translation methods to RubyExcel's classes
  #
  
  module Address
  
    #
    # Translates an address to a column index
    #
    # @param [String] address the address to translate
    # @return [Fixnum] the column index
    #
    
    def address_to_col_index( address )
      col_index( column_id( address ) )
    end
    
    #
    # Translates an address to indices
    #
    # @param [String] address the address to translate
    # @return [Array<Fixnum>] row index, column index
    #
    
    def address_to_indices( address )
      [ row_id( address ), address_to_col_index( address ) ]
    end
    
    #
    # Translates a column id to an index
    #
    # @param [String] letter the column id to translate
    # @return [Fixnum] the corresponding index
    #
    
    def col_index( letter )
      return letter if letter.is_a? Fixnum
      letter !~ /[^A-Z]/ && [1,2,3].include?( letter.length ) or fail ArgumentError, "Invalid column reference: #{ letter }"
      idx, a = 1, 'A'
      loop { return idx if a == letter; idx+=1; a.next! }
    end
  
    #
    # Translates an index to a column letter
    #
    # @param [Fixnum] index the index to translate
    # @return [String] the column letter
    #
  
    def col_letter( index, start='A' )
      return index if index.is_a? String
      index > 0 or fail ArgumentError, 'Indexing is 1-based'
      a = start.dup; ( index - 1 ).times { a.next! }; a
    end
  
    #
    # Translates an address to a column id
    #
    # @param [String] address the address to translate
    # @return [String] the column id
    #
  
    def column_id( address )
      address[/[A-Z]+/i].upcase
    end
    
    #
    # Expands an address to all contained addresses
    #
    # @param [String] address the address to translate
    # @return [Array<String>] all addresses included within the given address
    #

    def expand( address )
      return [[address]] unless address.include? ':'
      
      #Extract the relevant boundaries
      case address
      
      # Row
      when /\A(\d+):(\d+)\z/
      
        start_col, end_col, start_row, end_row = [ 'A', col_letter( sheet.maxcol ) ] + [ $1.to_i, $2.to_i ].sort
        
      # Column
      when /\A([A-Z]+):([A-Z]+)\z/
      
        start_col, end_col, start_row, end_row = [ $1, $2 ].sort + [ 1, sheet.maxrow ]
        
      # Range
      when /([A-Z]+)(\d+):([A-Z]+)(\d+)/
      
        start_col, end_col, start_row, end_row = [ $1, $3 ].sort + [ $2.to_i, $4.to_i ].sort
        
      # Invalid
      else
        fail ArgumentError, 'Invalid address: ' + address
      end
      
      # Return the array of addresses
      ( start_row..end_row ).map { |r| ( start_col..end_col ).map { |c| c + r.to_s } } 
      
    end
    
    #
    # Translates indices to an address
    #
    # @param [Fixnum] row_idx the row index
    # @param [Fixnum] column_idx the column index
    # @return [String] the corresponding address
    #
    
    def indices_to_address( row_idx, column_idx )
      [ row_idx, column_idx ].all? { |a| a.is_a?( Fixnum ) } or fail ArgumentError, 'Input must be Fixnum'
      col_letter( column_idx ) + row_idx.to_s
    end
    
    #
    # Checks whether an object is a multidimensional Array
    #
    # @param [Object] obj the object to test
    # @return [Boolean] whether the object is a multidimensional Array
    #
    
    def multi_array?( obj )
      obj.all? { |el| el.is_a?( Array ) } && obj.is_a?( Array ) rescue false
    end
    
    #
    # Offsets an address by row and column
    #
    # @param [String] address the address to offset
    # @param [Fixnum] row the number of rows to offset by
    # @param [Fixnum] col the number of columns to offset by
    # @return [String] the new address
    #
    
    def offset(address, row, col)
      ( col_letter( address_to_col_index( address ) + col ) ) + ( row_id( address ) + row ).to_s
    end
    
    #
    # Translates an address to a row id
    #
    # @param [String] address the address to translate
    # @return [Fixnum] the row id
    #
    
    def row_id( address )
      Integer( address[/\d+/] )
    end
    
    #
    # Step an index forward for an Array-style slice
    #
    # @param [Fixnum, String] start the index to start at
    # @param [Fixnum] slice the amount to advance to (1 means keep the same index)
    #
    
    def step_index( start, slice )
      if start.is_a?( Fixnum )
        start + slice - 1
      else
        x = start.dup
        ( slice - 1 ).times { x.next! }
        x
      end
    end
    
    #
    # Translates two objects to a range address
    #
    # @param [String, RubyExcel::Element] obj1 the first address element
    # @param [String, RubyExcel::Element] obj2 the second address element
    # @return [String] the new address
    #
    
    def to_range_address( obj1, obj2 )
      addr = obj1.respond_to?( :address ) ? obj1.address : obj1.to_s
      addr << ':' + ( obj2.respond_to?( :address ) ? obj2.address : obj2.to_s ) if obj2
      addr
    end

  end    

end