module RubyExcel

  #
  # A Range or Cell in a Sheet
  #

  class Element
    include Address
    include Enumerable

    #The parent Sheet
    attr_reader :sheet
    alias parent sheet
    
    #The address
    attr_reader :address
    
    #The Data underlying the Sheet
    attr_reader :data
    
    #The first Column id in the address
    attr_reader :column
    
    #The first Row id in the address
    attr_reader :row
    
    #
    # Creates a RubyExcel::Element instance
    #
    # @param [RubyExcel::Sheet] sheet the parent Sheet
    # @param [String] addr the address to reference
    #
    
    def initialize( sheet, addr )
      fail ArgumentError, "Invalid range: #{ addr }" unless addr =~ /\A[A-Z]{1,3}\d+:[A-Z]{1,3}\d+\z|\A[A-Z]{1,3}\d+\z|\A[A-Z]{1,3}:[A-Z]{1,3}\z|\A\d+:\d+\z/
      @sheet = sheet
      @data = sheet.data
      @address = addr
      @column = column_id( addr )
      @row = row_id( addr )
    end
    
    #
    # Delete the data referenced by self.address
    #
  
    def delete
      data.delete( self ); self
    end
    
    #
    # Return the value at this Element's address
    #
    # @return [Object, Array<Object>] the Object or Array of Objects within the data, referenced by the address
    #
  
    def value
      address.include?( ':' ) ? expand( address ).map { |ar| ar.map { |addr| data[ addr ] } } : data[ address ]
    end
  
    #
    # Set the value at this Element's address
    #
    # @param [Object, Array<Object>] val the Object or Array of Objects to write into the data
    #
  
    def value=( val )
    
      #Range
      if address.include? ':'
      
        addresses = expand( address )
        
        # 2D Array of Values
        if multi_array?( val )
        
          #Check the dimensions
          val_rows, val_cols, range_rows, range_cols = val.length, val.max_by(&:length).length, addresses.length, addresses.max_by(&:length).length
          val_rows == range_rows && val_cols == range_cols or fail ArgumentError, "Dimension mismatch! Value - rows: #{val_rows}, columns: #{ val_cols }. Range - rows: #{ range_rows }, columns: #{ range_cols }"
          
          #Write the values in order
          addresses.each_with_index { |row,idx| row.each_with_index { |el,i| data[ el ] = val[idx][i] } }
          
        # Single Value
        else
        
          #Write the same value to every cell in the Range
          addresses.each { |ar| ar.each { |addr| data[ addr ] = val } }
          
        end
      
      #Cell
      else
        data[ address ] = val
      end
      
      self
    end
    
    #
    # The data at address as a TSV String
    #
  
    def to_s
      value.is_a?( Array ) ? value.map { |ar| ar.join "\t" }.join($/) : value.to_s
    end
    
    #
    # View the object for debugging
    #
    
    def inspect
      "#{ self.class }:0x#{ '%x' % ( object_id << 1 ) }: #{ address }"
    end
  
    #
    # Yields each value in the data referenced by the address
    #
    
    def each
      return to_enum( :each ) unless block_given?
      expand( address ).flatten.each { |addr| yield data[ addr ] }
    end
    
    #
    # Yields each Element referenced by the address
    #
    
    def each_cell
      return to_enum( :each_cell ) unless block_given?
      expand( address ).flatten.each { |addr| yield Element.new( sheet, addr ) }
    end
    
    #
    # Checks whether the data referenced by the address is empty
    #
    
    def empty?
      all? { |v| v.to_s.empty? }
    end
  
    #
    # Replaces each value with the result of the block
    #
  
    def map!
      return to_enum( :map! ) unless block_given?
      expand( address ).flatten.each { |addr| data[ addr ] = yield data[ addr ] }
    end
  
  end

end