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
      address.upcase.match( /([A-Z]+)(\d+):([A-Z]+)(\d+)/i )
      start_col, end_col, start_row, end_row = [ $1, $3 ].sort + [ $2.to_i, $4.to_i ].sort
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
      ( col_letter( address_to_col_index( address ) + col ) ) + ( row_id( address ) + row ).to_s
    end
    
    def to_range_address( obj1, obj2 )
      if obj2
        obj2.is_a?( String ) ? ( obj1 + ':' + obj2 ) : "#{ obj1.address }:#{ obj2.address }"
      else
        obj1.is_a?( String ) ? obj1 : obj1.address
      end
    end
    
    def row_id( address )
      return nil unless address
      address[/\d+/].to_i
    end
    
  end
  
end