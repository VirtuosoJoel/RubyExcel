module RubyExcel

  def self.sample_data
    a=[];8.times{|t|b=[];c='A';5.times{b<<"#{c}#{t+1}";c.next!};a<<b};a
  end

  def <<( other )
    case other
    when RubyExcel
      other.each { |s| @sheets << s }
    when Sheet
      @sheets << other
    end
  end

  class Sheet
  
    def +( other )
      dup << other
    end
    
    def -( other )
      case other
      when Array
        Workbook.new.load( data.all - other )
      when Sheet
        Workbook.new.load( data.all - other.data.no_headers )
      else
        fail ArgumentError, "Unsupported class: #{ other.class }"
      end
    end
    
    def <<( other )
      case other
      when Array
        load( data.all + other, header_rows, header_cols )
      when Sheet
        load( data.all + other.data.no_headers, header_rows, header_cols )
      else
        fail ArgumentError, "Unsupported class: #{ other.class }"
      end
    end
    
    def compact!
      data.compact!
      self
    end
    
    def empty?
      data.empty?
    end
    
    def column_by_header( header )
      Column.new( self, data.colrf_by_header( header ) )
    end
    alias ch column_by_header

    def filter!( ref, &block )
      data.filter!( ref, &block )
      self
    end
    
  end
  
  class Data
  
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
      object
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
  
    def filter!( header )
      hrows = sheet.header_rows
      idx = col_index( hrows > 0 ? colref_by_header( header ) : header )
      @data = @data.select.with_index { |row, i| hrows > i || yield( row[ idx -1 ] ) }
    end
  
    def get_columns!( *headers )
      hrow = sheet.header_rows - 1
      ensure_shape
      @data = @data.transpose.select{ |col| headers.include?( col[hrow] ) }
      ensure_shape
      @data = @data.sort_by{ |col| headers.index( col[hrow] ) || col[hrow] }.transpose
    end

  end
  
  class Section

  end

  class Row < Section

  end

  class Column < Section

  end

  class Element

  end

end
