module RubyExcel

  class Element

    attr_reader :sheet, :address, :data, :column, :row
    alias parent sheet

    def initialize( sheet, addr )
      fail ArgumentError, "Invalid range: #{ addr }" unless addr =~ /\A[A-Z]+\d+:[A-Z]+\d+\z|\A[A-Z]+\d+\z/
      @sheet = sheet
      @data = sheet.data
      @address = addr
      @column = column_id( addr )
      @row = row_id( addr )
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
      "#{ self.class }:0x#{ '%x' % (object_id << 1) }: #{ address }"
    end
  
    include Enumerable
    
    def each
      expand( address ).flatten.each { |addr| yield data[ addr ] }
    end
    
    def each_cell
      expand( address ).flatten.each { |addr| yield Element.new( sheet, addr ) }
    end
    
    def empty?
      all? { |v| v.to_s.empty? }
    end
  
    def map!
      expand( address ).flatten.each { |addr| data[ addr ] = yield data[ addr ] }
    end
  
  end

end