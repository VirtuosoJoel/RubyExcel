module RubyExcel

=begin
Handy stuff to add:
Some way to support += and -=  and << with each class?
Get specific columns from an array (or arg list) of headers.
get the row number from a header (or other address type?) and a lookup value: =MATCH()
get the address of a value: =FIND()
filter the data with a column header and a block. Add a reverse-logic alternative for this?
unique the rows by a header
add upcase and strip options for the data
add tools to handle date conversion
add the ability to summarise a column
add a sumif and a countif
add something to the excel dump which takes a range and puts outer borders on it, plus optional inner borders.
add the ability to loop across a column or row while appending items. Maybe by referencing a section outside the existing range?

add the ability to import (recursively?) a nested hash into something like this:

{ Type1: { SubType1: 1, SubType2: 2, SubType3: 3 }, Type2: { SubType1: 4, SubType2: 5, SubType3: 6 } }

Type1 SubType1 1
      SubType2 2
      SubType3 3
Type2
      SubType1 4
      SubType2 5
      SubType3 6
      
=end

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
    
    def empty?
      data.empty?
    end
    
    def column_by_header( header )
      Column.new( self, data.colrf_by_header( header ) )
    end
    alias ch column_by_header

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
