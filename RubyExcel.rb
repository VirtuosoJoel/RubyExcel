require_relative 'rubyexcel/rubyexcel_components.rb'
require_relative 'rubyexcel/excel_tools.rb'

class Regexp
  def to_proc
    proc { |s| self =~ s.to_s }
  end
end

module RubyExcel

  class Workbook

    def initialize
      @sheets = []
    end
    
    def <<( other )
      case other
      when Workbook ; other.inject( @sheets, :<< )
      when Sheet ; @sheets << other
      end
    end
    
    def add( ref=nil )
      case ref
      when nil ; s = Sheet.new( 'Sheet' + ( @sheets.count + 1 ).to_s, self )
      when Sheet ; ( s = ref ).workbook = self
      when String ; s = Sheet.new( ref, self )
      else ; fail TypeError, "Unsupported Type: #{ ref.class }"
      end
      @sheets << s; s
    end
    alias add_sheet add
    
    def clear_all
      @sheets = []; self
    end
    
    def delete( ref )
      case ref
      when Fixnum ; @sheets.delete_at( ref - 1 )
      when String ; @sheets.reject! { |s| s.name == ref }
      when Regexp ; @sheets.reject! { |s| s.name =~ ref }
      when Sheet ; @sheets.reject! { |s| s == ref }
      else ; fail ArgumentError, "Unrecognised Argument Type: #{ ref.class }"
      end ; self
    end
    
    def dup
      wb = Workbook.new
      self.each { |s| wb.add s.dup }
      wb
    end
    
    def empty?
      @sheets.empty?
    end
    
    def load( *args )
      add.load( *args )
    end
    
    def sheets( ref=nil )
      return to_enum (:each) if ref.nil?
      ref.is_a?( Fixnum ) ? @sheets[ ref - 1 ] : @sheets.find { |s| s.name =~ /^#{ ref }$/i }
    end
    
    def sort!
      @sheets = @sheets.sort(&block)
    end
    
    def sort_by!( &block )
      @sheets = @sheets.sort_by(&block)
    end
    
    include Enumerable
    
    def each
      return to_enum(:each) unless block_given?
      @sheets.each { |s| yield s }
    end
    
  end

  class Sheet

    attr_reader :data
    attr_accessor :name, :header_rows, :workbook
    alias parent workbook; alias parent= workbook=
    alias headers header_rows; alias headers= header_rows=
    
    include Address
    
    def initialize( name, workbook )
      @workbook = workbook
      @name = name
      @header_rows = nil
      @data = Data.new( self, [[]] )
    end

    def[]( addr )
      range( addr ).value
    end

    def []=( addr, val )
      range( addr ).value = val
    end

 
    def +( other )
      dup << other
    end
    
    def -( other )
      case other
      when Array ; Workbook.new.load( data.all - other )
      when Sheet ; Workbook.new.load( data.all - other.data.no_headers )
      else ; fail ArgumentError, "Unsupported class: #{ other.class }"
      end
    end
    
    def <<( other )
      case other
      when Array ; load( data.all + other, header_rows )
      when Sheet ; load( data.all + other.data.no_headers, header_rows )
      else ; fail ArgumentError, "Unsupported class: #{ other.class }"
      end
    end
    
    def cell( row, col )
      Element.new( self, indices_to_address( row, col ) )
    end
    alias cells cell
    
    def column( index )
      Column.new( self, col_letter( index ) )
    end
    
    def column_by_header( header )
      Column.new( self, data.colref_by_header( header ) )
    end
    alias ch column_by_header
    
    def columns( start_column = 'A', end_column = data.cols )
      return to_enum(:columns, start_column, end_column) unless block_given?
      ( col_letter( start_column )..col_letter( end_column ) ).each { |idx| yield column( idx ) }; self
    end
    
    def compact!
      data.compact!; self
    end
    
    def delete
      workbook.delete self
    end
    
    def delete_rows_if
      rows.reverse_each { |r| r.delete if yield r }; self
    end
    
    def delete_columns_if
      columns.reverse_each { |c| c.delete if yield c }; self
    end
    
    def dup
      s = Sheet.new( name, workbook )
      d = data
      unless d.nil?
        d = d.dup
        s.load( d.all, header_rows )
        d.sheet = s
      end
      s
    end
    
    def empty?
      data.empty?
    end
    
    def filter( ref, &block )
      dup.filter!( ref, &block )
    end

    def filter!( ref, &block )
      data.filter!( ref, &block ); self
    end
    
    def get_columns( *headers )
      dup.data.get_columns!( *headers )
    end
    alias gc get_columns
    
    def get_columns!( *headers )
      data.get_columns!( *headers ); self
    end
    alias gc! get_columns!
    
    def insert_columns( *args )
      data.insert_columns( *args ); self
    end
    
    def insert_rows( *args )
      data.insert_rows( *args ); self
    end
    
    def inspect
      "#{ self.class }:0x#{ '%x' % (object_id << 1) }: #{ name }"
    end
    
    def load( input_data, header_rows=1 )
      @header_rows = header_rows
      @data = Data.new( self, input_data ); self
    end
    
    def match( header, &block )
      row_id( column_by_header( header ).find( &block ) )
    end
    
    def maxrow
      data.rows
    end
    
    def maxcol
      data.cols
    end
    alias maxcolumn maxcol
    
    def range( first_cell, last_cell=nil )
      Element.new( self, to_range_address( first_cell, last_cell ) )
    end
    
    def reverse_columns!
      data.reverse_columns!
    end
    
    def reverse_rows!
      data.reverse_rows!
    end

    def row( index )
      Row.new( self, index )
    end
    
    def rows( start_row = 1, end_row = data.rows )
      return to_enum(:rows, start_row, end_row) unless block_given?
      ( start_row..end_row ).each { |idx| yield row( idx ) }; self
    end
    
    def sort!( &block )
      data.sort!( &block ); self
    end
    
    def sort_by!( &block )
      data.sort_by!( &block ); self
    end
    
    def sumif( find_header, sum_header )
      find_col, sum_col  = ch( find_header ), ch( sum_header )
      find_col.each_cell.inject(0) { |sum,ce| yield( ce.value ) && ce.row > header_rows ? sum + sum_col[ ce.row ] : sum }
    end
    
    def to_a
      data.all
    end
    
    def to_excel
      workbook.dup.clear_all.add( self.dup ).workbook.to_excel
    end
    
    def to_s
      data.nil? ? '' : data.map { |ar| ar.join "\t" }.join( $/ )
    end
    
    def uniq!( header )
      data.uniq!( header ); self
    end
    alias unique! uniq!
    
    def vlookup( find_header, return_header, &block )
      find_col, return_col  = ch( find_header ), ch( return_header )
      return_col[ row_id( find_col.find( &block ) ) ] rescue nil
    end
    
  end
  
end