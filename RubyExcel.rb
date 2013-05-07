require_relative 'rubyexcel/rubyexcel_components.rb'
require 'cgi'

#
# Ruby's standard Regexp class.
#   Regexp#to_proc is a bit of "syntactic sugar" which allows shorthand Regexp blocks
#
# @example
#   sheet.filter!( 'Part', &/Type[13]/ )
#

class Regexp
  def to_proc
    proc { |s| self =~ s.to_s }
  end
end

#
# Namespace for all RubyExcel Classes and Modules
#

module RubyExcel

  #
  # A Workbook which can hold multiple Sheets
  #

  class Workbook
    include Enumerable

    #
    # Creates a RubyExcel::Workbook instance.
    #
    
    def initialize
      @sheets = []
    end
    
    #
    # Appends an object to the Workbook
    #
    # @param [RubyExcel::Workbook, RubyExcel::Sheet, Array<Array>] other the object to append to the Workbook
    #
    
    def <<( other )
      case other
      when Workbook ; other.each { |sht| sht.workbook = self; @sheets << sht }
      when Sheet    ; @sheets << other; other.workbook = self
      when Array    ; load( other )
      else          ; fail TypeError, "Unsupported Type: #{ other.class }"
      end
      self
    end
    
    #
    # Adds a Sheet to the Workbook.
    #   If no argument is given, names the Sheet 'Sheet' + total number of Sheets
    #
    # @example
    #   sheet = workbook.add
    #   #=> RubyExcel::Sheet:0x2b3a0b8: Sheet1
    #
    # @param [nil, RubyExcel::Sheet, String] ref the identifier or Sheet to add
    # @return [RubyExcel::Sheet] the Sheet which was added
    
    def add( ref=nil )
      case ref
      when nil    ; s = Sheet.new( 'Sheet' + ( @sheets.count + 1 ).to_s, self )
      when Sheet  ; ( s = ref ).workbook = self
      when String ; s = Sheet.new( ref, self )
      else        ; fail TypeError, "Unsupported Type: #{ ref.class }"
      end
      @sheets << s
      s
    end
    alias add_sheet add
    
    #
    # Removes all Sheets from the Workbook
    #
    
    def clear_all
      @sheets = []; self
    end
    alias delete_all clear_all
    
    #
    # Removes Sheet(s) from the Workbook
    #
    # @param [Fixnum, String, Regexp, RubyExcel::Sheet] ref the reference or object to remove
    #
    
    def delete( ref )
      case ref
      when Fixnum ; @sheets.delete_at( ref - 1 )
      when String ; @sheets.reject! { |s| s.name == ref }
      when Regexp ; @sheets.reject! { |s| s.name =~ ref }
      when Sheet  ; @sheets.reject! { |s| s == ref }
      else        ; fail ArgumentError, 'Unrecognised Argument Type: ' + ref.class.to_s
      end
      self
    end
    
    #
    # Return a copy of self
    #
    # @return [RubyExcel::Workbook]
    #
    
    def dup
      wb = Workbook.new
      self.each { |s| wb.add s.dup }
      wb
    end
    
    #
    # Check whether the workbook has Sheets
    #
    # @return [Boolean] if there are any Sheets in the Workbook
    #
    
    def empty?
      @sheets.empty?
    end
    
    # @overload load( input_data, header_rows=1 )
    #   Shortcut to create a Sheet and fill it with data
    #   @param [Array<Array>, Hash<Hash>] input_data the data to fill the Sheet with
    #   @param Fixnum] header_rows the number of Rows to be treated as headers
    #
    
    def load( *args )
      add.load( *args )
    end

    #
    # Select a Sheet or iterate through them
    #
    # @param [Fixnum, String, nil] ref the reference to select a Sheet by
    # @return [RubyExcel::Sheet] if a search term was given
    # @return [Enumerator] if nil or no argument given
    #
    
    def sheets( ref=nil )
      return to_enum (:each) if ref.nil?
      ref.is_a?( Fixnum ) ? @sheets[ ref - 1 ] : @sheets.find { |s| s.name =~ /^#{ ref }$/i }
    end
    
    # {Workbook#sort!}
    
    def sort( &block )
      dup.sort!( &block )
    end
    
    #
    # Sort Sheets according to a block
    #
    
    def sort!( &block )
      @sheets = @sheets.sort( &block )
    end
    
    # {Workbook#sort_by!}
    
    def sort_by( &block )
      dup.sort_by!( &block )
    end
    
    #
    # Sort Sheets by an attribute given in a block
    #
    
    def sort_by!( &block )
      @sheets = @sheets.sort_by( &block )
    end
    
    #
    # Yields each Sheet.
    #
    
    def each
      return to_enum( :each ) unless block_given?
      @sheets.each { |s| yield s }
    end
    
  end # Workbook

end # RubyExcel