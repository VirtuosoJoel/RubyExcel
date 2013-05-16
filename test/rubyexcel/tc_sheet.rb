require_relative '../../rubyexcel'
require 'test/unit'

class TestSheet < Test::Unit::TestCase
  
  def setup
    @s = RubyExcel.sample_sheet
  end
  
  def teardown
    @s = nil
  end

  def test_basics
    
    assert( @s['A1'] == @s.A1 )
    
    @s << RubyExcel.sample_hash
    assert( @s.maxrow == 20 && @s.maxcol == 5 )
    
    @s << RubyExcel.sample_data
    assert( @s.maxrow == 28 && @s.maxcol == 5 )
    
    @s << @s
    assert( @s.maxrow == 55 && @s.maxcol == 5 )
    
  end
  
  def test_advanced_filter
    
    @s.advanced_filter!( 'Part', :=~, /Type[13]/, 'Qty', :>, 1 )
    assert( @s.maxrow == 3 )
  
    setup
    @s.advanced_filter!( 'Part', :==, 'Type1', 'Ref1', :include?, 'X' )
    assert( @s.maxrow == 2 )
    
  end
  
  def test_advanced_filter
    
    assert_equal( 42, @s.averageif( 'Part', 'Cost', &/Type[13]/ ).to_i )
  
    assert_equal( 35, @s.averageif( 'Qty', 'Cost' ) { |i| i.between?( 2,3 ) }.to_i )
  
    assert_equal( 1, @s.averageif( 'Part', 'Qty' ) {true}.to_i )
  
  end

  def test_cell
  
    assert( @s.cell(1,1).value == 'Part' && @s.cell(1,1).address == 'A1' )
    
  end
  
  def test_column
  
    assert( @s.column('A')[1] == 'Part' )
  
  end
  
  def test_column_by_header
  
    assert( @s.ch( 'Part' )[1] == @s.ch( @s.column(1) )[1] )
    
  end
  
  def test_columns
  
    assert( @s.columns.class == Enumerator )
    
    assert( @s.columns( 'B' ).count == 4 )
    
    assert( @s.columns( 'B', 'D' ).to_a[0][1] == 'Ref1' )
  
  end
  
  def test_filter
  
    assert( @s.filter( 'Part', &/Type[13]/ ).maxrow == 5 )
  
  end
  
  def test_get_columns
  
    assert_equal( @s.get_columns( 'Ref2', 'Part' ).row(1).to_a, [ 'Ref2', 'Part' ] )
    
  end
  
  def test_insert_columns
    
    assert( @s.insert_columns( 2, 2 ).columns.count == 7 )
    
  end
  
  def test_insert_rows
  
    assert( @s.insert_rows( 2, 2 ).rows.count == 10 )
  
  end
  
  def test_match
  
    assert( @s.match( 'Part', &/Type2/ ) == 3 )
  
  end
  
  def test_method_missing
  
    assert( @s.a1 == 'Part' )
  
    assert_raise( NoMethodError ) { @s.abcd123 }
  
  end
  
  def test_respond_to?
  
    assert( @s.respond_to?(:A1) )
  
  end
  
  def test_partition
  
    s1, s2 = @s.partition( 'Qty' ) { |v| v > 1 }
    assert_equal( s1.maxrow + 1, s2.maxrow )

  end
  
  def test_range
  
    assert( @s.range( 'A1:A1' ).value == @s.range( @s.cell(1,1), @s.cell( 1,1 ) ).value )
    
    assert( @s.range( 'A1' ).value == @s.range( @s.cell(1,1) ).value )
 
  end
  
  def test_reverse
  
    assert( @s.reverse_columns!['A1'] == 'Cost' )
    
    assert( @s.reverse_rows!['A2'] == 104 )
    
  end
  
  def test_row
  
    assert( @s.row(1)['A'] == 'Part' )
  
  end
  
  def test_rows
  
    assert( @s.rows.class == Enumerator )
    
    assert( @s.rows( 2 ).count == 7 )
    
    assert( @s.rows( 2, 4 ).to_a[0]['A'] == 'Type1' )
  
  end
  
  def test_sort_by
  
    assert_equal( @s.A3, @s.sort_by( 'Part' ).A5 )
  
  end
  
  def test_split
  
    assert_equal( 4, @s.split('Part').sheets.count )
    
  end
  
  def test_sumif
  
    assert_equal( 169.15, @s.sumif( 'Part', 'Cost', &/Type1/ ) )
  
  end
  
  def test_uniq
  
    assert_equal( 5, @s.uniq( 'Part' ).maxrow )

  end

  def test_vlookup
  
    assert_equal( '231', @s.vlookup( 'Part', 'Ref2', &/Type1/ ) )
  
  end
  
end
