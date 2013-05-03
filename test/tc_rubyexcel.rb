require_relative '../rubyexcel'
require 'test/unit'

class TestRegexp < Test::Unit::TestCase
 
  def test_to_proc
    assert_equal(0, /a/.to_proc.call('a') )
  end
 
end
 
class TestWorkbook < Test::Unit::TestCase
  
  def setup
    @wb = RubyExcel::Workbook.new
    3.times { @wb.add.load RubyExcel.sample_data }
  end
  
  def teardown
    @wb = nil
  end

  def test_shovel
  
    @wb << @wb.dup
    assert_equal( 6, @wb.sheets.count )
    
    @wb << @wb.sheets(1)
    assert_equal( 7, @wb.sheets.count )
    
    @wb << @wb.sheets(1).data.all
    assert_equal( 8, @wb.sheets.count )
    
  end
  
  def test_add
    
    @wb.add 'Sheet4'
    assert_equal( 4, @wb.sheets.count )
    
    @wb.add
    assert_equal( 5, @wb.sheets.count )
    
    @wb.add @wb.sheets(1)
    assert_equal( 6, @wb.sheets.count )
    
    assert_raise( TypeError ) { @wb.add 1 }
    
  end
  
  def test_clear_all
    
    assert_equal( 0, @wb.clear_all.sheets.count )
    
  end
  
  def test_delete
    
    assert_equal( 2, @wb.delete(1).sheets.count )
    
    assert_equal( 1, @wb.delete( 'Sheet2' ).sheets.count )
    
    assert_equal( 0, @wb.delete( /Sheet/ ).sheets.count )
    
    assert_equal( 0, @wb.delete( @wb.add ).sheets.count )
    
  end
  
  def test_dup
  
    dup_wb = @wb.dup
    
    assert_equal( @wb.sheets(1).to_a, dup_wb.sheets(1).to_a )
    
    assert_not_equal( @wb.object_id, dup_wb.object_id )
  
  end
  
  def test_empty?
  
    assert( !@wb.empty? )
    
    assert( @wb.clear_all.empty? )
  
  end
  
  def test_load
  
    assert( @wb.load( [[]] ).class == RubyExcel::Sheet )
    
    assert( @wb.load( RubyExcel.sample_data )['A1'] == 'Part' )
    
  end
  
  def test_sheets
    
    assert( @wb.sheets.class == Enumerator )
    
    assert( @wb.sheets(2) == @wb.sheets('Sheet2') )
    
  end

end

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
  
  def test_split
  
    assert_equal( 4, @s.split('Part').sheets.count )
    
  end
  
  def test_sumif
  
    assert( @s.sumif( 'Part', 'Cost', &/Type1/ ) == 169.15 )
  
  end
  
  def test_uniq
  
    assert( @s.uniq( 'Part' ).maxrow ==  5 )

  end

  def test_
  
    assert( @s.vlookup( 'Part', 'Ref2', &/Type1/ ) == '231' )
  
  end
  
  
  
end
