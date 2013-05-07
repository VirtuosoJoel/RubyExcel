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