require_relative '../../lib/rubyexcel'
require 'test/unit'

class TestRowColumn < Test::Unit::TestCase
  
  def setup
    @s = RubyExcel.sample_sheet
    @r = @s.row(2)
    @c = @s.column(2)
  end
  
  def teardown
    @s = nil
    @r = nil
    @c = nil
  end
  
  def test_shovel
  
    @r << 1
    assert_equal( 5, @r.length )
    
    @r = @s.row(1)
    @r << 1
    assert_equal( 6, @r.length )
    
  end
  
  def test_cell
  
    assert_equal( @r.cell(2).address, @c.cell(2).address )
  
  end
  
  def test_cell_by_header
  
    assert_equal( @s.A2, @r.cell_h( 'Part' ).value )
  
  end
  
  def test_find
    
    assert_equal( 'B2', @r.find( &/QT1/ ) )
  
  end
  
  def test_summarise
  
    h = { 'Type1' => 3, 'Type2' => 2, 'Type3' => 1, 'Type4' => 1 }
    assert_equal( h, @s.column(1).summarise )
  
  end
  
  def test_getref
  
    assert_equal( 'A', @r.getref( 'Part' ) )
  
  end
  
  def test_value_by_header
  
    assert_equal( 'Type1', @r.val( 'Part' ) )
  
  end
  
  def test_set_value_by_header
  
    @r.set_val( 'Part', 'Moose' )
    assert_equal( 'Moose', @r[1] )
  
  end
  
end