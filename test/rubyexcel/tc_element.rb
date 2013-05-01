require_relative '../../rubyexcel'
require 'test/unit'

class TestElement < Test::Unit::TestCase
  
  def setup
    @s = RubyExcel.sample_sheet
  end
  
  def teardown
    @s = nil
  end
  
  def test_initialize
    
    r = @s.range( 'A1:B2' )
    assert_equal( @s, r.sheet )
    assert_equal( 'A1:B2', r.address )
    assert_equal( 'A', r.column )
    assert_equal( 1, r.row )
    
  end

  def test_value
    
    r = @s.range( 'A1' )
    assert_equal( 'Part', r.value )
    
    r = @s.range( 'A1:B2' )
    assert_equal( [['Part', 'Ref1'], ['Type1', 'QT1']], r.value )
    
    assert_raise( ArgumentError ) { r.value = [[1, 2]] }
  
  end
  
  def test_each
  
    assert_equal( 6, @s.range( 'A1:C2' ).each.count )
  
  end
  
end