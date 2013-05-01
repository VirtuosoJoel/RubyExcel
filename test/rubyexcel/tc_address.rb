require_relative '../../rubyexcel'
require 'test/unit'


class TestAddress < Test::Unit::TestCase
  
  def setup
    @s = RubyExcel.sample_sheet
  end
  
  def teardown
    @s = nil
  end
  
  def test_address_to_col_index
    
    assert( @s.address_to_col_index( 'A1' ) == 1 )
    
  end
  
  def test_address_to_indices
    
    assert_equal( [ 1, 1 ], @s.address_to_indices( 'A1' ) )
    
  end
  
  def test_col_index
  
    assert_equal( 1, @s.col_index( 'A' ) )
  
  end

  def test_col_letter

    assert_equal( 'A', @s.col_letter( 1 ) )
  
  end

  def test_column_id

    assert_equal( 'A', @s.column_id( 'A1' ) )
  
  end
  
  def test_expand

    assert_equal( [['A1','B1'],['A2', 'B2']], @s.expand( 'A1:B2' ) )
  
  end
  
  def test_indices_to_address
  
    assert_equal( 'A1', @s.indices_to_address( 1, 1 ) )
  
  end
  
  def test_multi_array?
  
    assert( @s.multi_array? RubyExcel.sample_data )
  
  end
  
  def test_offset
    
    assert_equal( 'B2', @s.offset( 'A1', 1, 1 ) )
    
  end
  
  def test_to_range_address

    assert_equal( 'A1:B2', @s.to_range_address( @s.cell(1,1), @s.cell(2,2) ) )

  end
  
  def test_row_id
  
    assert_equal( 2, @s.row_id( 'A2' ) )
  
  end
  
end