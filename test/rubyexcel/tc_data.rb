require_relative '../../rubyexcel'
require 'test/unit'

class TestData < Test::Unit::TestCase
  
  def setup
    @s = RubyExcel.sample_sheet
  end
  
  def teardown
    @s = nil
  end
  
  def test_advanced_filter!
    
    assert_equal( 3, @s.data.advanced_filter!( 'Part', :=~, /Type[13]/, 'Qty', :>, 1 ).rows )
    
  end

  def test_colref_by_header
  
    assert_equal( 'B', @s.data.colref_by_header( 'Ref1' ) )
  
  end
  
  def test_compact
  
    @s << [[1,2,3], [4,5,6]]
    assert( @s.data.cols == 5 && @s.data.rows == 10 )
  
    @s.rows( 9 ) { |r| r.map! { nil } }
    @s.data.compact!
    assert_equal( 8, @s.data.rows )
  
  end
  
  def test_delete
  
    @s.data.delete( @s.row(1) )
    assert_equal( 7, @s.data.rows )
    
    @s.data.delete( @s.column(1) )
    assert_equal( 4, @s.data.cols )
    
    @s.data.delete( @s.range('A:A') )
    assert_equal( 3, @s.data.cols )
    
    assert_raise( NoMethodError ) { @s.data.delete( [[]] ) }
  
  end
  
  def test_each
  
    assert_equal( 8, @s.data.each.count )
  
  end
  
  def test_filter!
  
    assert_equal( 3, @s.data.filter!( 'Part', &/Type2/ ).rows )
  
  end
  
  def test_get_columns!

    assert_equal( [ 'Ref2', 'Part' ], @s.data.get_columns!( 'Ref2', 'Part' ).sheet.row(1).to_a )

  end
  
  def test_headers
  
    assert_equal( 1, @s.data.headers.length )
  
    @s.headers = 0
    assert_equal( nil, @s.data.headers )
  
  end
  
  def test_index_by_header

    assert_equal( 1, @s.data.index_by_header( 'Part' ) )
  
    @s.headers = 0
    assert_raise( NoMethodError ) { @s.data.index_by_header( 'Part' ) }

  end
  
  def test_insert
  
    @s.data.insert_columns( 'A', 2 )
    assert_equal( 7, @s.maxcol )
    assert_equal( nil, @s['B2'] )
    
    @s.data.insert_rows( 2, 2 )
    assert_equal( 10, @s.maxrow )
    assert_equal( nil, @s['B4'] )
  
  end
  
  def test_no_headers

    assert_equal( 7, @s.data.no_headers.length )

  end
  
  def test_partition
  
    ar1, ar2 = @s.data.partition( 'Part', &/Type[13]/ )
    assert_equal( 5, ar1.length )
    assert_equal( 'Type1', ar1[1][0] )
    assert_equal( 4, ar2.length )
    
  end
  
  def test_read_write
  
    assert_equal( '123', @s.data.read( 'C3' ) )
    
    @s.data.write( 'C3', '321' )
    assert_equal( '321', @s.data.read( 'C3' ) )
  
  end
  
  def test_reverse
  
    @s.data.reverse_columns!
    assert_equal( 'Cost', @s.A1 )
    
    @s.data.reverse_rows!
    assert_equal( 'QT1', @s.d8 )
    
    @s.headers = 0
    @s.data.reverse_rows!
    assert_equal( 'Ref1', @s.d8 )
  
  end
  
  def test_skip_headers
  
    @s.load( @s.data.skip_headers { |data| data.map { |row| row.map { nil } } } )
    assert_equal( 'Part', @s.a1 )
    assert_equal( nil, @s.a2 )
  
  end
  
  def test_uniq!
  
    @s.data.uniq!( 'Part' )
    assert_equal( 5, @s.maxrow )
  
  end
  
end