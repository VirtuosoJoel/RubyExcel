require_relative '../../rubyexcel'
require 'test/unit'

class TestExcel < Test::Unit::TestCase
  
  def setup
    @wb = RubyExcel.sample_sheet.parent
  end
  
  def teardown
    @wb = nil
    @excel = @exwb.application rescue nil
    @exwb.close(0) rescue nil
    @excel.quit() rescue nil
  end

  def test_to_excel
    
    assert( @exwb = @wb.to_excel )
    
  end
  
end
