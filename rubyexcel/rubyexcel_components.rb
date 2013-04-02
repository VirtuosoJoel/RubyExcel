require_relative 'address.rb'
require_relative 'data.rb'
require_relative 'element.rb'
require_relative 'section.rb'

module RubyExcel

  def self.sample_data
    a=[];8.times{|t|b=[];c='A';5.times{b<<"#{c}#{t+1}";c.next!};a<<b};a
  end
  
  def self.sample_sheet
    Workbook.new.load RubyExcel.sample_data
  end
 
end