require_relative 'address.rb'
require_relative 'data.rb'
require_relative 'element.rb'
#require_relative 'excel_tools.rb'
require_relative 'section.rb'
require_relative 'sheet.rb'

module RubyExcel

  #
  # Example data to use in tests / demos
  #
  
  def self.sample_data
    [
      [ 'Part',  'Ref1', 'Ref2', 'Qty', 'Cost' ],
      [ 'Type1', 'QT1',  '231',  1,     35.15  ], 
      [ 'Type2', 'QT3',  '123',  1,     40     ], 
      [ 'Type3', 'XT1',  '321',  3,     0.1    ], 
      [ 'Type1', 'XY2',  '132',  1,     30.00  ], 
      [ 'Type4', 'XT3',  '312',  2,     3      ], 
      [ 'Type2', 'QY2',  '213',  1,     99.99  ], 
      [ 'Type1', 'QT4',  '123',  2,     104    ]
    ]
    
  end
  
  #
  # Example hash to demonstrate imports
  #
  
  def self.sample_hash
  
    {
      Part1: {
        Type1: {
          SubType1: 1, SubType2: 2, SubType3: 3
        },
        Type2: {
          SubType1: 4, SubType2: 5, SubType3: 6
        }
      },
      Part2: {
        Type1: {
          SubType1: 1, SubType2: 2, SubType3: 3
        },
        Type2: {
          SubType1: 4, SubType2: 5, SubType3: 6
        }
      }
    }
  
  end
  
  #
  # Shortcut to create a Sheet with example data
  #
  
  def self.sample_sheet
    Workbook.new.load RubyExcel.sample_data
  end
 
  #
  # Shortcut to import a WIN32OLE Workbook or Sheet
  #
 
  def self.import( *args )
    Workbook.new.import( *args )
  end
 
end