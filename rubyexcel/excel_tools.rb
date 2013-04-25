require 'win32ole' #Interface with Excel
require 'win32/registry' #Find Documents / My Documents for default directory

module ExcelConstants; end

module RubyExcel

  def self.borders( range, weight=1, inner=false )
    range.ole_respond_to?( :borders ) or fail ArgumentError, 'First Argument must be WIN32OLE Range'
    [0,1,2,3].include?( weight ) or fail ArgumentError, "Invalid line weight #{ weight }. Must be from 0 to 3"
    defined?( ExcelConstants::XlEdgeLeft ) or WIN32OLE.const_load( range.application, ExcelConstants )
    consts = [ ExcelConstants::XlEdgeLeft, ExcelConstants::XlEdgeTop, ExcelConstants::XlEdgeBottom, ExcelConstants::XlEdgeRight, ExcelConstants::XlInsideVertical, ExcelConstants::XlInsideHorizontal ]
    inner or consts.pop(2)
    weight = [ 0, ExcelConstants::XlThin, ExcelConstants::XlMedium, ExcelConstants::XlThick ][ weight ]
    consts.each { |const| weight.zero? ? range.Borders( const ).linestyle = ExcelConstants::XlNone : range.Borders( const ).weight = weight }
    range
  end

  class Workbook
  
    def dump_to_sheet( data, sheet=nil )
      data.is_a?( Array ) or fail ArgumentError, "Invalid data type: #{ data.class }"
      sheet ||= get_workbook.sheets(1)
      sheet.range( sheet.cells( 1, 1 ), sheet.cells( data.length, data[0].length ) ).value = data
      sheet
    end

    def get_excel( invisible = false )
      excel = WIN32OLE::connect( 'excel.application' ) rescue WIN32OLE::new( 'excel.application' )
      excel.visible = true unless invisible
      excel
    end
    
    def get_workbook( excel=nil )
      excel ||= get_excel
      wb = excel.workbooks.add
      ( ( wb.sheets.count.to_i ) - 1 ).times { |time| wb.sheets(2).delete }
      wb
    end

    def make_sheet_pretty( sheet )
      c = sheet.cells
      c.rowheight = 15
      c.entireColumn.autoFit
      c.horizontalAlignment = -4108
      c.verticalAlignment = -4108
      sheet
    end
    
    def save_excel( filename = 'Output', invisible = false )
      filename = filename.gsub('/','\\')
      unless filename.include?('\\')
        keypath = 'SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Shell Folders'
        documents = Win32::Registry::HKEY_CURRENT_USER.open(keypath)['Personal'] rescue Dir.pwd.gsub('/','\\')
        filename = documents + '\\' + filename 
      end
      wb = to_excel( invisible )
      wb.saveas filename
      wb
    end
    
    def to_excel
      self.sheets.count == self.sheets.map(&:name).uniq.length or fail NoMethodError, 'Duplicate sheet name'
      wb = get_workbook
      wb.parent.displayAlerts = false
      first_time = true
      self.each do |s|
        sht = ( first_time ? wb.sheets(1) : wb.sheets.add( { 'after' => wb.sheets( wb.sheets.count ) } ) ); first_time = false
        sht.name = s.name
        make_sheet_pretty( dump_to_sheet( s.data.all, sht ) )
      end
      wb.sheets(1).select
      wb
    end
    
  end
  
end