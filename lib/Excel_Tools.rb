
module Excel_Tools

  require 'win32ole'

  def get_excel
    excel = WIN32OLE::connect( 'excel.application' ) rescue WIN32OLE::new( 'excel.application' )
    excel.visible = true
    excel
  end
  
  def get_workbook( excel=nil )
    excel ||= get_excel
    wb = excel.workbooks.add
    ( ( wb.sheets.count.to_i ) - 1 ).times { |time| wb.sheets(2).delete }
    wb
  end
  
  def dump_to_sheet( data, sheet=nil )
    fail ArgumentError, "Invalid data type: #{ data.class }" unless data.is_a?( Array ) || data.is_a?( RubyExcel )
    data = data.to_a if data.is_a? RubyExcel
    sheet ||= get_workbook.sheets(1)
    sheet.range( sheet.cells( 1, 1 ), sheet.cells( data.length, data[0].length ) ).value = data
    sheet
  end
  
  def make_sheet_pretty( sheet )
    c = sheet.cells
    c.EntireColumn.AutoFit
    c.HorizontalAlignment = -4108
    c.VerticalAlignment = -4108
    sheet
  end
  
  def to_excel
    wb = get_workbook
    wb.parent.DisplayAlerts = false
    first_time = true
    self.each do |s|
      sht = ( first_time ? wb.sheets(1) : wb.sheets.add( { 'after' => wb.sheets( wb.sheets.count ) } ) ); first_time = false
      sht.name = s.name
      make_sheet_pretty( dump_to_sheet( s.data.all, sht ) )
    end
    wb.sheets(1).select
    wb
  end

  def save_excel( filename = 'Output.xlsx' )
    filename = Dir.pwd.gsub('/','\\') + '\\' + filename unless filename.include?('\\')
    to_excel.saveas filename
  end
  
end