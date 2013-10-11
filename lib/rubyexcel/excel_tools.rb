require 'win32ole' #Interface with Excel
require 'win32/registry' #Find Documents / My Documents for default directory

# Holder for WIN32OLE Excel Constants
module ExcelConstants; end

module RubyExcel

  #
  # Add borders to an Excel Range
  #
  # @param [WIN32OLE::Range] range the Excel Range to add borders to
  # @param [Fixnum] weight the weight of the borders
  # @param [Boolean] inner add inner borders
  # @raise [ArgumentError] 'First Argument must be WIN32OLE Range'
  # @return [WIN32OLE::Range] the range initially given
  #

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
  
  
    #
    # Add a single quote before any equals sign in the data.
    #   Disables any Strings which would have been interpreted as formulas by Excel
    #
    
    def disable_formulas!
      sheets { |s| s.rows { |r| r.each_cell { |ce|
        if ce.value.is_a?( String ) && ce.value[0] == '='
          ce.value = ce.value.sub( /\A=/,"'=" )
        end
      } } }; self
    end
  
    #
    # Find the Windows "Documents" or "My Documents" path, or return the present working directory if it can't be found.
    #
    # @return [String]
    #
  
    def documents_path
      Win32::Registry::HKEY_CURRENT_USER.open( 'SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Shell Folders' )['Personal'] rescue Dir.pwd.gsub('/','\\')
    end
  
    #
    # Drop a multidimensional Array into an Excel Sheet
    #
    # @param [Array<Array>] data the data to place in the Sheet
    # @param [WIN32OLE::Worksheet, nil] sheet optional WIN32OLE Worksheet to use
    # @return [WIN32OLE::Worksheet] the Worksheet containing the data
    #
  
    def dump_to_sheet( data, sheet=nil )
      data.is_a?( Array ) or fail ArgumentError, "Invalid data type: #{ data.class }"
      sheet ||= get_workbook.sheets(1)
      sheet.range( sheet.cells( 1, 1 ), sheet.cells( data.length, data.max_by(&:length).length ) ).value = data
      sheet
    end
    
    #
    # Open or connect to an Excel instance
    #
    # @param [Boolean] invisible leave Excel invisible if creating a new instance
    # @return [WIN32OLE::Excel] the first available Excel application
    #

    def get_excel( invisible = false )
      excel = WIN32OLE::connect( 'excel.application' ) rescue WIN32OLE::new( 'excel.application' )
      excel.visible = true unless invisible
      excel
    end
    
    #
    # Create a new Excel Workbook
    #
    # @param [WIN32OLE::Excel, nil] excel an Excel object to use
    # @param [Boolean] invisible leave Excel invisible if creating a new instance
    # @return [WIN32OLE::Workbook] the new Excel Workbook
    #
    
    def get_workbook( excel=nil, invisible = false )
      excel ||= get_excel( invisible )
      wb = excel.workbooks.add
      ( ( wb.sheets.count.to_i ) - 1 ).times { |time| wb.sheets(2).delete }
      wb
    end

    #
    # Import a WIN32OLE Object as a Workbook or Sheet
    #
    # @param [WIN32OLE::Workbook, WIN32OLE::Sheet, String] other The WIN32OLE Object, either Sheet or Workbook, to import, or a path to the file.
    # @param [String] sheetname the name of a specific Sheet to import.
    # @param [Boolean] keep_formulas Retain Excel formulas rather than importing their current values
    # @return [self] self with the data and name(s) imported.
    #
    
    def import( other, sheetname=nil, keep_formulas=false )
      operation = ( keep_formulas ? :formula : :value )
    
      if other.is_a?( String )
      
        # Filename
        File.exists?( other ) || fail( ArgumentError, "Unable to find file: #{ other }" )
        
        #Open the file with Excel
        excel = WIN32OLE.new( 'excel.application' )
        excel.displayalerts = false
        
        begin
          wb = excel.workbooks.open({'filename'=> other, 'readOnly' => true})
        rescue WIN32OLERuntimeError
          excel.quit
          raise
        end
        
        # Only one sheet, or the entire Workbook?
        if sheetname
          add( sheetname ).load( wb.sheets( sheetname ).usedrange.send( operation ) )
        else
          self.name = File.basename( other, '.*' )
          wb.sheets.each { |sh| add( sh.name ).load( sh.usedrange.send( operation ) ) }
        end
        
        # Cleanup
        wb.close
        excel.quit
        
      elsif !other.respond_to?( :ole_respond_to? )
      
        fail ArgumentError, "Invalid input: #{other.class}"
        
      elsif other.ole_respond_to?( :sheets )
      
        # Workbook
        
        # Only one sheet, or the entire Workbook?
        if sheetname
          add( sheetname ).load( other.sheets( sheetname ).usedrange.send( operation ) )
        else
          self.name = File.basename( other.name, '.*' )
          other.sheets.each { |sh| add( sh.name ).load( sh.usedrange.send( operation ) ) }
        end
        
      elsif other.ole_respond_to?( :usedrange )
      
        # Sheet
        add( other.name ).load( other.usedrange.send( operation ) )
        
      else
      
        fail ArgumentError, "Object not recognised as a WIN32OLE Workbook or Sheet.\n#{other.inspect}"
        
      end
      
      self
    end
    
    #
    # Take an Excel Sheet and standardise some of the formatting
    #
    # @param [WIN32OLE::Worksheet] sheet the Sheet to add formatting to
    # @return [WIN32OLE::Worksheet] the sheet with formatting added
    #
    
    def make_sheet_pretty( sheet )
      c = sheet.cells
      c.rowheight = 15
      c.entireColumn.autoFit
      c.horizontalAlignment = -4108
      c.verticalAlignment = -4108
      sheet.UsedRange.Columns.each { |col| col.ColumnWidth = 30 if col.ColumnWidth > 50 }
      RubyExcel.borders( sheet.usedrange, 1, true )
      sheet
    end
    
    #
    # Save the RubyExcel::Workbook as an Excel Workbook
    #
    # @param [String] filename the filename to save as
    # @param [Boolean] invisible leave Excel invisible if creating a new instance
    # @return [WIN32OLE::Workbook] the Workbook, saved as filename.
    #
    
    def save_excel( filename = nil, invisible = false )
      filename ||= name
      filename = filename.gsub('/','\\')
      unless filename.include?('\\')
        filename = documents_path + '\\' + filename 
      end
      wb = to_excel( invisible )
      wb.saveas filename
      wb
    end
    
    #
    # Output the RubyExcel::Workbook to Excel
    #
    # @param [Boolean] invisible leave Excel invisible if creating a new instance
    # @return [WIN32OLE::Workbook] the Workbook in Excel
    #
    
    def to_excel( invisible = false )
      self.sheets.count == sheets.map(&:name).uniq.length or fail NoMethodError, 'Duplicate sheet name'
      wb = get_workbook( nil, true )
      wb.parent.displayAlerts = false
      first_time = true
      each do |s|
        sht = ( first_time ? wb.sheets(1) : wb.sheets.add( { 'after' => wb.sheets( wb.sheets.count ) } ) ); first_time = false
        sht.name = s.name
        make_sheet_pretty( dump_to_sheet( s.to_a, sht ) )
      end
      wb.sheets(1).select rescue nil
      wb.application.visible = true unless invisible
      wb
    end
    
    # {Workbook#to_safe_format!}
    
    def to_safe_format
      dup.to_safe_format!
    end
    
    #
    # Standardise the data for safe export to Excel.
    #   Set each cell contents to a string and remove leading equals signs.
    #
    
    def to_safe_format!
      sheets { |s| s.rows { |r| r.map! { |v|
        if v.is_a?( String )
          v[0] == '=' ? v.sub( /\A=/,"'=" ) : v
        else
          v.to_s
        end
      } } }; self
    end
    
  end # Workbook
  
end # RubyExcel