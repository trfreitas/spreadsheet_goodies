# A cached copy of an Excel worksheet (a single workbook "tab"), stored as an
# array of arrays. Individual cells can be accessed by column title, e.g.:
#   worksheet[0]['A column title']
module SpreadsheetGoodies::Excel
  class Workbook
    attr_reader :workbook, :sheets, :worksheet

    def initialize(workbook_file_pathname)
      @workbook = case workbook_file_pathname.to_s
        when /\.xls[^x]/ then Roo::Excel.new(workbook_file_pathname, file_warning: :ignore)
        when /\.xlsx/ then Roo::Excelx.new(workbook_file_pathname, file_warning: :ignore)
      end
    end

    def sheets
      @workbook.sheets
    end

    def worksheet(worksheet_title_or_index = 0, num_header_rows = 1)
      Worksheet.new(@workbook, worksheet_title_or_index, num_header_rows)
    end
  end
end
