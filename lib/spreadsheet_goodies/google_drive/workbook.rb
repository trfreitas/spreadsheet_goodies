# A Google Spreadsheet (a document with one or more worksheets, or "tabs").
# Worksheets are read on demand with #worksheet.
module SpreadsheetGoodies::GoogleDrive
  class Workbook
    # @param connector [Connector] Optional; shared by all worksheets read from
    #   this workbook. Built from SpreadsheetGoodies.configuration if omitted.
    def initialize(spreadsheet_key, connector: nil)
      @spreadsheet_key = spreadsheet_key
      @connector = connector
    end

    # The spreadsheet metadata, as a Google::Apis::SheetsV4::Spreadsheet
    def workbook
      connector.read_workbook(@spreadsheet_key)
    end

    def sheets
      connector.sheet_titles(@spreadsheet_key)
    end

    # @param worksheet_title_or_index [String, Integer] Sheet title or
    #   zero-based index
    def worksheet(worksheet_title_or_index = 0, num_header_rows = 1)
      Worksheet.new(@spreadsheet_key, worksheet_title_or_index, num_header_rows, connector: connector)
    end

    private

    def connector
      @connector ||= Connector.new
    end
  end
end
