# A cached copy of an Excel worksheet (a single workbook "tab"), stored as an
# array of arrays. Individual cells can be accessed by column title, e.g.:
#   worksheet[0]['A column title']
module SpreadsheetGoodies::GoogleDrive
  class Workbook
    attr_reader :workbook

    def initialize(spreadsheet_key)
      @spreadsheet_key = spreadsheet_key
    end

    def sheets
      underlying_adapter_worksheet.worksheets.map { |ws| ws.title }
    end

    def worksheet(worksheet_title_or_index = 0, num_header_rows = 1)
      Worksheet.new(@spreadsheet_key, worksheet_title_or_index, num_header_rows)
    end

    private

    # Reads a GoogleDrive::Spreadsheet object from Google Drive, using the
    # google_drive gem. Documentation for this class can be found here:
    # https://www.rubydoc.info/github/gimite/google-drive-ruby/GoogleDrive/Spreadsheet
    def underlying_adapter_worksheet
      @underlying_adapter_worksheet ||= Connector.new.read_workbook(@spreadsheet_key)
    end
  end
end
