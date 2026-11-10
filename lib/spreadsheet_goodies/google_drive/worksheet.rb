require_relative '../abstract_base_worksheet'
require_relative '../row'

# A cached copy of a Google Spreadsheets worksheet (i.e., a single workbook "tab"),
# Individual cells of a row can be accessed by column title.
# Example:
#   worksheet[0]['A column title']
module SpreadsheetGoodies::GoogleDrive
  class Worksheet < SpreadsheetGoodies::AbstractBaseWorksheet
    attr_reader :title

    # @param spreadsheet_key [String] Spreadsheet identifier, which can be retrieved
    #   from the spreasheet's url
    # @param worksheet_title_or_index [String, Integer, nil] Sheet title or
    #   zero-based index; if unspecified, the first sheet will be used.
    # @param num_header_rows [Integer] Number of rows at the top of the sheet that
    #   contain headers or stuff other than data. Optional; if unspecified, assumes
    #   that a single top row contains the header and all rows below are data.
    # @param connector [Connector] Optional; built from
    #   SpreadsheetGoodies.configuration if omitted.
    def initialize(spreadsheet_key, worksheet_title_or_index = nil, num_header_rows = 1, connector: nil)
      @spreadsheet_key = spreadsheet_key
      @num_header_rows = num_header_rows
      @connector = connector || Connector.new
      @title = @connector.resolve_sheet_title(spreadsheet_key, worksheet_title_or_index)
      @pending_writes = {}

      read_from_google_drive
    end

    # Writes to a given cell identified by row and column indexes (they start at 1).
    # The change is only sent to Google when #commit_writes! is called.
    # @override
    def write_to_cell(row_index, col_index, value)
      @pending_writes[[row_index, col_index]] = value
    end

    # Commit writes so they are propagated to the real spreadsheet
    # @override
    def commit_writes!
      @connector.write_values(@spreadsheet_key, @title, @pending_writes)
      @pending_writes = {}
    end

  private

    # Reads sheet contents and caches it into instance variables to so that it
    # can be accessed later without accessing Google Drive.
    def read_from_google_drive
      rows = @connector.read_values(@spreadsheet_key, @title)

      if SpreadsheetGoodies.configuration.strip_values_on_read
        rows = rows.map do |row|
          row.map { |element| element&.strip }
        end
      end

      @header_row = rows[@num_header_rows-1]
      @rows = rows.collect.with_index do |row, index|
        SpreadsheetGoodies::Row.new(@header_row, index+1, self, *row)
      end

      self
    end
  end
end
