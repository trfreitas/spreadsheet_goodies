
module SpreadsheetGoodies::GoogleDrive

  # Accesses GoogleDrive and returns a SpreadsheetGoodies::GoogleDrive::Worksheet
  def self.read_workbook(spreadsheet_key:)
    Workbook.new(spreadsheet_key)
  end

end

# loads all files in google_drive folder
Dir[File.join(File.dirname(__FILE__), "google_drive/**/*.rb")].each { |f| require f }
