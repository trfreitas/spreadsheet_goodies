
module SpreadsheetGoodies::Excel

  def self.read_workbook(file_pathname:)
    Workbook.new(file_pathname)
  end

end

# loads all files in excel folder
Dir[File.join(File.dirname(__FILE__), "excel/**/*.rb")].each { |f| require f }
