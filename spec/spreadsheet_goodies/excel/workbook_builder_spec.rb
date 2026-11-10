require 'spec_helper'
require 'tmpdir'
require 'zip'
require 'nokogiri'

describe SpreadsheetGoodies::Excel::WorkbookBuilder do
  around do |example|
    Dir.mktmpdir do |dir|
      @path = File.join(dir, 'out.xlsx')
      example.run
    end
  end

  let(:builder) { described_class.new(@path) }

  before do
    builder.add_worksheet('Employees')
    builder.add_header_row(['Employee', 'Salary'])
    builder.add_data_row(['John', 10000.0])
    builder.add_data_row(['Mary', 12000.5])
    builder.setup_auto_filter
    builder.write_to_file
  end

  def sheet_xml
    xml = Zip::File.open(@path) do |zip|
      zip.read('xl/worksheets/sheet1.xml')
    end
    Nokogiri::XML(xml).remove_namespaces!
  end

  it 'writes the file' do
    expect(File.exist?(@path)).to be true
  end

  it 'exposes the underlying Axlsx workbook' do
    expect(builder.workbook).to be_a Axlsx::Workbook
  end

  it 'writes the data that can be read back' do
    workbook = SpreadsheetGoodies::Excel.read_workbook(file_pathname: @path)
    expect(workbook.sheets).to eq ['Employees']

    worksheet = workbook.worksheet(0)
    expect(worksheet.header_row).to eq ['Employee', 'Salary']
    expect(worksheet.data_rows.map { |r| r['Employee'] }).to eq ['John', 'Mary']
    expect(worksheet.data_rows.map { |r| r['Salary'] }).to eq [10000.0, 12000.5]
  end

  it 'adds an auto filter covering the data' do
    expect(sheet_xml.at_xpath('//autoFilter')['ref']).to eq 'A1:B3'
  end

  it 'freezes the top row' do
    pane = sheet_xml.at_xpath('//sheetView/pane')
    expect(pane['topLeftCell']).to eq 'A2'
    expect(pane['ySplit']).to eq '1'
    expect(pane['state']).to eq 'frozenSplit'
    expect(pane['activePane']).to eq 'bottomRight'
  end
end
