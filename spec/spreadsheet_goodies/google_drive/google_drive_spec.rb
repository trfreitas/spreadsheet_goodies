require 'spec_helper'
require 'json'
require 'net/http'
require 'openssl'

describe 'SpreadsheetGoodies::GoogleDrive' do
  let(:key) { 'spreadsheet-key' }
  let(:service) { instance_double(Google::Apis::SheetsV4::SheetsService) }
  let(:connector) { SpreadsheetGoodies::GoogleDrive::Connector.new(service: service) }

  # Values as returned by the Sheets API: interior empty cells are '' and
  # trailing empty cells are omitted.
  let(:api_values) do
    [
      ['Employee', 'Salary', 'Notes'],
      ['John', '10,000.00'],
      ['', '12,000.50', ' remote '],
    ]
  end

  before do
    spreadsheet = Google::Apis::SheetsV4::Spreadsheet.new(
      sheets: [
        Google::Apis::SheetsV4::Sheet.new(properties: Google::Apis::SheetsV4::SheetProperties.new(title: 'Employees')),
        Google::Apis::SheetsV4::Sheet.new(properties: Google::Apis::SheetsV4::SheetProperties.new(title: "Bob's tab")),
      ]
    )
    allow(service).to receive(:get_spreadsheet)
      .with(key, fields: 'spreadsheetId,properties.title,sheets.properties')
      .and_return(spreadsheet)
    allow(service).to receive(:get_spreadsheet_values)
      .with(key, "'Employees'", value_render_option: 'FORMATTED_VALUE')
      .and_return(Google::Apis::SheetsV4::ValueRange.new(values: api_values))
    allow(service).to receive(:get_spreadsheet_values)
      .with(key, "'Bob''s tab'", value_render_option: 'FORMATTED_VALUE')
      .and_return(Google::Apis::SheetsV4::ValueRange.new(values: nil))
  end

  after { SpreadsheetGoodies.reset }

  describe SpreadsheetGoodies::GoogleDrive::Workbook do
    let(:workbook) { described_class.new(key, connector: connector) }

    it 'lists the sheet titles' do
      expect(workbook.sheets).to eq ['Employees', "Bob's tab"]
    end

    it 'reads a worksheet by index' do
      expect(workbook.worksheet(0).title).to eq 'Employees'
      expect(workbook.worksheet(1).title).to eq "Bob's tab"
    end

    it 'reads a worksheet by title' do
      expect(workbook.worksheet('Employees').title).to eq 'Employees'
    end

    it 'raises for an unknown title or index' do
      expect { workbook.worksheet('Nope') }
        .to raise_error "Worksheet named 'Nope' not found in spreadsheet."
      expect { workbook.worksheet(5) }
        .to raise_error 'Worksheet with index 5 not found in spreadsheet.'
    end

    it 'fetches the spreadsheet metadata only once' do
      workbook.sheets
      workbook.worksheet(0)
      expect(service).to have_received(:get_spreadsheet).once
    end
  end

  describe SpreadsheetGoodies::GoogleDrive::Worksheet do
    let(:worksheet) { described_class.new(key, nil, 1, connector: connector) }

    it 'uses the first sheet when no title is given' do
      expect(worksheet.title).to eq 'Employees'
    end

    it 'pads rows and turns empty cells into nil, like a CSV export' do
      expect(worksheet.rows).to eq [
        ['Employee', 'Salary', 'Notes'],
        ['John', '10,000.00', nil],
        [nil, '12,000.50', ' remote '],
      ]
      expect(worksheet.header_row).to eq ['Employee', 'Salary', 'Notes']
      expect(worksheet.data_rows.map { |r| r['Salary'] }).to eq ['10,000.00', '12,000.50']
    end

    it 'strips values when strip_values_on_read is set' do
      SpreadsheetGoodies.configure { |c| c.strip_values_on_read = true }
      expect(worksheet[2]).to eq [nil, '12,000.50', 'remote']
    end

    it 'handles an empty sheet' do
      sheet = described_class.new(key, "Bob's tab", 1, connector: connector)
      expect(sheet.rows).to eq []
    end

    it 'sends pending writes in one batch on commit' do
      allow(service).to receive(:batch_update_values)

      worksheet[1]['Notes'] = 'on site'
      worksheet[2][0] = 'Mary'
      expect(service).not_to have_received(:batch_update_values)

      worksheet.commit_writes!

      expect(service).to have_received(:batch_update_values).once do |spreadsheet_key, request|
        expect(spreadsheet_key).to eq key
        expect(request.value_input_option).to eq 'USER_ENTERED'
        expect(request.data.map { |d| [d.range, d.values] }).to eq [
          ["'Employees'!C2", [['on site']]],
          ["'Employees'!A3", [['Mary']]],
        ]
      end
      expect(worksheet[2]['Employee']).to eq 'Mary'
    end

    it 'does not call the API when there is nothing to commit' do
      allow(service).to receive(:batch_update_values)
      worksheet.commit_writes!
      expect(service).not_to have_received(:batch_update_values)
    end

    it 'converts column numbers beyond Z' do
      allow(service).to receive(:batch_update_values)
      worksheet.write_to_cell(1, 28, 'x')
      worksheet.commit_writes!
      expect(service).to have_received(:batch_update_values) do |_, request|
        expect(request.data.first.range).to eq "'Employees'!AB1"
      end
    end
  end

  describe SpreadsheetGoodies::GoogleDrive::Connector do
    it 'returns a sheet as CSV' do
      expect(connector.read_worksheet_as_string(key, 'Employees'))
        .to eq "Employee,Salary,Notes\nJohn,\"10,000.00\",\n,\"12,000.50\", remote \n"
    end

    describe 'authentication' do
      it 'requires a refresh token for oauth2' do
        SpreadsheetGoodies.configure do |c|
          c.login_method = :oauth2
          c.google_client_id = 'id'
          c.client_secret = 'secret'
        end
        expect { described_class.new }.to raise_error(/Missing refresh_token.*generate_refresh_token/)
      end

      it 'builds oauth2 credentials with the refresh token' do
        SpreadsheetGoodies.configure do |c|
          c.login_method = :oauth2
          c.google_client_id = 'id'
          c.client_secret = 'secret'
          c.refresh_token = 'refresh'
        end
        auth = described_class.new.service.authorization
        expect(auth).to be_a Google::Auth::UserRefreshCredentials
        expect(auth.refresh_token).to eq 'refresh'
        expect(auth.scope).to eq described_class::SCOPES
      end

      it 'builds service account credentials from the JSON key' do
        key_json = {
          type: 'service_account',
          client_email: 'robot@example.iam.gserviceaccount.com',
          private_key: OpenSSL::PKey::RSA.new(2048).to_pem,
          project_id: 'project',
        }.to_json
        SpreadsheetGoodies.configure do |c|
          c.login_method = :service_account
          c.service_account_key_json = key_json
        end
        auth = described_class.new.service.authorization
        expect(auth).to be_a Google::Auth::ServiceAccountCredentials
        expect(auth.issuer).to eq 'robot@example.iam.gserviceaccount.com'
      end
    end

    describe '.generate_refresh_token' do
      let(:output) { StringIO.new }

      before do
        allow_any_instance_of(Google::Auth::UserRefreshCredentials).to receive(:fetch_access_token!)
        allow_any_instance_of(Google::Auth::UserRefreshCredentials)
          .to receive(:refresh_token).and_return('new-refresh-token')
      end

      def run_flow
        thread = Thread.new do
          described_class.generate_refresh_token(client_id: 'id', client_secret: 'secret', output: output)
        end
        sleep 0.01 until output.string.include?('Waiting')
        url = Addressable::URI.parse(output.string[%r{https://\S+}])
        [thread, url.query_values]
      end

      it 'receives the code on the loopback redirect and returns the refresh token' do
        thread, params = run_flow
        expect(params['redirect_uri']).to match %r{\Ahttp://127\.0\.0\.1:\d+\z}
        expect(params['access_type']).to eq 'offline'
        expect(params['scope']).to eq described_class::SCOPES.join(' ')

        redirect = URI(params['redirect_uri'])
        expect(Net::HTTP.get_response(redirect + '/favicon.ico').code).to eq '404'
        response = Net::HTTP.get_response(redirect + "/?state=#{params['state']}&code=abc")

        expect(response.body).to include 'Authorization received'
        expect(thread.value).to eq 'new-refresh-token'
      end

      it 'rejects a redirect with the wrong state' do
        thread, params = run_flow
        thread.report_on_exception = false
        Net::HTTP.get_response(URI(params['redirect_uri'] + '/?state=forged&code=abc'))
        expect { thread.value }.to raise_error 'Authorization failed: state mismatch.'
      end
    end
  end
end
