require 'csv'
require 'logger'
require 'securerandom'
require 'socket'
require 'stringio'
require 'uri'
require 'googleauth'
require 'google/apis/sheets_v4'

module SpreadsheetGoodies::GoogleDrive
  # Thin wrapper around the Google Sheets API v4
  # (https://developers.google.com/sheets/api/reference/rest).
  class Connector
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets'].freeze

    attr_reader :logger, :service

    # @param service [Google::Apis::SheetsV4::SheetsService] Optional; an
    #   already authorized service. If omitted, one is built from
    #   SpreadsheetGoodies.configuration.
    def initialize(service: nil)
      @logger = Logger.new(STDOUT)
      @service = service || build_service
      @workbooks = {}
    end

    # Returns the spreadsheet metadata (no cell data) as a
    # Google::Apis::SheetsV4::Spreadsheet.
    def read_workbook(spreadsheet_key)
      @workbooks[spreadsheet_key] ||= @service.get_spreadsheet(
        spreadsheet_key, fields: 'spreadsheetId,properties.title,sheets.properties'
      )
    end

    def sheet_titles(spreadsheet_key)
      read_workbook(spreadsheet_key).sheets.map { |sheet| sheet.properties.title }
    end

    # @param title_or_index [String, Integer, nil] Sheet title, zero-based
    #   index, or nil for the first sheet.
    # @return [String] The sheet title
    def resolve_sheet_title(spreadsheet_key, title_or_index)
      titles = sheet_titles(spreadsheet_key)

      case title_or_index
      when nil
        titles.first or raise 'Spreadsheet has no worksheets.'
      when Integer
        titles[title_or_index] or
          raise "Worksheet with index #{title_or_index} not found in spreadsheet."
      else
        title = title_or_index.to_s
        titles.include?(title) or
          raise "Worksheet named '#{title}' not found in spreadsheet."
        title
      end
    end

    # Reads all cells of a sheet as displayed in the UI (formatted values).
    # Empty cells are nil and every row has the same length, matching a CSV
    # export of the sheet.
    # @return [Array<Array<String, nil>>]
    def read_values(spreadsheet_key, sheet_title)
      response = @service.get_spreadsheet_values(
        spreadsheet_key, quote_sheet_title(sheet_title),
        value_render_option: 'FORMATTED_VALUE'
      )
      rows = response.values || []
      width = rows.map(&:length).max || 0

      rows.map do |row|
        row = row.map { |value| value == '' ? nil : value }
        row + [nil] * (width - row.length)
      end
    end

    # Reads a sheet and returns it as a CSV string
    def read_worksheet_as_string(spreadsheet_key, worksheet_title_or_index = nil)
      title = resolve_sheet_title(spreadsheet_key, worksheet_title_or_index)
      CSV.generate { |csv| read_values(spreadsheet_key, title).each { |row| csv << row } }
    end

    # Writes several cells in a single request. Values are interpreted as if
    # typed by a user (e.g. '=SUM(A1:A3)' becomes a formula).
    # @param updates [Hash{Array(Integer, Integer) => Object}] Maps
    #   [row_index, col_index] (both starting at 1) to the new value.
    def write_values(spreadsheet_key, sheet_title, updates)
      return if updates.empty?

      data = updates.map do |(row_index, col_index), value|
        Google::Apis::SheetsV4::ValueRange.new(
          range: "#{quote_sheet_title(sheet_title)}!#{column_letter(col_index)}#{row_index}",
          values: [[value]]
        )
      end
      request = Google::Apis::SheetsV4::BatchUpdateValuesRequest.new(
        value_input_option: 'USER_ENTERED', data: data
      )
      @service.batch_update_values(spreadsheet_key, request)
    end

    # Runs the OAuth2 authorization flow for a "Desktop app" OAuth client and
    # returns a refresh token. It opens a temporary HTTP server on 127.0.0.1
    # to receive Google's redirect (loopback flow), so it must run on a
    # machine with a browser.
    # @return [String] The refresh token
    def self.generate_refresh_token(client_id:, client_secret:, port: 0, output: $stdout)
      server = TCPServer.new('127.0.0.1', port)
      redirect_uri = "http://127.0.0.1:#{server.addr[1]}"
      state = SecureRandom.hex(16)

      credentials = Google::Auth::UserRefreshCredentials.new(
        client_id: client_id, client_secret: client_secret,
        scope: SCOPES, redirect_uri: redirect_uri
      )
      url = credentials.authorization_uri(access_type: 'offline', prompt: 'consent', state: state)
      output.puts "Open this page in your browser and authorize access:\n#{url}\n\n"
      output.puts "Waiting for the authorization on #{redirect_uri} ..."

      code = wait_for_authorization_code(server, state)
      credentials.code = code
      credentials.fetch_access_token!
      output.puts "Your refresh token is: #{credentials.refresh_token}"
      output.puts 'Set it as refresh_token in your SpreadsheetGoodies configuration.'
      credentials.refresh_token
    ensure
      server&.close
    end

    def self.wait_for_authorization_code(server, state)
      loop do
        client = server.accept
        request_line = client.gets.to_s
        path = request_line.split(' ')[1].to_s
        params = URI.decode_www_form(URI(path).query.to_s).to_h

        # Ignore unrelated requests such as /favicon.ico
        unless params.key?('code') || params.key?('error')
          client.print "HTTP/1.1 404 Not Found\r\nContent-Length: 0\r\nConnection: close\r\n\r\n"
          client.close
          next
        end

        ok = params['state'] == state && params['code']
        body = ok ? 'Authorization received. You can close this window.' : 'Authorization failed.'
        client.print "HTTP/1.1 200 OK\r\nContent-Type: text/plain\r\n" \
          "Content-Length: #{body.bytesize}\r\nConnection: close\r\n\r\n#{body}"
        client.close

        raise "Authorization failed: #{params['error']}" if params['error']
        raise 'Authorization failed: state mismatch.' unless params['state'] == state
        return params['code']
      end
    end
    private_class_method :wait_for_authorization_code

  private

    def build_service
      service = Google::Apis::SheetsV4::SheetsService.new
      service.authorization = credentials
      service
    end

    def credentials
      config = SpreadsheetGoodies.configuration
      case config.login_method
      when :service_account then service_account_credentials(config)
      when :oauth2 then oauth2_credentials(config)
      else raise "Unknown login_method: #{config.login_method.inspect}"
      end
    end

    # Authenticate via OAuth2 with a previously obtained refresh token.
    # Your OAuth2 client id and secret are at
    # https://console.cloud.google.com/apis/credentials. To obtain a refresh
    # token, run SpreadsheetGoodies::GoogleDrive::Connector.generate_refresh_token.
    def oauth2_credentials(config)
      unless config.refresh_token
        raise 'Missing refresh_token in SpreadsheetGoodies configuration. Obtain one with ' \
          'SpreadsheetGoodies::GoogleDrive::Connector.generate_refresh_token(' \
          'client_id: ..., client_secret: ...)'
      end

      Google::Auth::UserRefreshCredentials.new(
        client_id: config.google_client_id,
        client_secret: config.client_secret,
        refresh_token: config.refresh_token,
        scope: SCOPES
      )
    end

    # Authenticate via a service account. The spreadsheet must be shared with
    # the service account's e-mail address.
    def service_account_credentials(config)
      Google::Auth::ServiceAccountCredentials.make_creds(
        json_key_io: StringIO.new(config.service_account_key_json),
        scope: SCOPES
      )
    end

    def quote_sheet_title(title)
      "'#{title.gsub("'", "''")}'"
    end

    # 1 -> A, 26 -> Z, 27 -> AA
    def column_letter(index)
      letters = ''
      while index > 0
        index, remainder = (index - 1).divmod(26)
        letters = (65 + remainder).chr + letters
      end
      letters
    end
  end
end
