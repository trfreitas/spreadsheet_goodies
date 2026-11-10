# SpreadsheetGoodies

SpreadsheetGoodies is a collection of tools to help work with spreadsheets in
Excel and Google Spreadhseets formats. It relies heavily on other gems to make
the actual work of reading and writing to spreadsheet documents. It main
features are:

* Read a spreadseet as an array of arrays to allow aceessing its data without
  using the original document
* Access a row's elements using the column titles as keys"

## Installation

Add this line to your application's Gemfile:

```ruby
gem 'spreadsheet_goodies'
```

And then execute:

    $ bundle

Or install it yourself as:

    $ gem install spreadsheet_goodies

## Usage

Read a Google Spreadsheet:
```ruby
workbook = SpreadsheetGoodies::GoogleDrive.read_workbook(
  spreadsheet_key: '1UC43X6aZwlWPCnn...', # required, taken from the spreadsheet url
)
workbook.sheets            # => ['sheet1', 'sheet2']
sheet = workbook.worksheet # first worksheet
sheet = workbook.worksheet('sheet1') # or by title, or by zero-based index
```

Values are read as displayed in Google Sheets (formatted strings). Empty cells
are `nil`, and every row has the same length.

Read an Excel workbook:
```ruby
sheet = SpreadsheetGoodies::Excel.read_worksheet(
  file_pathname: '~/workbook.xlsx', # required,
  worksheet_title_or_index: 'sheet1', # optional, first worksheet is loaded if not specified
)
```

Iterate over every data row (i.e., all but the header row) and print the value
of a column titled 'Total':
```ruby
sheet.data_rows.each do |row|
  puts "#{row.row_number} -- #{row['Total']}"
end
```

Writing values to cells (only available for GoogleDrive adapter right now):
```ruby
row = sheet[0]
row[0] = 'First cell'
row[1] = 'Second cell'
sheet.commit_writes! # changes are applied to real spreadsheet
```

## Logging in to Google Drive
If you need to access a spreadsheet on Google Drive that is not publicly accessible,
you are required to setup an authentication method. Currently, there are two available
authentication methods.

Both methods use the Google Sheets API, which must be enabled in your project at
https://console.cloud.google.com/apis/library/sheets.googleapis.com. The only
scope requested is `https://www.googleapis.com/auth/spreadsheets`.

### OAuth2
1. At https://console.cloud.google.com/apis/credentials, create an OAuth client
   ID of type **Desktop app** and note its client id and client secret.
2. On a machine with a browser, generate a refresh token once:
   ```ruby
   require 'spreadsheet_goodies'
   SpreadsheetGoodies::GoogleDrive::Connector.generate_refresh_token(
     client_id: 'YOUR_CLIENT_ID',
     client_secret: 'YOUR_CLIENT_SECRET',
   )
   ```
   It prints an authorization URL and waits on a temporary local address
   (`http://127.0.0.1:<port>`). Open the URL, authorize access, and Google
   redirects back to that address. The refresh token is printed and returned.
3. Configure the gem with the three values:
   ```ruby
   SpreadsheetGoodies.configure do |config|
     config.login_method = :oauth2
     config.google_client_id = ENV['GOOGLE_CLIENT_ID']
     config.client_secret = ENV['GOOGLE_CLIENT_SECRET']
     config.refresh_token = ENV['GOOGLE_REFRESH_TOKEN']
   end
   ```

Refresh tokens obtained with versions before 0.2.0 keep working.

### Service Accounts
Share the spreadsheet with the service account's e-mail address, then:
```ruby
SpreadsheetGoodies.configure do |config|
  config.login_method = :service_account
  config.service_account_key_json = File.read('service_account_key.json')
end
```

## Upgrading to 0.2.0
* Requires Ruby >= 3.1.
* The `google_drive` gem was replaced by `google-apis-sheets_v4` and
  `googleauth ~> 1.9`. `Connector` no longer exposes a `GoogleDrive::Session`;
  `Connector#read_workbook` and `Workbook#workbook` now return a
  `Google::Apis::SheetsV4::Spreadsheet`, and `Connector#service` returns the
  `Google::Apis::SheetsV4::SheetsService`. `Connector#read_worksheet` and
  `Connector#worksheets` were removed.
* `Workbook#worksheet` accepts a zero-based index (it used to fail with one).
* The interactive refresh token prompt was replaced by
  `Connector.generate_refresh_token`; a missing refresh token now raises an error.
* `axlsx` was replaced by `caxlsx`.

## Contributing

Bug reports and pull requests are welcome on GitHub at https://github.com/ricardo-jasinski/spreadsheet_goodies.


## License

The gem is available as open source under the terms of the [Unlicense](http://unlicense.org/UNLICENSE).
