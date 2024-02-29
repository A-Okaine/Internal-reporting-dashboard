# Internal-reporting-dashboard

It was requested that a shift report be created for an internal Team to track daily tasks.

This scope of the code was to:
1. Excel to Google Sheets Conversion:
- Searches for unread emails labeled "data-reporters-responsible-play-dashboard."
- Extract attachments from these emails, which are assumed to be Excel files.
- Parses the data from these Excel files.
- Insert the parsed data into corresponding sheets in a Google Spreadsheet

  2. Folder Management:
- Creates folders in Google Drive based on the current date.
- Moves the Google Spreadsheet containing processed data to the appropriate folder.

3. Sending Email:
- Sends an email using Gmail service based on the configured template.
- The email includes a link to a Google Sheet.

4. Data Cleanup:
- Clears specific sheets' content in the original spreadsheet after processing to maintain data integrity.
