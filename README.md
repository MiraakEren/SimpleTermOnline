# SimpleTermOnline: Collaborative Term Base Management

**SimpleTermOnline** is an enhanced version of SimpleTerm, designed to streamline your interaction with term bases by offering a minimalistic interface that reduces the need for cumbersome spreadsheets.

## Getting Started

1. **Download the Executable**
   - Obtain the `.exe` file from the link provided below.
  1.1 Optional Step
     Place the executable in a dedicated folder with the key file to keep the files organized.

2. **Enter Sheet ID**
   - Input the Sheet ID, which can be found in the URL between "d/" and "/edit?...".

3. **Select JSON Key File**
   - Choose the JSON key file that grants access to the Google Sheets API.

## Features

**SimpleTermOnline** is engineered to **efficiently manage and access term bases collaboratively** with a user interface optimized for simplicity and minimal mouse usage.

- **API Key and Sheet ID Required**: To access your Google Sheet, you must provide both the API key and Sheet ID. Ensure your sheet includes "Source Term", "Target Term", and "Notes" as headers in the first row.

- **Local Data Handling**: The tool does not store user data online, and therefore cannot verify usernames. It is designed for small groups and maintains a record of term entries and reviews within the spreadsheet, minimizing on-screen clutter.

### Key Features:

- **Efficient Term Management**: Add or edit terms quickly without opening the spreadsheet.
- **Search Navigation**: Navigate through multiple results for a search term seamlessly.
- **Direct Download**: Download the active spreadsheet directly from the application.
- **Copy Functionality**: Easily copy target terms to your clipboard.
- **Reviewer Mode**: Review and confirm entries to ensure accuracy and trustworthiness.

### Libraries Used:

- **pandas**: For interacting with sheets.
- **tkinter**: For the graphical user interface (GUI).
- **pyperclip**: For copying terms to the clipboard.
- **gspread**: For Google Sheets API interactions.

## Keyboard Shortcuts:

- **Tab Key, Left/Right Arrows**: Navigate between results.
- **Ctrl+C**: Copy the displayed result.
- **Ctrl+N**: Add a new term.
- **Ctrl+O**: Open the current Google Sheet.
- **Ctrl+E**: Edit the current entry.
- **Ctrl+D**: Download the sheet as a `.csv` file.
- **Ctrl+R**: Open Reviewer Mode.
- **Ctrl+ +/-**: Adjust font size.
- **Ctrl+Shift+ +/-**: Adjust notes font size.
- **F5**: Refresh the Google Sheet.
- **F3**: Display the shortcut help menu.
- **F2**: Change the active sheet.

## Additional Notes:

**Upcoming Features**: The next update will introduce enhancements to Reviewer Mode, adding new functionality and improvements.
