Google Drive Search Tool

This tool provides a convenient way to search for documents within a specified Google Drive folder and Google Sheet. It allows you to link URLs to the Google Sheet for efficient organization and management of your files.
Code Optimization

    .Memoization: The program utilizes memoization to cache folder structures and values retrieved from Google APIs. This improves performance by reducing redundant API calls.
    .Efficient API Usage: API requests are made efficiently, minimizing unnecessary data retrieval and processing.
    .Modular Design: The code is organized into functions, making it easier to understand, maintain, and extend.
    .Error Handling: Comprehensive error handling ensures robustness and prevents crashes due to unexpected inputs or API failures.
    .GUI Responsiveness: The graphical user interface (GUI) is responsive and provides real-time feedback during operations.

Usage

    1.Select Google Drive Folder: Click the "Select Folder" button to choose the Google Drive folder you want to search within.
    2.Select Google Sheet: Enter the link to the Google Sheet and select the desired tab.
    3.Specify Columns: Enter the column names for document IDs and phone numbers in the Google Sheet.
    4.Search: Click the "Search" button to start searching for documents within the specified folder and update the results in the GUI.
    5.Link URLs: After performing a search, you can click the "Link URLs" button to link the URLs of matching documents to the Google Sheet.
    6.Clear Results: Click the "Clear Results" button to clear the search results displayed in the GUI.

Google APIs

This tool utilizes the following Google APIs:

    .Google Drive API: Allows access to files and folders stored in Google Drive.
    .Google Sheets API: Enables reading data from and writing data to Google Sheets.

Dependencies

Ensure the following dependencies are installed:

    .google-auth
    .google-auth-oauthlib
    .google-auth-httplib2
    .google-api-python-client
    .tkinter

Setup

    1.Obtain OAuth 2.0 credentials by creating a project in the Google Developers Console and enabling the Google Drive API and Google Sheets API.
    2.Download the credentials.json file and place it in the same directory as the script.
    3.Run the script and follow the authentication prompts to authorize access to your Google account.

Code Overview

The code consists of several functions:

  1.  authenticate(): Authenticates with Google APIs using OAuth 2.0. It checks for existing credentials or prompts the user to log in if necessary.

  2.  get_data(sheet_id, tab_name): Retrieves all values from the specified Google Sheet tab using the Google Sheets API.

 3.   select_folder(): Opens a dialog to select a Google Drive folder. It utilizes the Google Drive API to list available folders.

 4.   list_folders(service): Lists available folders in Google Drive along with the total number of documents. It caches the folder list for performance optimization.

 5.   list_files(service, folder_id): Lists all files in a Google Drive folder using pagination. It caches the folder files list for performance optimization.

 6.   select_sheet(): Opens a dialog to select a Google Sheet from a list. It utilizes the Google Drive API to list Google Sheets.

 7.   list_google_sheets(service): Lists all Google Sheets in Google Drive using the Google Drive API.

8.    list_tabs(service, sheet_id): Lists tabs of a Google Sheet using the Google Sheets API.

9.    column_to_letter(column): Converts a column number to its corresponding letter representation (e.g., 1 to 'A', 2 to 'B', etc.).

10.    list_columns(service, sheet_id, tab_name): Lists columns of a Google Sheets tab using the Google Sheets API. It caches the column information for performance optimization.

11.    select_columns(service, sheet_id): Opens a dialog to select columns for IDs and phone numbers from a Google Sheet. It utilizes the Google Sheets API to list columns.

12.    start_search(): Initiates the search process by calling search_and_update_drive(). It retrieves user inputs and calls necessary functions.

13.    search_and_update_drive(folder_id, column_name, phone_column, sheet_link, sheet_name, tab_name): Searches Google Drive for matching documents and updates the hits in the GUI. It utilizes authentication, Google Drive API, and Google Sheets API for data retrieval.

14.    clear_results(): Clears search results from the GUI.

15.    get_non_empty_columns(sheet_id, tab_name): Retrieves non-empty columns in the specified Google Sheet tab using the Google Sheets API.

16.    link(): Initiates the process of linking URLs to empty columns in the Google Sheet. It retrieves user inputs and calls necessary functions.

17.    link_urls(sheet_id, tab_name, columns_to_fill): Pastes URLs from the hit log into specified empty columns in the Google Sheet. It utilizes the Google Sheets API for batch updates.

18.    on_closing(): Handles the closing event of the application window.

Notes

    .This tool requires proper authentication to access your Google Drive and Google Sheets data.
    .Ensure that the specified folder and Google Sheet exist and that you have appropriate permissions to access them.
    .Double-clicking on a URL in the search results copies the URL to the clipboard for easy sharing and navigation.

