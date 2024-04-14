import os
import re
import tkinter as tk
from tkinter import messagebox, ttk, simpledialog
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# Define Google Drive API and Google Sheets API scopes
SCOPES = ['https://www.googleapis.com/auth/drive.readonly', 'https://www.googleapis.com/auth/spreadsheets']

# Memoization cache
memo = {}
urls = []  # Initialize the list to store URLs

def authenticate():
    """Authenticate with Google APIs."""
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json')
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            # Construct the path to credentials.json dynamically
            script_dir = os.path.dirname(os.path.realpath(__file__))
            credentials_path = os.path.join(script_dir, 'credentials.json')
            flow = InstalledAppFlow.from_client_secrets_file(credentials_path, SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    return creds
def get_data(sheet_id, tab_name):
    """Retrieve all values in the specified Google Sheet tab."""
    creds = authenticate()
    service = build('sheets', 'v4', credentials=creds)
    range_name = f"{tab_name}!A:Z"
    result = service.spreadsheets().values().get(spreadsheetId=sheet_id, range=range_name).execute()
    return result.get('values', [])
def select_folder():
    """Select Google Drive folder."""
    creds = authenticate()
    service = build('drive', 'v3', credentials=creds)
    folders, total_documents = list_folders(service)
    if folders:
        folder_names = [folder['name'] for folder in folders]
        selected_folder = simpledialog.askstring("Select Folder", f"Select Google Drive Folder\nTotal Documents: {total_documents}", initialvalue=folder_names[0],
                                                 parent=root)
        if selected_folder:
            folder_id = [folder['id'] for folder in folders if folder['name'] == selected_folder][0]
            folder_entry.delete(0, tk.END)
            folder_entry.insert(tk.END, folder_id)
    else:
        messagebox.showerror("Error", "No folders found in Google Drive.")

def list_folders(service):
    """List available folders in the Google Drive along with the total number of documents."""
    # Check if folders list is already cached
    if 'folders' in memo:
        return memo['folders']

    results = service.files().list(
        q="mimeType='application/vnd.google-apps.folder'",
        fields="files(id, name)").execute()
    folders = results.get('files', [])
    folder_count = len(folders)

    # Iterate over each folder to count the total number of documents
    total_documents = 0
    for folder in folders:
        folder_id = folder['id']
        # Reset the total_documents count for each folder iteration
        documents_result = list_files(service, folder_id)
        total_documents += len(documents_result)/2

    # Cache the folders list
    memo['folders'] = (folders, total_documents)
    return memo['folders']

def list_files(service, folder_id):
    """List all files in the Google Drive folder."""
    # Check if folder files list is already cached
    if folder_id in memo:
        return memo[folder_id]

    files = []
    page_token = None
    while True:
        response = service.files().list(q=f"'{folder_id}' in parents",
                                        fields="nextPageToken, files(id, name)",
                                        pageToken=page_token).execute()
        files.extend(response.get('files', []))
        page_token = response.get('nextPageToken')
        if not page_token:
            break

    # Cache the folder files list
    memo[folder_id] = files
    return files

def select_sheet():
    """Select Google Sheet from a list."""
    creds = authenticate()
    service = build('drive', 'v3', credentials=creds)
    sheets = list_google_sheets(service)
    if sheets:
        sheet_names = [sheet['name'] for sheet in sheets]
        sheet_list_window = tk.Toplevel(root)
        sheet_list_window.title("Select Google Sheet")
        sheet_list_window.geometry("300x200")
        selected_sheet = tk.StringVar(value=sheet_names[0])
        sheet_listbox = tk.Listbox(sheet_list_window, listvariable=selected_sheet, selectmode="single")
        sheet_listbox.pack(expand=True, fill="both")
        for sheet_name in sheet_names:
            sheet_listbox.insert(tk.END, sheet_name)

        selected_columns = {'IDs': None, 'Phone number': None}

        def on_ok():
            selected_sheet_name = sheet_listbox.get(tk.ACTIVE)
            sheet_entry.delete(0, tk.END)
            sheet_entry.insert(tk.END, selected_sheet_name)  # Update Sheet Name entry field

            sheet = [sheet for sheet in sheets if sheet['name'] == selected_sheet_name][0]
            sheet_id = sheet['id']
            sheet_entry.delete(0, tk.END)
            sheet_entry.insert(tk.END, f"https://docs.google.com/spreadsheets/d/{sheet_id}")
            service = build('sheets', 'v4', credentials=creds)  # Use Google Sheets API
            tabs = list_tabs(service, sheet_id)
            if tabs:
                select_tab_window = tk.Toplevel(root)
                select_tab_window.title("Select Tab")
                select_tab_window.geometry("300x200")
                selected_tab = tk.StringVar(value=tabs[0])
                tab_listbox = tk.Listbox(select_tab_window, listvariable=selected_tab, selectmode="single")
                tab_listbox.pack(expand=True, fill="both")
                for tab_name in tabs:
                    tab_listbox.insert(tk.END, tab_name)

                def on_tab_ok():
                    selected_tab_name = tab_listbox.get(tk.ACTIVE)
                    tab_entry.delete(0, tk.END)
                    tab_entry.insert(tk.END, selected_tab_name)
                    # Now let's list the columns for the selected tab
                    columns = list_columns(service, sheet_id, selected_tab_name)
                    if columns:
                        select_columns(service, sheet_id)
                    else:
                        messagebox.showerror("Error", "No columns found in the selected Google Sheets tab.")

                    select_tab_window.destroy()

                tab_ok_button = tk.Button(select_tab_window, text="OK", command=on_tab_ok)
                tab_ok_button.pack()

            else:
                messagebox.showerror("Error", "No tabs found in the selected Google Sheet.")

            sheet_list_window.destroy()

        ok_button = tk.Button(sheet_list_window, text="OK", command=on_ok)
        ok_button.pack()

    else:
        messagebox.showerror("Error", "No Google Sheets found in Google Drive.")

def list_google_sheets(service):
    """List all Google Sheets in Google Drive."""
    results = service.files().list(
        q="mimeType='application/vnd.google-apps.spreadsheet'",
        fields="files(id, name)").execute()
    sheets = results.get('files', [])
    return sheets

def list_tabs(service, sheet_id):
    """List tabs of a Google Sheet."""
    sheet = service.spreadsheets().get(spreadsheetId=sheet_id).execute()
    sheets = sheet.get('sheets', [])
    tabs = [sheet['properties']['title'] for sheet in sheets]
    return tabs

def column_to_letter(column):
    """Convert column number to letter."""
    letters = ''
    while column > 0:
        column, remainder = divmod(column - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters

def list_columns(service, sheet_id, tab_name):
    """List columns of a Google Sheets tab."""
    # Get the spreadsheet
    sheet = service.spreadsheets().get(spreadsheetId=sheet_id).execute()
    sheets = sheet.get('sheets', [])

    # Find the specified tab by name
    tab = None
    for s in sheets:
        if s['properties']['title'] == tab_name:
            tab = s
            break

    if not tab:
        messagebox.showerror("Error", f"Tab '{tab_name}' not found in the Google Sheet.")
        return []

    # Get the total number of columns in the tab
    total_columns = tab['properties']['gridProperties']['columnCount']

    # Construct the range from column A to the last column
    range_name = f"{tab_name}!A1:{column_to_letter(total_columns)}1"

    # Retrieve values from the specified range to get column names
    result = service.spreadsheets().values().get(spreadsheetId=sheet_id, range=range_name).execute()
    values = result.get('values', [])
    columns = values[0] if values else []

    # Create a list of tuples containing column name and its letter identifier
    column_info = [(col_name, column_to_letter(idx + 1)) for idx, col_name in enumerate(columns)]

    return column_info

def select_columns(service, sheet_id):
    """Select columns for IDs and phone numbers."""
    select_column_window = tk.Toplevel(root)
    select_column_window.title("Select Columns")
    select_column_window.geometry("300x200")
    selected_columns = []

    def on_column_ok_id():
        selected_column_value = column_listbox.get(tk.ACTIVE)  # Get the selected column value
        if selected_column_value:
            # Extract the column letter from the selected column value
            selected_column_letter = re.search(r'\((\w+)\)', selected_column_value).group(1)
            selected_columns.append(selected_column_letter)
            if len(selected_columns) == 1:
                column_entry.delete(0, tk.END)
                column_entry.insert(tk.END, selected_column_value.split()[0])
                select_column_window.destroy()
                # Call select_columns again to select the second column
                select_columns(service, sheet_id)
            elif len(selected_columns) == 2:
                phone_entry.delete(0, tk.END)
                phone_entry.insert(tk.END, selected_column_value.split()[0])
                select_column_window.destroy()
            else:
                messagebox.showerror("Error", "Please select only two columns.")

    def on_column_ok_phone():
        selected_column_value = column_listbox.get(tk.ACTIVE)  # Get the selected column value
        if selected_column_value:
            # Extract the column letter from the selected column value
            selected_column_letter = re.search(r'\((\w+)\)', selected_column_value).group(1)
            selected_columns.append(selected_column_letter)
            if len(selected_columns) == 1:
                column_entry.delete(0, tk.END)
                column_entry.insert(tk.END, selected_column_value.split()[0])
                select_column_window.destroy()
                # Call select_columns again to select the second column
                select_columns(service, sheet_id)
            elif len(selected_columns) == 2:
                phone_entry.delete(0, tk.END)
                phone_entry.insert(tk.END, selected_column_value.split()[0])
                select_column_window.destroy()
            else:
                messagebox.showerror("Error", "Please select only two columns.")

    columns = list_columns(service, sheet_id, tab_entry.get())
    column_listbox = tk.Listbox(select_column_window, selectmode="single")
    column_listbox.pack(expand=True, fill="both")
    for column_info in columns:
        column_listbox.insert(tk.END, f"{column_info[0]} ({column_info[1]})")  # Display name and letter

    column_ok_button = tk.Button(select_column_window, text="OK", command=on_column_ok_id)
    column_ok_button.pack()

def start_search():
    """Start searching Google Drive."""
    folder_id = folder_entry.get()
    column_name = column_entry.get()
    phone_column = phone_entry.get()
    sheet_link = sheet_entry.get()
    sheet_name = sheet_entry.get()
    tab_name = tab_entry.get()
    search_and_update_drive(folder_id, column_name, phone_column, sheet_link, sheet_name, tab_name)

def search_and_update_drive(folder_id, column_name, phone_column, sheet_link, sheet_name, tab_name):
    """Search Google Drive for matching documents and update the hits in the GUI."""
    creds = authenticate()
    service = build('drive', 'v3', credentials=creds)

    # Retrieve all files in the specified folder
    files = list_files(service, folder_id)

    # Initialize hit count
    hit_count = 0

    # Retrieve the selected sheet
    sheet_id = re.search(r'/spreadsheets/d/([a-zA-Z0-9-_]+)', sheet_link)
    if sheet_id:
        sheet_id = sheet_id.group(1)
    else:
        messagebox.showerror("Error", "Invalid Google Sheet link.")
        return

    service = build('sheets', 'v4', credentials=creds)

    # Retrieve data from the selected sheet and tab
    id_range = f"{tab_name}!{column_name}:{column_name}"
    phone_range = f"{tab_name}!{phone_column}:{phone_column}"

    id_data = service.spreadsheets().values().get(spreadsheetId=sheet_id, range=id_range).execute()
    phone_data = service.spreadsheets().values().get(spreadsheetId=sheet_id, range=phone_range).execute()

    ids = id_data.get('values', [])
    phones = phone_data.get('values', [])

    if ids:
        # Extract the important numbers from the document names
        document_numbers = [re.search(r'(\d+)#', file['name']).group(1) for file in files if re.search(r'(\d+)#', file['name'])]

        # Extract IDs and phone numbers from the Google Sheet
        sheet_ids = [row[0].lstrip('0') if row else "" for row in ids]  # Remove leading zeros
        phone_numbers = [row[0] if row else "" for row in phones]  # Use phone number if ID is blank

        # Count hits
        for doc_number in document_numbers:
            if doc_number.lstrip('0') in sheet_ids or doc_number.lstrip('0') in phone_numbers:
                hit_count += 1

                # Get the URL of the matching file
                matching_file = next(file for file in files if re.search(rf'{doc_number}#', file['name']))
                file_url = f"https://drive.google.com/file/d/{matching_file['id']}/view"

                # Append the URL to the list
                urls.append((matching_file['name'], file_url))

    # Display the hit count in the GUI
    hit_count_label.config(text=f"Total Hits: {hit_count}")

    # Clear the Treeview
    tree.delete(*tree.get_children())

    # Display the hit files in the Treeview
    for idx, url_info in enumerate(urls):
        file_name, file_url = url_info
        index_match = re.search(r'^(\d+)#', file_name)  # Extract index using regex
        index = index_match.group(1) if index_match else ""  # Extracted index or empty string if not found

        # Find the column name and row in the Google Sheets where the index appears
        gs_name = ""
        if index:
            # Remove leading zeros from the index
            index_without_zero = index.lstrip('0')
            row = None
            for i, (row_value1, row_value2) in enumerate(zip(ids, phones), start=1):
                if row_value1 and index_without_zero in row_value1:
                    row = i
                    gs_name = f"{column_name}{row}"  # Construct the Google Sheets column name and row
                    break
                elif row_value2 and index_without_zero in row_value2:
                    row = i
                    gs_name = f"{phone_column}{row}"  # Construct the Google Sheets column name and row
                    break

        # Insert the values into the Treeview
        tree.insert("", tk.END, values=(index, file_name, gs_name, file_url))

def clear_results():
    """Clear search results."""
    tree.delete(*tree.get_children())
    hit_count_label.config(text="Total Hits: 0")
    urls.clear()

def get_non_empty_columns(sheet_id, tab_name):
    """Retrieve non-empty columns in the specified Google Sheet tab."""
    creds = authenticate()
    service = build('sheets', 'v4', credentials=creds)

    # Define the range to fetch only the first row of each column
    range_name = f"{tab_name}!1:1"
    
    # Make the request to retrieve the values
    result = service.spreadsheets().values().batchGet(spreadsheetId=sheet_id, ranges=[range_name]).execute()
    value_ranges = result.get('valueRanges', [])

    non_empty_columns = []
    if value_ranges:
        for value_range in value_ranges:
            values = value_range.get('values', [])
            if values:
                for col_idx, cell in enumerate(values[0], start=1):
                    if cell:  # If cell is not empty
                        non_empty_columns.append(column_to_letter(col_idx))
    return non_empty_columns

def link():
    """Link URLs to empty columns in the Google Sheet."""
    # Get the selected Google Sheet and tab
    sheet_id_match = re.search(r'/spreadsheets/d/([a-zA-Z0-9-_]+)', sheet_entry.get())
    if sheet_id_match:
        sheet_id = sheet_id_match.group(1)
    else:
        messagebox.showerror("Error", "Invalid Google Sheet link.")
        return

    tab_name = tab_entry.get()

    # Get non-empty columns
    non_empty_columns = get_non_empty_columns(sheet_id, tab_name)

    if non_empty_columns:
        # Prompt user to select an empty column using GUI
        selected_column = simpledialog.askstring("Select Empty Column", "Enter the column letter where URLs will be pasted (e.g., A, B, C):")
        if selected_column:
            # Check if the selected column is valid
            if selected_column.upper() not in non_empty_columns:
                # Link URLs to the selected column
                link_urls(sheet_id, tab_name, [selected_column.upper()])
            else:
                messagebox.showerror("Error", "Selected column is not empty. Please select an empty column.")
    else:
        messagebox.showerror("Error", "No non-empty columns found in the selected Google Sheet.")

def list_files(service, folder_id, page_token=None):
    """List all files in the Google Drive folder with pagination."""
    # Check if folder files list is already cached
    if folder_id in memo:
        return memo[folder_id]

    files = []
    while True:
        response = service.files().list(q=f"'{folder_id}' in parents",
                                        fields="nextPageToken, files(id, name)",
                                        pageToken=page_token).execute()
        files.extend(response.get('files', []))
        page_token = response.get('nextPageToken')
        if not page_token:
            break

    # Cache the folder files list
    memo[folder_id] = files
    return files

"""def link_urls(sheet_id, tab_name, columns_to_fill):
    
    creds = authenticate()
    service = build('sheets', 'v4', credentials=creds)
    
    # Display notification when URL pasting starts
    messagebox.showinfo("URL Pasting", "URL pasting process started.")
    
    # Get the hit log URLs
    urls_to_link = [url[1] for url in urls]

    # Get the existing data in the Google Sheet
    range_name = f"{tab_name}!A:Z"
    values_cache_key = (sheet_id, range_name)  # Cache key for values
    if values_cache_key in memo:
        values = memo[values_cache_key]  # Retrieve values from cache
    else:
        result = service.spreadsheets().values().get(spreadsheetId=sheet_id, range=range_name).execute()
        values = result.get('values', [])
        memo[values_cache_key] = values  # Cache the values
    
    print("Existing values in the Google Sheet:", values)  # Debug print
    
    if values:
        # Prepare batch update request
        batch_update_values_request = {
            'value_input_option': 'RAW',
            'data': []
        }

        # Iterate over each URL and add it to the batch update request
        for idx, url in enumerate(urls_to_link):
            # Find the first empty row in the specified column
            for col in columns_to_fill:
                col_idx = ord(col.upper()) - ord('A')  # Convert column letter to index
                column_data = [row[col_idx] for row in values if len(row) > col_idx]  # Get data for the specified column

                print(f"Column {col}: Data = {column_data}")  # Debug print

                if not any(column_data):  # If column is empty
                    # Find the first empty row in the column
                    empty_row_index = len(values) + 1

                    print(f"Column {col}: Empty row index = {empty_row_index}")  # Debug print

                    # Add URL to batch update request
                    range_to_update = f"{tab_name}!{col}{empty_row_index}"
                    batch_update_values_request['data'].append({
                        'range': range_to_update,
                        'values': [[url]]
                    })

                    # Add a green check button beside the pasted URL
                    tree.insert('', 'end', values=(url, ""), tags=("GREEN_BUTTON",))
                    break

        print("Batch update request:", batch_update_values_request)  # Debug print

        # Execute batch update request
        if batch_update_values_request['data']:
            service.spreadsheets().values().batchUpdate(spreadsheetId=sheet_id, body=batch_update_values_request).execute()

    # Apply styles
    tree.tag_configure("GREEN_BUTTON", background="green")
    
    # Display notification when URL pasting finishes
    messagebox.showinfo("URL Pasting", "URL pasting process finished.")"""
def link_urls(sheet_id, tab_name, columns_to_fill):
    #Paste URLs from the hit log into the specified empty columns in the Google Sheet.
    creds = authenticate()
    service = build('sheets', 'v4', credentials=creds)
    
    # Display notification when URL pasting starts
    messagebox.showinfo("URL Pasting", "URL pasting process started.")
    
    # Get the hit log URLs
    urls_to_link = [url[1] for url in urls]

    # Get the existing data in the Google Sheet
    range_name = f"{tab_name}!A:Z"
    values_cache_key = (sheet_id, range_name)  # Cache key for values
    if values_cache_key in memo:
        values = memo[values_cache_key]  # Retrieve values from cache
    else:
        result = service.spreadsheets().values().get(spreadsheetId=sheet_id, range=range_name).execute()
        values = result.get('values', [])
        memo[values_cache_key] = values  # Cache the values
    
    #print("Existing values in the Google Sheet:", values)  # Debug print
    
    if values:
        # Prepare batch update request
        batch_update_values_request = {
            'value_input_option': 'RAW',
            'data': []
        }

        # Iterate over each URL and add it to the batch update request
        for idx, url in enumerate(urls_to_link):
            gs_name = tree.item(tree.get_children()[idx], 'values')[2]  # Get the G-S Name from the Treeview
            if gs_name:
                # Extract the row number from G-S Name (e.g., 'D23' -> row 23)
                row_number = int(''.join(filter(str.isdigit, gs_name)))
                # Iterate over the specified columns to fill
                for col in columns_to_fill:
                    # Add URL to batch update request
                    range_to_update = f"{tab_name}!{col}{row_number}"
                    batch_update_values_request['data'].append({
                        'range': range_to_update,
                        'values': [[url]]
                    })

                    # Add a green check button beside the pasted URL
                    #tree.insert('', 'end', values=(url, ""), tags=("GREEN_BUTTON",))
                    tree.item(tree.get_children()[idx], values=("✔️", *tree.item(tree.get_children()[idx], 'values')[1:]), tags=("GREEN_BUTTON",))
        #print("Batch update request:", batch_update_values_request)  # Debug print

        # Execute batch update request
        if batch_update_values_request['data']:
            service.spreadsheets().values().batchUpdate(spreadsheetId=sheet_id, body=batch_update_values_request).execute()

    # Apply styles
    tree.tag_configure("GREEN_BUTTON", background="green")
    
    # Display notification when URL pasting finishes
    messagebox.showinfo("URL Pasting", "URL pasting process finished.")




def on_closing():
    """Close the application."""
    if messagebox.askokcancel("Quit", "Do you want to quit?"):
        root.destroy()

# Create the main window
root = tk.Tk()
root.title("Google Drive Search Tool")

# Style for the Treeview widget
style = ttk.Style()
style.theme_use("clam")
style.configure("Treeview", background="#D3D3D3", foreground="black", rowheight=25, fieldbackground="#D3D3D3")
style.configure("Treeview.Heading", background="blue", foreground="white")
style.configure("Green.TButton", background="green")

# Create folder selection widgets
folder_label = tk.Label(root, text="Google Drive Folder ID:")
folder_label.grid(row=0, column=0, sticky="e")

folder_entry = tk.Entry(root)
folder_entry.grid(row=0, column=1)

folder_button = tk.Button(root, text="Select Folder", command=select_folder)
folder_button.grid(row=0, column=2)

# Create sheet selection widgets
sheet_label = tk.Label(root, text="Google Sheet Link:")
sheet_label.grid(row=1, column=0, sticky="e")

sheet_entry = tk.Entry(root)
sheet_entry.grid(row=1, column=1)

sheet_button = tk.Button(root, text="Select Sheet", command=select_sheet)
sheet_button.grid(row=1, column=2)

# Create tab selection widgets
tab_label = tk.Label(root, text="Tab Name:")
tab_label.grid(row=2, column=0, sticky="e")

tab_entry = tk.Entry(root)
tab_entry.grid(row=2, column=1)

# Create column selection widgets
column_label = tk.Label(root, text="Column Name:")
column_label.grid(row=3, column=0, sticky="e")

column_entry = tk.Entry(root)
column_entry.grid(row=3, column=1)

phone_label = tk.Label(root, text="Phone Number Column:")
phone_label.grid(row=4, column=0, sticky="e")

phone_entry = tk.Entry(root)
phone_entry.grid(row=4, column=1)

# Create search and clear buttons
search_button = tk.Button(root, text="Search", command=start_search)
search_button.grid(row=5, column=0)

# Create Link URLs button
link_button = tk.Button(root, text="Link URLs", command=link)
link_button.grid(row=5, column=0, columnspan=3)

clear_button = tk.Button(root, text="Clear Results", command=clear_results)
clear_button.grid(row=5, column=2)

# Create Treeview widget to display search results
tree = ttk.Treeview(root, columns=('Check','Index', 'File Name', 'G-S Name', 'URL'), selectmode="browse")
tree.heading('#0', text='')
tree.heading('Check', text='Check')
tree.heading('Index', text='Index')
tree.heading('File Name', text='File Name')
tree.heading('G-S Name', text='G-S Name')
tree.heading('URL', text='URL')
tree.column('#0', width=50, anchor='center')
tree.column('Check', width=50, anchor='center')
tree.column('Index', width=50, anchor='center')
tree.column('File Name', width=200, anchor='center')
tree.column('G-S Name', width=200, anchor='center')
tree.column('URL', width=300, anchor='center')
tree.grid(row=6, column=0, columnspan=3)

# Function to paste URLs into Treeview
def paste_urls(urls):
    for idx, url_info in enumerate(urls):
        file_name, file_url = url_info
        item_id = tree.insert("", tk.END, values=("",idx+1, file_name, "", file_url))

        # Verify URL is pasted correctly in the URL column (column index 4)
        url_column_value = tree.set(item_id, column='URL')
        if url_column_value == file_url:
            # If URL is pasted correctly, add the checkmark in the first column
            tree.item(item_id, values=("✔️ Check", idx + 1, file_name, "", file_url), tags=("GREEN_BUTTON",))
        else:
            # If URL is not pasted correctly, remove the checkmark
            tree.item(item_id, values=("❌ Check", idx + 1, file_name, "", file_url), tags=("RED_BUTTON",))

# Insert values into Treeview
for idx, url_info in enumerate(urls):
    file_name, file_url = url_info
    tree.insert("", tk.END, values=("",idx+1, file_name,"", file_url))

def copy_url(event):
    """Copy the URL to the clipboard when a user double-clicks on an item in the URL column."""
    item = tree.selection()[0]  # Get the selected item
    url = tree.item(item, 'values')[3]  # Get the URL from the fourth column (index 3)
    root.clipboard_clear()  # Clear the clipboard
    root.clipboard_append(url)  # Append the URL to the clipboard

# Bind the double-click event to the Treeview to copy the URL
tree.bind("<Double-1>", copy_url)


# Create hit count label
hit_count_label = tk.Label(root, text="Total Hits: 0")
hit_count_label.grid(row=7, column=0, columnspan=3)

# Create Link URLs button
#link_button = tk.Button(root, text="Link URLs", command=link)
#link_button.grid(row=8, column=0, columnspan=3)

# Bind the closing event
root.protocol("WM_DELETE_WINDOW", on_closing)

# Start the application
root.mainloop()
