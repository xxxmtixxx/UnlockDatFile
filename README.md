# UnlockDatFile

UnlockDatFile is a PowerShell script that provides a graphical user interface for managing and closing open files on a Windows system. It's particularly useful for system administrators and power users who need to monitor and manage file locks.

## Features

- Lists all open files on the system
- Allows users to search for specific open files
- Enables closing of selected open files
- Provides a summary of successfully closed and failed-to-close files
- Supports sorting of the file list by different columns
- Scalable window that adjusts to different screen sizes
- Automatically loads data when the form opens
- Dynamic search functionality
- Verifies and ensures it's running with administrator privileges

## Requirements

- Windows operating system
- PowerShell
- Administrator privileges
- `psfile.exe` utility (should be in the same directory as the script)
  - You can download psfile from the [Microsoft Sysinternals PsFile page](https://learn.microsoft.com/en-us/sysinternals/downloads/psfile)

## Download and Setup

1. Download the script and its dependencies:
   - [Download UnlockDatFile ZIP](https://github.com/xxxmtixxx/UnlockDatFile/archive/refs/heads/main.zip)
2. Extract the ZIP file to a directory of your choice.
3. Download `psfile.exe` from the Microsoft Sysinternals link above and place it in the same directory as the extracted files.

## How to Use

1. Right-click on the `UnlockDatFile.ps1` script and select "Run with PowerShell".
   - The script will automatically check if it has administrator privileges and restart with elevated permissions if necessary.
2. The GUI will load, automatically displaying a list of all open files.
3. Use the search box to filter the list of files.
4. Select one or more files and click "Close Selected File(s)" to attempt to close them.
5. A summary message will show which files were successfully closed and which failed.

## GUI Components

- **Search Box**: Enter text to filter the list of open files.
- **Search Button**: Click to apply the search filter.
- **Close Selected File(s) Button**: Attempts to close the selected files.
- **File List**: Displays ID, User Name, Access type, and File Path for each open file.

## Features in Detail

1. **Admin Privilege Check**: The script verifies it's running with administrator privileges and automatically elevates if needed.
2. **File Listing**: Uses `psfile.exe` to retrieve a list of all open files on the system.
3. **Search Functionality**: 
   - If the search box is empty, all open files are displayed.
   - Entering text in the search box and pressing Enter or clicking Search refreshes the form before filtering and displaying matching files.
   - Filters based on content in any column (ID, User Name, Access, File Path).
4. **File Closing**: Attempts to close selected files and verifies if the closure was successful.
5. **Sortable Columns**: Click on column headers to sort the list by that column.
6. **Auto-refresh**: The file list is automatically updated when the application starts and after closing files.
7. **Scalable UI**: The window and its components automatically adjust to different screen sizes and resolutions.
8. **Initial Data Load**: Open files are automatically loaded and displayed when the form first opens.
9. **Error Handling**: Provides error messages for various scenarios, such as failure to retrieve open files or close selected files.
10. **Confirmation Dialog**: Asks for confirmation before attempting to close selected files.
11. **Keyboard Navigation**: Supports using the Enter key in the search box to initiate a search.

## Performance Considerations

- The script includes a 3-second delay after attempting to close a file to allow time for the system to process the closure.
- Large numbers of open files may impact the initial load time and search performance.

## Note

While the script checks for and requests administrator privileges, ensure you have the necessary permissions to run PowerShell scripts on your system.

## Disclaimer

Use this tool carefully, as closing open files may lead to data loss or application instability if not done properly. Always ensure you understand the implications of closing a file before proceeding.
