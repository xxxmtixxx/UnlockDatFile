# Check if the script is running as administrator
if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Warning "This script must be run as an administrator. Restarting with elevated privileges..."
    Start-Process -FilePath "powershell.exe" -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    exit
}

# Get the directory of the current script
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

# Import necessary assemblies for GUI
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Function to get open files using psfile
function Get-OpenFiles {
    try {
        # Run psfile to get the list of open files using the full path
        $psfilePath = Join-Path -Path $scriptDir -ChildPath "psfile.exe"
        $psfileOutput = & cmd /c "`"$psfilePath`" -nobanner" 2>&1
        if ($LASTEXITCODE -ne 0 -or [string]::IsNullOrWhiteSpace($psfileOutput)) {
            throw "Error running psfile command: $psfileOutput"
        }

        # Convert the output to an array of lines
        $psfileOutput = $psfileOutput.Split([Environment]::NewLine)

        # Parse the output to get file information
        $fileEntries = @()
        for ($i = 0; $i -lt $psfileOutput.Count; $i++) {
            $line = $psfileOutput[$i].Trim()

            # Check if the line contains the file ID and path
            if ($line -match '^\[(\d+)\]\s+(.+)$') {
                $fileId = $matches[1].Trim()
                $filePath = $matches[2].Trim()

                # Initialize placeholders for user, locks, and access
                $userName = "Unknown"
                $locks = "Unknown"
                $access = "Unknown"

                # Parse the next three lines for user, locks, and access
                for ($j = 1; $j -le 3; $j++) {
                    if (($i + $j) -lt $psfileOutput.Count) {
                        $nextLine = $psfileOutput[$i + $j].Trim()

                        if ($nextLine -match '^User:\s+(\S+)$') {
                            $userName = $matches[1].Trim()
                        } elseif ($nextLine -match '^Locks:\s+(\d+)$') {
                            $locks = $matches[1].Trim()
                        } elseif ($nextLine -match '^Access:\s+(Read|Write|Read Write)$') {
                            $access = $matches[1].Trim()
                        }
                    }
                }

                $fileEntries += [PSCustomObject]@{
                    FileId    = $fileId
                    UserName  = $userName
                    FilePath  = $filePath
                    Locks     = $locks
                    Access    = $access
                }

                # Skip the next three lines since they have been processed
                $i += 3
            }
        }

        return $fileEntries
    } catch {
        return @()
    }
}

function Close-OpenFiles {
    param (
        [Parameter(Mandatory=$true)]
        [System.Windows.Forms.ListView+SelectedListViewItemCollection]$selectedItems
    )

    try {
        if ($selectedItems.Count -eq 0) {
            throw "No files selected for closing."
        }

        # Show the wait form
        $waitForm.Show()
        $waitForm.Refresh()

        $psfilePath = Join-Path -Path $scriptDir -ChildPath "psfile.exe"
        $closedFiles = @()
        $failedFiles = @()

        foreach ($item in $selectedItems) {
            $fileId = $item.Text
            $filePath = $item.SubItems[3].Text  # Full path is in the fourth column

            try {
                # Update wait message
                $waitLabel.Text = "Closing file: $([System.IO.Path]::GetFileName($filePath))"
                $waitForm.Refresh()

                # Run psfile with the correct arguments to close the file
                $arguments = @("$fileId", "-c", "-nobanner")
                Start-Process -FilePath $psfilePath -ArgumentList $arguments -NoNewWindow -Wait

                # Add delay to give the system time to close the file
                Start-Sleep -Seconds 3

                # Verify if the file is still open
                $verifyOutput = & cmd /c "`"$psfilePath`" $fileId -nobanner" 2>&1

                # Check if the file is still open
                $fileStillOpen = $verifyOutput -match "\[$fileId\]"

                if ($fileStillOpen) {
                    # File is still open
                    $failedFiles += [System.IO.Path]::GetFileName($filePath)
                } else {
                    # File was successfully closed
                    $closedFiles += [System.IO.Path]::GetFileName($filePath)
                }
            } catch {
                $failedFiles += "$([System.IO.Path]::GetFileName($filePath)) : $($_.Exception.Message)"
            }
        }

        # Hide the wait form
        $waitForm.Hide()

        # Prepare the summary message
        $summaryMessage = "Operation Complete`n`n"
        
        if ($closedFiles.Count -gt 0) {
            $summaryMessage += "Successfully closed files:`n"
            $summaryMessage += $closedFiles -join "`n"
            $summaryMessage += "`n`n"
        }
        
        if ($failedFiles.Count -gt 0) {
            $summaryMessage += "Failed to close files:`n"
            $summaryMessage += $failedFiles -join "`n"
        }

        # Show the summary message box
        [System.Windows.Forms.MessageBox]::Show($summaryMessage, "File Closing Summary", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)

        # Update the list of open files after processing all selected items
        Update-OpenFiles
    } catch {
        # Hide the wait form in case of an error
        $waitForm.Hide()
        throw  # Re-throw the exception to be caught by the caller
    }
}

# Function to update the list of open files
function Update-OpenFiles {
    # Reset sorting before updating the list
    $listBox.ListViewItemSorter = $null
    $listBox.Tag = $null
    $listBox.Items.Clear()

    try {
        # Retrieve the latest list of open files
        $fileEntries = Get-OpenFiles

        if ($fileEntries.Count -eq 0) {
            throw "No open files found to display"
        }

        # Populate the list view with the new data
        foreach ($entry in $fileEntries) {
            $item = [System.Windows.Forms.ListViewItem]::new($entry.FileId.ToString())
            $item.SubItems.Add($entry.UserName)
            $item.SubItems.Add($entry.Access)
            $item.SubItems.Add($entry.FilePath)
            $listBox.Items.Add($item)
        }

        foreach ($column in $listBox.Columns) {
            $column.AutoResize([System.Windows.Forms.ColumnHeaderAutoResizeStyle]::ColumnContent)
        }
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error retrieving open files: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}

# Function to filter ListView items based on search text
function Filter-ListViewItems {
    param($searchText)

    # If the search text is empty, we've already updated the list in Perform-Search
    if ([string]::IsNullOrWhiteSpace($searchText)) {
        return
    }

    # Create a temporary collection for filtered items
    $filteredItems = @()

    foreach ($item in $listBox.Items) {
        $match = $false
        foreach ($subItem in $item.SubItems) {
            if ($subItem.Text -like "*$searchText*") {
                $match = $true
                break
            }
        }

        if ($match) {
            $filteredItems += $item
        }
    }

    # Clear and add filtered items back to the ListView
    $listBox.Items.Clear()
    $listBox.Items.AddRange($filteredItems)
}

# Form definition
$form = New-Object System.Windows.Forms.Form
$form.Text = "UnlockDatFile"
$form.Size = New-Object System.Drawing.Size(900, 600)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Sizable
$form.MaximizeBox = $true

# Add this at the beginning of your script, after the form definition
$waitForm = New-Object System.Windows.Forms.Form
$waitForm.Text = "Please Wait"
$waitForm.Size = New-Object System.Drawing.Size(300, 100)
$waitForm.StartPosition = "CenterScreen"
$waitForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
$waitForm.ControlBox = $false

$waitLabel = New-Object System.Windows.Forms.Label
$waitLabel.Location = New-Object System.Drawing.Point(10, 20)
$waitLabel.Size = New-Object System.Drawing.Size(280, 40)
$waitLabel.Text = "Closing files. Please wait..."
$waitLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$waitForm.Controls.Add($waitLabel)

# Set the form icon
$iconPath = Join-Path -Path $scriptDir -ChildPath "UnlockDatFile.ico"
if (Test-Path $iconPath) {
    $form.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($iconPath)
} else {
    Write-Warning "Icon file not found: $iconPath"
}

# Close button (moved to right)
$btnClose = New-Object System.Windows.Forms.Button
$btnClose.Text = "Close Selected File(s)"
$btnClose.Location = New-Object System.Drawing.Point(710, 20)
$btnClose.Size = New-Object System.Drawing.Size(150, 30)
$btnClose.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($btnClose)

# Search button and text box (moved to left)
$btnSearch = New-Object System.Windows.Forms.Button
$btnSearch.Text = "Search"
$btnSearch.Location = New-Object System.Drawing.Point(200, 20)
$btnSearch.Size = New-Object System.Drawing.Size(60, 23)
$btnSearch.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
$form.Controls.Add($btnSearch)

$txtSearch = New-Object System.Windows.Forms.TextBox
$txtSearch.Location = New-Object System.Drawing.Point(20, 20)
$txtSearch.Size = New-Object System.Drawing.Size(170, 23)
$txtSearch.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
$form.Controls.Add($txtSearch)

# List view for open files
$listBox = New-Object System.Windows.Forms.ListView
$listBox.View = 'Details'
$listBox.Location = New-Object System.Drawing.Point(20, 60)
$listBox.Size = New-Object System.Drawing.Size(840, 500)  # Increased height
$listBox.Columns.Add("ID", 50, [System.Windows.Forms.HorizontalAlignment]::Left)
$listBox.Columns.Add("User Name", 150, [System.Windows.Forms.HorizontalAlignment]::Left)
$listBox.Columns.Add("Access", 100, [System.Windows.Forms.HorizontalAlignment]::Left)
$listBox.Columns.Add("File Path", 500, [System.Windows.Forms.HorizontalAlignment]::Left)
$listBox.FullRowSelect = $true
$listBox.Sorting = [System.Windows.Forms.SortOrder]::None
$listBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($listBox)

# Function to perform search
function Perform-Search {
    # Always refresh the data first
    Update-OpenFiles

    # Then apply the search filter
    Filter-ListViewItems -searchText $txtSearch.Text
}

# Event handlers
$btnSearch.Add_Click({
    Perform-Search
})

$btnClose.Add_Click({
    if ($listBox.SelectedItems.Count -gt 0) {
        $dialogResult = [System.Windows.Forms.MessageBox]::Show("Are you sure you want to close the selected files?", "Confirm Close", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)

        if ($dialogResult -eq [System.Windows.Forms.DialogResult]::Yes) {
            try {
                Close-OpenFiles -selectedItems $listBox.SelectedItems
            } catch {
                [System.Windows.Forms.MessageBox]::Show("Error closing files: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        }
    } else {
        [System.Windows.Forms.MessageBox]::Show("Please select files to close.", "No Selection", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    }
})

$txtSearch.Add_KeyDown({
    param($sender, $e)
    if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
        Perform-Search
        $e.SuppressKeyPress = $true  # Prevents the ding sound
    }
})

# Custom sorter for ListView columns
class ListViewItemComparer : System.Collections.IComparer {
    [int]$Column
    [string]$Order

    ListViewItemComparer([int]$column, [string]$order) {
        $this.Column = $column
        $this.Order = $order
    }

    [int] Compare($x, $y) {
        $item1 = $x.SubItems[$this.Column].Text
        $item2 = $y.SubItems[$this.Column].Text

        if ($this.Column -eq 0) {
            # Sort the ID column numerically using Decimal to avoid Int32 and Int64 overflow issues
            [decimal]$val1 = [decimal]::Parse($item1)
            [decimal]$val2 = [decimal]::Parse($item2)
            $result = [decimal]::Compare($val1, $val2)
        } else {
            # Sort other columns as strings
            $result = [System.String]::Compare($item1, $item2)
        }

        if ($this.Order -eq "desc") {
            return -$result
        }
        return $result
    }
}

# Add column click event for sorting
$listBox.Add_ColumnClick({
    param($sender, $e)
    $column = $e.Column
    $currentTag = $listBox.Tag

    if ($currentTag -eq $null -or $currentTag -notmatch '^\d+:') {
        $currentOrder = "asc"
        $lastColumn = -1
    } else {
        $parts = $currentTag -split ':'
        $lastColumn = [int]$parts[0]
        $lastOrder = $parts[1]

        if ($lastColumn -eq $column) {
            $currentOrder = if ($lastOrder -eq "asc") { "desc" } else { "asc" }
        } else {
            $currentOrder = "asc"
        }
    }

    $listBox.Tag = ($column.ToString() + ":" + $currentOrder)
    $listBox.ListViewItemSorter = [ListViewItemComparer]::new($column, $currentOrder)
})

# Load initial data when form opens
$form.Add_Shown({
    Update-OpenFiles
})

# Run the form
[void]$form.ShowDialog()