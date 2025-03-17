# -------------------------------
# Mailbox Size & Dumpster Report
# -------------------------------

# Check if the ImportExcel module is installed; if not, install it.
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Output "ImportExcel module not found. Installing..."
    try {
        Install-Module -Name ImportExcel -Scope CurrentUser -Force -AllowClobber
    } catch {
        Write-Error "Failed to install ImportExcel module. Please install it manually."
        return
    }
}

# Define the path for the list of users and the output Excel report.
$userListPath = "C:\Azhar\Legal Team\chinook\Users.txt"        # <-- Update this path to your users list file (one email/alias per line)
$excelReportPath = "C:\Azhar\Legal Team\chinook\MailboxReport.xlsx"   # <-- Update to your desired report output path

# Read the list of users from file.
if (Test-Path $userListPath) {
    $users = Get-Content $userListPath
} else {
    Write-Error "User list file not found at $userListPath. Please create the file with one user per line."
    return
}

$results = @()

foreach ($user in $users) {
    Write-Output "Processing mailbox for user: $user"
    try {
        # Retrieve mailbox object (using alias or email provided).
        $mailbox = Get-Mailbox -Identity $user -ErrorAction Stop
        
        # Retrieve mailbox statistics.
        $stats = Get-MailboxStatistics -Identity $mailbox.Identity -ErrorAction Stop
        
        # Retrieve Recoverable Items (dumpster) statistics.
        $recoverableStats = Get-MailboxFolderStatistics -Identity $mailbox.Identity -FolderScope RecoverableItems -ErrorAction Stop
        
        # Calculate total dumpster size in bytes by summing the FolderSize from each recoverable folder.
        $totalDumpsterBytes = 0
        foreach ($folder in $recoverableStats) {
            if ($folder.FolderSize -and $folder.FolderSize.Value) {
                $totalDumpsterBytes += $folder.FolderSize.Value.ToBytes()
            }
        }
        # Convert total dumpster size to MB (rounded to 2 decimal places).
        $dumpsterSizeMB = [math]::Round($totalDumpsterBytes / 1MB, 2)
        
        # Sum up the number of items in all recoverable folders.
        $totalDumpsterItemCount = ($recoverableStats | Measure-Object -Property ItemsInFolder -Sum).Sum
        
        # Create an object with the required properties.
        $result = [PSCustomObject]@{
            Input                = $user
            Email                = $mailbox.PrimarySmtpAddress
            TotalItemSize        = $stats.TotalItemSize.ToString()
            TotalDeletedItemSize = $stats.TotalDeletedItemSize.ToString()
            DumpsterSizeMB       = "$dumpsterSizeMB MB"
            DumpsterItemCount    = $totalDumpsterItemCount
        }
    }
    catch {
        Write-Warning "Failed to retrieve data for user: $user. Error: $_"
        $result = [PSCustomObject]@{
            Input                = $user
            Email                = "N/A"
            TotalItemSize        = "N/A"
            TotalDeletedItemSize = "N/A"
            DumpsterSizeMB       = "N/A"
            DumpsterItemCount    = "N/A"
        }
    }
    $results += $result
}

# Export the results to an Excel file.
$results | Export-Excel -Path $excelReportPath -AutoSize -WorksheetName "Mailbox Report"

Write-Output "Mailbox report generated at: $excelReportPath"
