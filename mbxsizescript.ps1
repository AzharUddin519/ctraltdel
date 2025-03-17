<#
.SYNOPSIS
Generates a mailbox size report for specified users in Exchange Online.

.DESCRIPTION
This script connects to Exchange Online, retrieves mailbox statistics for specified users, and generates an Excel report with detailed size information.

.NOTES
- Requires ExchangeOnlineManagement and ImportExcel modules
- Input file should be CSV with 'UserInput' column containing email addresses or aliases
- Outputs Excel file with detailed mailbox size information

.PARAMETER InputFile
Path to CSV file containing user identifiers (email or alias)

.PARAMETER OutputFile
Path for output Excel file (default: MailboxSizeReport.xlsx)
#>

param(
    [Parameter(Mandatory=$true)]
    [string]$InputFile,
    [string]$OutputFile = "MailboxSizeReport.xlsx"
)

# Check required modules
$modules = @('ExchangeOnlineManagement', 'ImportExcel')
$missingModules = $modules | Where-Object { -not (Get-Module -ListAvailable $_) }

if ($missingModules) {
    Write-Host "The following modules are required: $($missingModules -join ', ')" -ForegroundColor Red
    $install = Read-Host "Would you like to install them now? (Y/N)"
    if ($install -eq 'Y') {
        foreach ($module in $missingModules) {
            Install-Module -Name $module -Scope CurrentUser -Force -SkipPublisherCheck
        }
    }
    else {
        Write-Host "Exiting script. Required modules not installed." -ForegroundColor Red
        exit
    }
}

# Import modules
Import-Module ExchangeOnlineManagement
Import-Module ImportExcel

# Connect to Exchange Online
try {
    Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
}
catch {
    Write-Error "Failed to connect to Exchange Online: $_"
    exit
}

# Process mailboxes
$users = Import-Csv -Path $InputFile
$report = foreach ($user in $users) {
    $inputIdentity = $user.UserInput
    try {
        # Get mailbox details
        $mailbox = Get-Mailbox -Identity $inputIdentity -ErrorAction Stop
        $stats = Get-MailboxStatistics -Identity $mailbox.UserPrincipalName

        # Parse size values
        $sizeProperties = @{
            TotalItemSize        = $stats.TotalItemSize
            TotalDeletedItemSize = $stats.TotalDeletedItemSize
            TotalDumpsterSize    = $stats.TotalDumpsterSize
        }

        $parsedSizes = $sizeProperties.GetEnumerator() | ForEach-Object {
            $result = @{
                "$($_.Key)"      = $_.Value
                "$($_.Key)Bytes" = $null
            }
            
            if ($_.Value -match '\(([\d,]+) bytes\)') {
                $bytes = $Matches[1] -replace ',',''
                $result["$($_.Key)Bytes"] = [long]$bytes
            }
            $result
        }

        # Prepare output object
        [PSCustomObject]@{
            Input                    = $inputIdentity
            PrimaryEmail            = $mailbox.PrimarySmtpAddress
            TotalItemSize           = $parsedSizes[0].TotalItemSize
            TotalItemSizeBytes      = $parsedSizes[0].TotalItemSizeBytes
            TotalDeletedItemSize    = $parsedSizes[1].TotalDeletedItemSize
            TotalDeletedItemSizeBytes = $parsedSizes[1].TotalDeletedItemSizeBytes
            DumpsterSize            = $parsedSizes[2].TotalDumpsterSize
            DumpsterSizeBytes       = $parsedSizes[2].TotalDumpsterSizeBytes
            RecoverableItemsQuota   = $mailbox.RecoverableItemsQuota
            ItemCount               = $stats.ItemCount
            DeletedItemCount        = $stats.DeletedItemCount
            LastLogonTime           = $stats.LastLogonTime
        }
    }
    catch {
        [PSCustomObject]@{
            Input                    = $inputIdentity
            PrimaryEmail            = "Error: $_"
            TotalItemSize           = $null
            TotalItemSizeBytes      = $null
            TotalDeletedItemSize    = $null
            TotalDeletedItemSizeBytes = $null
            DumpsterSize            = $null
            DumpsterSizeBytes       = $null
            RecoverableItemsQuota   = $null
            ItemCount               = $null
            DeletedItemCount        = $null
            LastLogonTime           = $null
        }
    }
}

# Generate report
try {
    $excelParams = @{
        Path          = $OutputFile
        WorksheetName = "Mailbox Sizes"
        AutoSize      = $true
        TableName     = "MailboxSizeReport"
        FreezeTopRow  = $true
        BoldTopRow    = $true
        Numberformat  = '###,###,##0'
    }

    $report | Export-Excel @excelParams
    Write-Host "Report generated successfully: $OutputFile" -ForegroundColor Green
}
catch {
    $csvPath = $OutputFile -replace '\.xlsx$', '.csv'
    $report | Export-Csv -Path $csvPath -NoTypeInformation
    Write-Host "Excel generation failed. CSV report created: $csvPath" -ForegroundColor Yellow
}

# Disconnect session
Disconnect-ExchangeOnline -Confirm:$false | Out-Null
