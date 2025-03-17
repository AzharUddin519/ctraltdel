Get-SPOSite -IncludePersonalSite $true -Limit All | Where-Object { 
    $_.Url -like "*-my.sharepoint.com/personal/*" -and 
    $_.Owner -like "*tyrone-admin@chinooktx.com*" -and 
    $_.Url -match "ebjerkholt|aking|kschumacher"
}
