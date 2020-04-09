#Script to backup local files at logoff

$Mdriveroot = "$env:HOMESHARE"
$localprofile = "$env:USERPROFILE"
$roamappdata = "$env:APPDATA"
$localappdata = "$env:LOCALAPPDATA"
$chrome = Test-Path "$Mdriveroot\Backup\Chrome"

#Copy-Item -Recurse "$localprofile\Desktop" "$Mdriveroot\Backup\Desktop" -ErrorAction Ignore
Copy-Item -Recurse "$localprofile\Favorites" "$Mdriveroot\Backup\Favorites" -Force -ErrorAction Ignore
Copy-Item -Recurse "$roamappdata\Mozilla\Firefox\Profiles" "$Mdriveroot\Backup\Firefox" -Force -ErrorAction Ignore
Copy-Item -Recurse "$roamappdata\Microsoft\Templates" "$Mdriveroot\Backup\Microsoft\Templates" -Force -ErrorAction Ignore
Copy-Item -Recurse "$roamappdata\Microsoft\Signatures" "$Mdriveroot\Backup\Microsoft\Signatures" -Force -ErrorAction Ignore
Copy-Item -Recurse "$roamappdata\Microsoft\Document Building Blocks" "$Mdriveroot\Backup\Microsoft\Document Building Blocks" -Force -ErrorAction Ignore
Copy-Item -Recurse "$localappdata\Microsoft\Office" "$Mdriveroot\Backup\Microsoft\Office" -Force -ErrorAction Ignore

if ($chrome -eq $true) {
    Copy-Item "$localappdata\Google\Chrome\User Data\Default\Bookmarks" "$Mdriveroot\Backup\Chrome" -Force -ErrorAction Ignore
} else {
    New-Item -ItemType Directory -Path "$Mdriveroot\Backup\Chrome"
    Copy-Item "$localappdata\Google\Chrome\User Data\Default\Bookmarks" "$Mdriveroot\Backup\Chrome" -Force -ErrorAction Ignore
}

#End Script