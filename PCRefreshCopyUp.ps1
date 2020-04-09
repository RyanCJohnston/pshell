#Script to copy local files and folders to user's shared drive(M:) and provide a few other outputs

$Mdriveroot = $null
$Mdriveroot = "$env:HOMESHARE"
$localprofile = $null
$localprofile = "$env:USERPROFILE"
$roamappdata = $null
$roamappdata = "$env:APPDATA"
$localappdata = $null
$localappdata = "$env:LOCALAPPDATA"

Start-Transcript -path "$Mdriveroot\PC Refresh\UploadLog.txt" -append

Copy-Item -Recurse "$localprofile\Videos" "$Mdriveroot\PC Refresh\My Videos" -Verbose
Copy-Item -Recurse "$localprofile\Documents" "$Mdriveroot\PC Refresh\Documents" -ErrorAction SilentlyContinue -Verbose
Copy-Item -Recurse "$localprofile\Music" "$Mdriveroot\PC Refresh\Music" -Verbose
Copy-Item -Recurse "$localprofile\Pictures" "$Mdriveroot\PC Refresh\Pictures" -Verbose
Copy-Item -Recurse "$localprofile\Desktop" "$Mdriveroot\PC Refresh\Desktop" -Verbose
Copy-Item -Recurse "$localprofile\Favorites" "$Mdriveroot\PC Refresh\Favorites" -Verbose
Copy-Item -Recurse "$roamappdata\Mozilla\Firefox\Profiles\" "$Mdriveroot\PC Refresh\Firefox" -ErrorAction SilentlyContinue -Verbose
Copy-Item -Recurse "$roamappdata\Microsoft\Templates" "$Mdriveroot\PC Refresh\Templates" -Verbose
Copy-Item -Recurse "$roamappdata\Microsoft\Signatures" "$Mdriveroot\PC Refresh\Signatures" -Verbose
Copy-Item -Recurse "$roamappdata\Microsoft\Document Building Blocks\1033\15" "$Mdriveroot\PC Refresh\Document Building Blocks" -Verbose
Copy-Item -Recurse "$localappdata\Microsoft\Outlook\RoamCache" "$Mdriveroot\PC Refresh\RoamCache" -Verbose
Copy-Item -path "$localappdata\Microsoft\Office\" -Recurse -Filter "*.officeUI" -Destination "$Mdriveroot\PC Refresh\OfficeUI" -Verbose
Copy-Item -Recurse "C:\Windows\Fonts" "$Mdriveroot\PC Refresh\Fonts" -Verbose
Get-ChildItem HKCU:\Software\Microsoft\Office\15.0\Outlook\Profiles\* | Out-File "$Mdriveroot\PC Refresh\Profiles.txt" -Verbose
Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Select-Object DisplayName, Publisher | Format-Table –AutoSize > "$Mdriveroot\PC Refresh\InstalledSoftware32bit.txt" -Verbose
Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Select-Object DisplayName, Publisher | Format-Table –AutoSize > "$Mdriveroot\PC Refresh\InstalledSoftware64bit.txt" -Verbose
Get-ChildItem -Path C:\ -Recurse -Include *.pst -ErrorAction SilentlyContinue | Select-Object fullname | Out-File "$Mdriveroot\PC Refresh\PST.txt" -Force -Verbose
New-Item -ItemType Directory -Path "$Mdriveroot\PC Refresh\Chrome" -Verbose
Copy-Item "$localappdata\Google\Chrome\User Data\Default\Bookmarks" "$Mdriveroot\PC Refresh\Chrome\Bookmarks" -Verbose

Stop-Transcript

 Write-host "

PCRefreshCopyUp script has finished, log saved at $Mdriveroot\PC Refresh\UploadLog.txt
            " -ForegroundColor Green

 Pause

#End Script

<#################################################################################
$sourcepath = "C:\"
$destination = "$Mdriveroot\PC Refresh\"
$sourcefiles = get-childitem $sourcepath -Include *.pst -Recurse -ErrorAction Ignore

foreach ($pst in $sourcefiles)
{
Copy-Item $pst.FullName -Destination "$destination\$($pst.Name)"
}
##################################################################################>