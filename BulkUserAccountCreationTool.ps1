Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$Form = New-Object system.Windows.Forms.Form
$Form.ClientSize = '638,657'
$Form.TopMost = $false

$PictureBox1 = New-Object system.Windows.Forms.PictureBox
$PictureBox1.width = 188
$PictureBox1.height = 111
$PictureBox1.location = New-Object System.Drawing.Point(486, 2)
$PictureBox1.imageLocation = "https://www.yourimagehere.com"
$PictureBox1.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::zoom
$Panel1 = New-Object system.Windows.Forms.Panel
$Panel1.height = 116
$Panel1.width = 641
$Panel1.BackColor = "#9b9b9b"
$Panel1.Anchor = 'right,left'
$Panel1.location = New-Object System.Drawing.Point(-3, 0)

$Label1 = New-Object system.Windows.Forms.Label
$Label1.text = "New User Account Creation Tool"
$Label1.AutoSize = $true
$Label1.width = 25
$Label1.height = 10
$Label1.location = New-Object System.Drawing.Point(29, 49)
$Label1.Font = 'Microsoft Sans Serif,16,style=Underline'

$Label2 = New-Object system.Windows.Forms.Label
$Label2.text = "SCOIT"
$Label2.AutoSize = $true
$Label2.width = 25
$Label2.height = 10
$Label2.location = New-Object System.Drawing.Point(11, 9)
$Label2.Font = 'MingLiU-ExtB,10,style=Bold'
$Label2.ForeColor = "#4a4a4a"

$Label3 = New-Object system.Windows.Forms.Label
$Label3.text = "Select mailbox database:"
$Label3.AutoSize = $true
$Label3.width = 25
$Label3.height = 10
$Label3.location = New-Object System.Drawing.Point(33, 133)
$Label3.Font = 'Microsoft Sans Serif,10'

$ComboBox1 = New-Object system.Windows.Forms.ComboBox
$ComboBox1.width = 100
$ComboBox1.height = 20
$ComboBox1.location = New-Object System.Drawing.Point(33, 156)
$ComboBox1.Font = 'Microsoft Sans Serif,10'

$ComboBox2 = New-Object system.Windows.Forms.ComboBox
$ComboBox2.width = 153
$ComboBox2.height = 20
$ComboBox2.location = New-Object System.Drawing.Point(33, 208)
$ComboBox2.Font = 'Microsoft Sans Serif,10'

$Label4 = New-Object system.Windows.Forms.Label
$Label4.text = "Select OU:"
$Label4.AutoSize = $true
$Label4.width = 25
$Label4.height = 10
$Label4.location = New-Object System.Drawing.Point(33, 185)
$Label4.Font = 'Microsoft Sans Serif,10'

$Label5 = New-Object system.Windows.Forms.Label
$Label5.text = "Enter password for user(s):"
$Label5.AutoSize = $true
$Label5.width = 25
$Label5.height = 10
$Label5.location = New-Object System.Drawing.Point(33, 241)
$Label5.Font = 'Microsoft Sans Serif,10'

$Label7 = New-Object system.Windows.Forms.Label
$Label7.text = "Re-enter password:"
$Label7.AutoSize = $true
$Label7.width = 25
$Label7.height = 10
$Label7.location = New-Object System.Drawing.Point(33, 294)
$Label7.Font = 'Microsoft Sans Serif,10'

$MaskedTextBox2 = New-Object system.Windows.Forms.MaskedTextBox
$MaskedTextBox2.multiline = $false
$MaskedTextBox2.width = 120
$MaskedTextBox2.height = 20
$MaskedTextBox2.location = New-Object System.Drawing.Point(33, 319)
$MaskedTextBox2.Font = 'Microsoft Sans Serif,10'
$MaskedTextBox2.PasswordChar = '*'

$Button1 = New-Object system.Windows.Forms.Button
$Button1.BackColor = "#9b9b9b"
$Button1.text = "Click to Import CSV"
$Button1.width = 134
$Button1.height = 41
$Button1.enabled = $false
$Button1.location = New-Object System.Drawing.Point(364, 133)
$Button1.Font = 'Microsoft Sans Serif,10,style=Bold'

$Button2 = New-Object system.Windows.Forms.Button
$Button2.BackColor = "#9b9b9b"
$Button2.text = "Click to Create Account(s)"
$Button2.width = 186
$Button2.height = 41
$Button2.enabled = $false
$Button2.location = New-Object System.Drawing.Point(364, 199)
$Button2.Font = 'Microsoft Sans Serif,10,style=Bold'

$MaskedTextBox1 = New-Object system.Windows.Forms.MaskedTextBox
$MaskedTextBox1.multiline = $false
$MaskedTextBox1.width = 119
$MaskedTextBox1.height = 20
$MaskedTextBox1.location = New-Object System.Drawing.Point(33, 266)
$MaskedTextBox1.Font = 'Microsoft Sans Serif,10'
$MaskedTextBox1.PasswordChar = '*'

$Label6 = New-Object system.Windows.Forms.Label
$Label6.text = "             "
$Label6.AutoSize = $false
$Label6.width = 255
$Label6.height = 63
$Label6.location = New-Object System.Drawing.Point(209, 251)
$Label6.Font = 'Microsoft Sans Serif,9'
$Label6.ForeColor = "#d0021b"

$Panel2 = New-Object system.Windows.Forms.Panel
$Panel2.height = 5
$Panel2.width = 635
$Panel2.BackColor = "#4a4a4a"
$Panel2.location = New-Object System.Drawing.Point(2, 358)

$Label8 = New-Object system.Windows.Forms.Label
$Label8.text = "Post account creation tasks:"
$Label8.AutoSize = $true
$Label8.width = 25
$Label8.height = 10
$Label8.location = New-Object System.Drawing.Point(188, 375)
$Label8.Font = 'Microsoft Sans Serif,14,style=Bold'

$Label9 = New-Object system.Windows.Forms.Label
$Label9.text = "Search for account to configure:"
$Label9.AutoSize = $true
$Label9.width = 25
$Label9.height = 10
$Label9.location = New-Object System.Drawing.Point(207, 414)
$Label9.Font = 'Microsoft Sans Serif,10'

$Label10 = New-Object system.Windows.Forms.Label
$Label10.text = "Version 2.0"
$Label10.AutoSize = $true
$Label10.width = 25
$Label10.height = 10
$Label10.location = New-Object System.Drawing.Point(583, 638)
$Label10.Font = 'Microsoft Sans Serif,6'

$Label11 = New-Object system.Windows.Forms.Label
$Label11.text = "Move user to OU:"
$Label11.AutoSize = $true
$Label11.width = 25
$Label11.height = 10
$Label11.location = New-Object System.Drawing.Point(33, 507)
$Label11.Font = 'Microsoft Sans Serif,10'

$ComboBox3 = New-Object system.Windows.Forms.ComboBox
$ComboBox3.width = 153
$ComboBox3.height = 20
$ComboBox3.location = New-Object System.Drawing.Point(33, 528)
$ComboBox3.Font = 'Microsoft Sans Serif,10'

$Button4 = New-Object system.Windows.Forms.Button
$Button4.BackColor = "#9b9b9b"
$Button4.text = "OK"
$Button4.width = 45
$Button4.height = 27
$Button4.enabled = $false
$Button4.location = New-Object System.Drawing.Point(189, 527)
$Button4.Font = 'Microsoft Sans Serif,10,style=Bold'

$Button5 = New-Object system.Windows.Forms.Button
$Button5.BackColor = "#9b9b9b"
$Button5.text = "Exit"
$Button5.width = 89
$Button5.height = 27
$Button5.location = New-Object System.Drawing.Point(267, 624)
$Button5.Font = 'Microsoft Sans Serif,10,style=Bold'

$Label12 = New-Object system.Windows.Forms.Label
$Label12.AutoSize = $true
$Label12.width = 25
$Label12.height = 10
$Label12.location = New-Object System.Drawing.Point(207, 326)
$Label12.Font = 'Microsoft Sans Serif,9,style=Bold'
$Label12.ForeColor = "#7ed321"

$ComboBox4 = New-Object system.Windows.Forms.ComboBox
$ComboBox4.text = ""
$ComboBox4.width = 153
$ComboBox4.height = 20
$ComboBox4.location = New-Object System.Drawing.Point(223, 443)
$ComboBox4.Font = 'Microsoft Sans Serif,10'

$Label13 = New-Object system.Windows.Forms.Label
$Label13.AutoSize = $true
$Label13.width = 25
$Label13.height = 10
$Label13.location = New-Object System.Drawing.Point(39, 563)
$Label13.Font = 'Microsoft Sans Serif,10'
$Label13.ForeColor = "#7ed321"

$Button3 = New-Object system.Windows.Forms.Button
$Button3.BackColor = "#9b9b9b"
$Button3.text = "Configure AD"
$Button3.width = 104
$Button3.height = 44
$Button3.enabled = $true
$Button3.location = New-Object System.Drawing.Point(311, 510)
$Button3.Font = 'Microsoft Sans Serif,10,style=Bold'

$Button6 = New-Object system.Windows.Forms.Button
$Button6.BackColor = "#9b9b9b"
$Button6.text = "Configure DFS"
$Button6.width = 104
$Button6.height = 44
$Button6.enabled = $true
$Button6.location = New-Object System.Drawing.Point(478, 510)
$Button6.Font = 'Microsoft Sans Serif,10,style=Bold'

$Panel1.controls.AddRange(@($PictureBox1, $Label1, $Label2))
$Form.controls.AddRange(@($Panel1, $Label3, $ComboBox1, $ComboBox2, $Label4, $Label5, $Label7, $MaskedTextBox2, $Button1, $Button2, $MaskedTextBox1, $Label6, $Panel2, $Label8, $Label9, $Label10, $Label11, $ComboBox3, $Button4, $Button5, $Label12, $ComboBox4, $Label13, $Button3, $Button6))

#$MaskedTextBox1.PasswordChar     = '*'
#$MaskedTextBox2.PasswordChar     = '*'
#$TextBox1.AutoCompleteSource     = 'CustomSource'
#$TextBox1.AutoCompleteMode       = 'SuggestAppend'

$Button1.Add_MouseClick( { ImportCSV })

$Button2.Add_MouseClick( { CreateAccounts })

$ComboBox4.Add_MouseClick( { searchAD })

$Button4.Add_MouseClick( { Move-Account })

$MaskedTextBox2.Add_KeyUp( { PasswordCheck })

$MaskedTextBox1.Add_KeyUp( { PasswordCheck })

$ComboBox4.Add_KeyUp( { EnableOK })

$ComboBox3.add_dropdownclosed( { EnableOK })

$ComboBox3.Add_KeyUp( { EnableOK })

$Button5.Add_MouseClick( { Exit-Window })

$Button3.Add_MouseClick( { Open-AD })

$Button6.Add_MouseClick( { Open-DFS })

#$ErrorActionPreference = "SilentlyContinue"

Function PasswordCheck {
    $Button1.Enabled = $false
    if ($MaskedTextBox1.Text -ne "" -and $MaskedTextBox2.Text -ne "") {
        $Password1 = $MaskedTextBox1.Text | ConvertTo-SecureString -AsPlainText -Force
        $Password2 = $MaskedTextBox2.Text | ConvertTo-SecureString -AsPlainText -Force
        $pwd1_text = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($Password1))
        $pwd2_text = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($Password2))
        if ($pwd1_text -ne $pwd2_text) { 
            $Label6.Text = "Passwords don't match, please re-enter passwords."
            Return
        }
        if ($pwd1_text -eq $pwd2_text) { 
            $Label6.Text = ""
        }
    
        $numComplexityTestResultsPassed = @(
            $pwd1_text -cmatch '[A-Z]'
            $pwd1_text -cmatch '[a-z]'
            $pwd1_text -cmatch '[0-9]'
            $pwd1_text -cmatch '[^a-zA-Z0-9]'
            $pwd1_text -match '\d'
        ).Where{ $_ }.Count

        if ($pwd1_text.Length -lt 9 -or $numComplexityTestResultsPassed -lt 5) {
            $Label6.Text = "Password complexity not met."
            Return
        }
        $Label6.Text = ""
        $Button1.Enabled = $true
    }
}

Function CreateAccounts {
    $Label12.Text = ""
    try {
        $Password1 = $MaskedTextBox1.Text | ConvertTo-SecureString -AsPlainText -Force
        $Databasename = $ComboBox1.Text
        #$OU = $ComboBox2.Value
        $OU = $ComboBox2.SelectedItem.DistinguishedName
        Import-CSV $Button1.Text |
        #Import-CSV ($FileBrowser).FileName |
        ForEach-Object { New-Mailbox -Name $_.name -firstname $_.firstname -lastname $_.lastname -userPrincipalName $_.userprincipalname -OrganizationalUnit $OU -Database $Databasename -DisplayName $_.name -Password $Password1 }
        Remove-PSSession $Session
        $Label12.Text = "Account creation(s) complete!" 
        $Button1.Enabled = $false
        $Button2.Enabled = $false
        $ComboBox1.Text = ""
        $ComboBox2.Text = ""
        $MaskedTextBox1.Text = ""
        $MaskedTextBox2.Text = ""
    }
    catch {
        $Label12.Text = "Account creation(s) Failed!" 
        $Label12.Text += $_
        Return
    }
}

Function ImportCSV {
    $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{
        Title            = "Select CSV containing user accounts"
        InitialDirectory = [Environment]::GetFolderPath('Desktop')
        Filter           = 'Comma-separated Values (*.csv)|*.csv'
    }
    $Result = $FileBrowser.ShowDialog()
    $Result
    $Result2 = $FileBrowser.FileName
    $correctHeaders = @(
        'Name', 'FirstName', 'LastName', 'UserPrincipalName'
    )
	   
    if ($Result -eq "OK") {  
            
        # put all the headers into a comma separated array
        $headers = (Get-Content $Result2 -TotalCount 1).Split(",")
        for ($i = 0; $i -lt $headers.Count; $i++) {
            # trim any leading white space and compare the headers
            if ($headers[$i].TrimStart() -ne $correctHeaders[$i]) {
                $Label6.Text = "$Result2 failed to validate headers." #because header number $i showed:$($headers[$i].TrimStart()) instead of:$($correctHeaders[$i]). Please fix the headers and try again."
                return
            }
            else {
                continue
            }
        }
        #Write-Host "Selected CSV File: $Result2"  -ForegroundColor Green  
        #Write-Host "CSV Successfully Imported." -ForegroundColor Green 
        $Button1.Text = $FileBrowser.FileName
        $Label6.Text = ""
        $Label12.Text = "$Result2 successfully imported."
        $Button2.Enabled = $true
    } 
    else { $Label6.Text = "CSV import canceled." }
}

function EnableOK {
    if ($ComboBox3.text -ne "" -and $ComboBox4 -ne "") {
        $Button4.enabled = $true
    }
    else {
        $Button4.enabled = $false 
    }
}

function searchAD {
    $ComboBox4.Items.Clear()
    #$ComboBox4.ResetText()
    $Userlist = Get-ADUser -SearchBase "OU=Accounts,DC=domain,DC=com" -Filter * | Sort-Object
    Foreach ($User in $Userlist) {
        [void] $ComboBox4.Items.Add($User);
        $ComboBox4.ValueMember = "DistinguishedName"
        $ComboBox4.DisplayMember = "Name"
    }
}

function Move-Account {
    $Identity = $ComboBox4.SelectedItem.distinguishedname
    $TargetPath = $ComboBox3.SelectedItem.DistinguishedName
    Move-ADObject -Identity $Identity -TargetPath $TargetPath
    $Label13.Text = "Account successfully moved."
}

function Exit-Window {
    $Form.Close()
}

function Open-AD {
    dsa.msc "/RDN=OU=Accounts"
}

function Open-DFS {
    Start-Process dfsmgmt.msc
}

#Create remote powershell session to Exchange server
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://exchange -Authentication Kerberos
Import-PSSession $Session -DisableNameChecking

$dbases = Get-MailboxDatabase | Select-Object -Property Name
Foreach ($dbase in $dbases) {
    [void] $ComboBox1.Items.Add($dbase.Name);
}

$OUList = Get-ADOrganizationalUnit -SearchBase "OU=Accounts,DC=domain,DC=com" -Filter * | Sort-Object Name # -Properties Name,DistinguishedName | Select-Object -Property Name,DistinguishedName | Sort-Object Name
Foreach ($orgunit in $OUList) {
    [void] $ComboBox2.Items.Add($orgunit);
    $ComboBox2.ValueMember = "DistinguishedName"
    $ComboBox2.DisplayMember = "Name"
    [void] $ComboBox3.Items.Add($orgunit);
    $ComboBox3.ValueMember = "DistinguishedName"
    $ComboBox3.DisplayMember = "Name"
}

[void] $Form.ShowDialog()