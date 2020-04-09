$Credential = Get-Credential -Message "Enter DOMAIN\username for the domain you would like to check against: DOMAIN1, DOMAIN2 or DOMAIN3"       # | select-object -expandproperty username
$Domain = $Credential.getNetworkCredential().domain

if ($Domain -eq "DOMAIN1") {
    $SearchBase = "OU=Servers,DC=DOMAIN1,DC=COM"
    $Server = "SERVER1.DOMAIN1.COM"
}
elseif ($Domain -eq "DOMAIN2") {
    $SearchBase = "OU=Servers,DC=DOMAIN2,DC=COM"
    $Server = "SERVER1.DOMAIN2.COM"
}
elseif ($Domain -eq "DOMAIN3") {
    $SearchBase = "OU=Servers,DC=DOMAIN3,DC=COM"
    $Server = "SERVER1.DOMAIN3.COM"
}
elseif ($Domain -ne "DOMAIN1", "DOMAIN2", "DOMAIN3") {
    Write-Host "
Please enter a correct domain." -ForegroundColor Red
    Stop
}


$OUList = Get-ADOrganizationalUnit -Credential $Credential -SearchBase $SearchBase -Server $Server -Filter * -Properties Name, DistinguishedName | Select-Object -Property Name, DistinguishedName
$OU = $OUList | Out-GridView -Title "Select OU and Click OK" -OutputMode Single | Select-Object -expandproperty distinguishedname

function Hotfixreport {  
    $computers = Get-ADComputer -Credential $Credential -Filter * -SearchBase $OU -Server $Server | Select-Object -ExpandProperty DNSHostName
    ForEach ($computer in $computers) {     
        $data = Get-HotFix -ComputerName $computer -Credential $Credential | Select-Object PSComputerName, InstalledOn | Sort-Object InstalledOn -Descending
        $data[0] | Format-Table -AutoSize -HideTableHeaders
    }
}

Hotfixreport

Write-Host "The search is finished.

" -ForegroundColor Green

Pause