$Date = [DateTime]::Today.addmonths(-12)
#Run on the DC as ADMIN, should pull the correct domain component
$domainCom = Get-ADDomain -Current LocalComputer | Select-Object -Property DistinguishedName
if (!(Get-ADOrganizationalUnit -Filter "Name -eq 'Disabled Computers'")) {
    New-ADOrganizationalUnit -Name 'Disabled Computers' -Path $domainCom.DistinguishedName
}
$pcNames = Get-ADComputer -Filter 'LastLogonDate -lt $Date' -Properties LastLogonDate | Select-Object -Property Name,DistinguishedName

foreach ($name in $pcNames.Name){
    $result = Get-ADComputer -Filter 'Name -eq $name' -SearchBase "OU=Disabled Computers,$($domainCom.DistinguishedName)"
    if ($result -eq $null){
        Get-ADComputer $name | Move-ADObject -TargetPath "OU=Disabled Computers,$($domainCom.DistinguishedName)" 1>$null 2>$null
        Get-ADComputer $name | Disable-ADAccount 1>$null 2>$null
    }
}