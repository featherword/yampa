$filePath = "<Path>\License_Keys.csv"
$csv = Import-Csv -Path $filePath

$locationSelection = ""
$productKeySelection = ""

function activateProduct($productKey) {
    cmd /c "cscript //b C:\Windows\System32\slmgr.vbs -ipk $productkey"
    cmd /c "cscript //b C:\Windows\System32\slmgr.vbs -ato"

    $Status = (Get-CimInstance -ClassName SoftwareLicensingProduct -Filter "Name like 'Windows%'" | where PartialProductKey).licensestatus
    If ($Status -ne 1) {
        throw "Windows unable to activate"
    } 
}

function confirmSelection {
    $confirmSelection = [System.Windows.MessageBox]::Show("You selected the Product Key for Location: $locationSelection Please speak with account manager to verify if unsure.`nWindows will activate with the selected MAK.","Confirm Selection",'YesNo','Information')            
    switch  ($confirmSelection) {
        'Yes' {
            activateProduct($productKeySelection)
        }
        'No' {
            exit 0;
        }
    }
}

function selectLocation {
    $selection = $csv | Out-GridView -PassThru -Title 'Select Product Key Location'
    Switch ($selection) {
        {$selection.index -ne 0} 
        {   
            $Script:locationSelection = $selection.Organization
            $Script:productKeySelection = $selection.'Product Key'
            confirmSelection
        }
    }
}


function main {
    try{
        $biosKey = (Get-WmiObject -Class SoftwareLicensingService).OA3xOriginalProductKey
        activateProduct($biosKey)
    } catch {
        [System.Windows.MessageBox]::Show('Windows was unable to activate automatically w/bios key. Please select the client key next window or cancel if client does not exist!');
        selectLocation
    }
}

main