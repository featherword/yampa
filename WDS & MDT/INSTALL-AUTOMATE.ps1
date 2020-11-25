$tempPath = "$env:TEMP"
$currentUser = "$env:UserName"
$serverPass = "<SERVERPASSWORD>"
$serverAddress = "https://<SERVER>"
$filePath = "<PATH-TO-.CSV>\Automate-Location-IDs.csv"
$downloadURL = "https://<SERVER>/LabTech/Service/LabTechRemoteAgent.msi"
$csv = Import-Csv -Path $filePath

$clientSelection = ""
$locationSelection = ""
$locationIDSelection = ""


function runInstaller {
    Try {
        if (Test-Path -Path $tempPath\Agent_Install.msi) {
            $MSIArguments = @(
                        "/i"
                        "$tempPath\Agent_Install.msi"
                        "SERVERADDRESS=$serverAddress"
                        "SERVERPASS=$serverPass"
                        "LOCATION=$locationIDSelection"
                        "-L*V $tempPath\automateMSILogFile.txt"
            )
            Start-Process "msiexec.exe" -ArgumentList $MSIArguments -Wait -ErrorAction Stop
        } else {
            throw "File does not exist at runInstaller step."
        }
        Exit 0
    } Catch {
        $_.Exception | Out-File "$tempPath\automateinstlog.log" -Append
    }
}

function downloadInstaller {
    Try {
        Invoke-WebRequest $downloadURL -OutFile "$tempPath\Agent_Install.msi" -ErrorAction Stop | Out-Null
        Start-Sleep -Seconds 5
        if (Test-Path -Path $tempPath\Agent_Install.msi) {
            runInstaller
        } else {
            throw "Error in downloading the file, Agent_Install.msi does not exist at downloadInstaller step."
        }
    } Catch {
        $_.Exception | Out-File "$tempPath\automateinstlog.log" -Append
    }
}

function confirmSelection {
    $confirmSelection = [System.Windows.MessageBox]::Show("You selected the Automate agent for -`nClient: $clientSelection `nLocation: $locationSelection `nID: $locationIDSelection `n`nContinue? (No to return.)","Confirm Selection",'YesNo','Information')            
    switch  ($confirmSelection) {
        'Yes' {
            downloadInstaller
        }
        'No' {
            selectLocation
        }
    }
}

function selectLocation {
    $selection = $csv | Out-GridView -PassThru -Title 'Select Automate Agent to Install'
    Switch ($selection) {
        {$selection.index -ne 0} 
        {   
            $Script:clientSelection = $selection.client 
            $Script:locationSelection = $selection.location
            $Script:locationIDSelection = $selection.'location id'
            confirmSelection
        }
    }
}


selectLocation
