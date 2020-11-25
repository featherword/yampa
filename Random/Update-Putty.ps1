$downloadURL = (Invoke-WebRequest -Uri "https://www.chiark.greenend.org.uk/~sgtatham/putty/latest.html").Links.Href | where { $_ -like "https://the.earth.li/~sgtatham/putty/latest/w64/putty-64bit*.msi" };
$tempPath = "$env:TEMP";
$winDirPath = "$env:windir";

function updPUTTY {
    if((Test-Path -Path "C:\Program Files\PuTTY\putty.exe") -Or (Test-Path -Path "C:\Program Files (x86)\PuTTY\putty.exe")) {
            #For versions before .68 you must kill the task and delete putty.exe to silently uninstall.
            Get-Process -Name "putty.exe" | Stop-Process -EV Err -EA SilentlyContinue;
            Remove-Item -Force "C:\Program Files\PuTTY\putty.exe" -EV Err -EA SilentlyContinue;
            Remove-Item -Force "C:\Program Files (x86)\PuTTY\putty.exe" -EV Err -EA SilentlyContinue;
            & 'C:\Program Files\PuTTY\unins000.exe' /VERYSILENT /VERYSILENT /SUPPRESSMSGBOXES /NORESTART /SP- -EV Err -EA SilentlyContinue;
            & 'C:\Program Files (x86)\PuTTY\unins000.exe' /VERYSILENT /VERYSILENT /SUPPRESSMSGBOXES /NORESTART /SP- -EV Err -EA SilentlyContinue;

            #For newer versions installed by .msi package
            Get-Package *putty* | Uninstall-Package -EV Err -EA SilentlyContinue;
    };
    Try {
        #Downloads and installs the latest stable version of putty.
        Invoke-WebRequest $downloadURL -OutFile "$tempPath\PUTTY_INST.msi" -ErrorAction Stop | Out-Null;
        if (Test-Path -Path $tempPath\PUTTY_INST.msi) {
            $MSIArguments = (@("/i `"$tempPath\PUTTY_INST.msi`"","-L*V `"$tempPath\PuTTYLogFile.txt`"","/qn") | Where-Object {$_}) -join ' ';
            Start-Process -FilePath "$winDirPath\system32\msiexec.exe" -ArgumentList $MSIArguments -WorkingDirectory $tempPath -Wait;
            Remove-Item -Force "$tempPath\PUTTY_INST.msi" -EV Err -EA SilentlyContinue;
        } else {
            throw "Error in downloading the file, file non-existant in downloaded location.";
        }
    } Catch {
        $_.Exception | Out-File "$tempPath\puttyinstlog.log" -Append;
    }
    exit 0;
}
updPUTTY;