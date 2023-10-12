<#
Objective: 
Start each application when session is opened
Description: 
Start usefull application along with values
Date:
04/10/2018
#>

#Test script right on the computer by writing a file at the script location
#$text = 'test'
#$text | Out-File 'file.txt'

#Call script functions
Tools
Firefox
OpenExcelFiles
BackupFiles 'H:\SOC' 'U:\'

#Open IBM Lotus Notes, OneNotes, NotePad++, KeePass
Function Tools {   
    Invoke-Item "*.nsf"

    Invoke-Item "*.onetoc2"

    #Launch NotePad++
    *\Notepad++Portable.exe "*.py" 

    #Open KeePass
    *\KeePass.exe "*.kdbx"
    #to open the commun KeePass
    #H:\PortableApps\PortableApps\KeePass-2.34\KeePass.exe "S:\27-Telecom & Sécu\04-OCD -Sécurité (SECU)\Documentation reversibilité\BDD_MDP_I3T\Base_Commun_MDP_I3T.kdbx"
}

#Open usefull Firefox tabs when lauching the web browser
Function Firefox {    
    $FirefoxTab = @(
                    "https://google.com",
                  )
    #sleep=5s to open a new tab instead of opening a new windows
    For($j=0; $j -lt $FirefoxTab.Length; $j++)
    {
        *\FirefoxPortable.exe $FirefoxTab[$j]
        sleep(5)
    }
}

#Open Excel files
Function OpenExcelFiles {
    $filepath = @(
                    "*.xlsx"
                )
    For($i=0; $i -lt $filepath.Length; $i++)
    {
        if (Test-Path -Path $filepath[$i])
        {
            Invoke-Item $filepath[$i]
            sleep(1)
        }
        else{"File "+$filepath[$i]+" not Found"}
    }
}

#Backup folder from H: to U: every friday
#Two parameters, the source folder to zip and the destination with "\" at the end
Function BackupFiles {

    Param([string]$src, [string]$drive)

    $date = get-date -f ddd
    if ($date -match 'ven'){
    
        #syntax like "19900101"
        $dateNum = Get-Date -UFormat %Y%m%d 
    
        $dst = $drive+$dateNum+"_bckSOC"+".zip"
    
        #To zip files to the user network filesystem , used "system.io.compression" .NET class
        #To load the assembly, command "Add-Type"
        Add-Type -assembly "system.io.compression.filesystem"
        #ZipFile only has static method so "::", no "New-Object"
        [io.compression.zipfile]::CreateFromDirectory($src, $dst)
    }
}

