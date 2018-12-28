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

#Open IBM Lotus Notes
Invoke-Item "*.nsf"

#Open OneNotes
Invoke-Item "*.onetoc2"


#Open KeePass
*\KeePass.exe ".kdbx"

#Open usefull Firefox tabs when lauching the web browser
$FirefoxTab = @(
                "https://www.google.com",
                )
#sleep=4s to open a new tab instead of opening a new windows
For($j=0; $j -lt $FirefoxTab.Length; $j++)
{
    *\FirefoxPortable.exe $FirefoxTab[$j]
    sleep(5)
}

#Open Excel files
$filepath = @(
                #"*.xlsx",
            )
For($i=0; $i -lt $filepath.Length; $i++)
{
    Invoke-Item $filepath[$i]
    sleep(1)
}

#Backup folder from H: to U: every friday
$date = get-date -f ddd
if ($date -match 'ven')
{
    #syntax like "19900101"
    $dateNum = Get-Date -UFormat %Y%m%d 
    
    $src = "WhatToSave"
    $dst = "WhereToSave"+$dateNum+"_bck"+".zip"
    
    #To zip file format
    Add-Type -assembly "system.io.compression.filesystem"
    [io.compression.zipfile]::CreateFromDirectory($src, $dst)
}
