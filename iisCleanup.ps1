#ISS Clean up Script. Requires 7zip to be installed.
#import IIS WebAdministration Module - PS Script Must be run as admin
Import-Module WebAdministration


$7zipPath = "$env:ProgramFiles\7-Zip\7z.exe"
if (-not (test-path "$7zipPath")) {throw "$7zipPath needed"}

$PMonth = (Get-Date).Month-1
$CMonth = (Get-Date).Month
$Year = (Get-Date).Year
$PYear = (Get-Date).Year-1

$StartDate = "$PMonth/1/$year"
$EndDate = "$CMonth/1/$Year"

foreach($WebSite in $(get-website)){
	#Get IIS Log file Path
    $IISlogFilePath = "$($Website.logFile.directory)\W3SVC$($website.id)".replace("%SystemDrive%",$env:SystemDrive)
    $LogFileOldestLastWrite = (Get-ChildItem -Path $IISlogFilePath -Filter *.log ).LastWriteTime.Month | Measure-Object -Minimum
	#Remove Previous Years Log file
	remove-item -Path "$IISlogFilePath\W3SVC$($website.id)-$PMonth-$PYear.zip" | Out-null
    $MCount = $LogFileOldestLastWrite.Minimum
    do {
        $StartDate = "$MCount/1/$year"
        $EndDate = "$($MCount+1)/1/$Year"
                
        $ZipName = "W3SVC$($website.id)-$MCount-$Year.zip"
        $OutZipFile = "$IISlogFilePath\$ZipName"
        $ToZip = (Get-ChildItem -Path $IISlogFilePath -Filter *.log) | Where-Object {$_.LastWriteTime -ge $StartDate -and $_.LastWriteTime -le $EndDate}
        If ($ToZip.count -ne $Null) {
            Foreach($File in $ToZip){
                $FileToZip = $File.FullName
                & $7zipPath a $OutZipFile $FileToZip
            }
            Foreach($File in $ToZip){
                $FileName = $File.FullName
                Remove-Item -Path $FileName
            }
        }  
        $MCount++      
    } until ($MCount -eq $CMonth)
	
} 	