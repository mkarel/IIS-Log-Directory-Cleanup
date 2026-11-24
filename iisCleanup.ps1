#IIS Clean up Script. Requires Windows PowerShell 5.0 or later.
#import IIS WebAdministration Module - PS Script Must be run as admin
Import-Module WebAdministration


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
            if (Test-Path -LiteralPath $OutZipFile) {
                Compress-Archive -Path $ToZip.FullName -DestinationPath $OutZipFile -Update -ErrorAction Stop
            }
            else {
                Compress-Archive -Path $ToZip.FullName -DestinationPath $OutZipFile -ErrorAction Stop
            }
            $ToZip | Remove-Item -ErrorAction Stop
        }  
        $MCount++      
    } until ($MCount -eq $CMonth)
	
} 	