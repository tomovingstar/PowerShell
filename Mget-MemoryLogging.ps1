#---------------------------------------------------------------------------
# This script will open IE, navigate to $navP page and stop for user input (number of reading attempts).
# It will then collect specific counters ($count) and store recorded values in $out file. Session Transcript is stored in MemReading.txt.
$CurrentPath = split-path $SCRIPT:MyInvocation.MyCommand.Path -parent
Start-Transcript "$CurrentPath\MemReading.txt"
function Output {
# To use, adjust the following settings:
$RefMS = 60 # Max Samples
$navP="about:blank"
$count="\Process(iexplore)\Private Bytes","\Process(iexplore#1)\Private Bytes" # Set counter
$RefSI = 1 #Sample Interval (s)
$out="$CurrentPath\MemReadingData.csv"
Write-Output "Script will create new IE window, clear IE chache and perform IE counter collection"
$ie = New-Object -ComObject InternetExplorer.Application
$ie.navigate($navP)
while($ie.ReadyState -ne 4) {start-sleep 1}
$ie.visible = $true
Write-Output "Current iexplore processes. Stop (Ctrl + C) if you see more than 2 iexplore instances"
Get-Process iexplore* | Select-Object id, name, mainwindowtitle | Format-Table -auto
$RefTime = $RefSI * $RefMS
Write-Output "Data collection will run for $RefTime sec"
[int]$RefNumb = Read-Host -Prompt 'Insert URL for page of interest in IE window; type integer for number of reading attempts in PS window, press "Enter" and mouse-click IE window afterward'
Write-Output "Clearing IE cache"
RunDll32.exe InetCpl.cpl, ClearMyTracksByProcess 8
RunDll32.exe InetCpl.cpl, ClearMyTracksByProcess 2
start-sleep 2
For($i=1; $i -le $RefNumb; $i++){
Write-Output "Reading attempt $i"
Write-Output "Data collection started"
$Res=Get-counter $count -SampleInterval $RefSI -MaxSamples $RefMS
Write-Output "Data collection finished for this attempt"
$ResTot = $ResTot + $Res }
$ResTot | Export-counter -Force -FileFormat CSV -Path $out
Get-Process iexplore* | Select-Object id, name, mainwindowtitle | Format-Table -auto
Write-Output "Data stored in file $out"
Read-Host -Prompt 'Press "Enter" to close IE and exit'
$ie.Quit() }
Output
Stop-Transcript
#---------------------------------------------------------------------------
