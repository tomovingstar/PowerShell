#https://gist.github.com/jehugaleahsa/e23d90f65f378aff9aa254e774b40bc7
function join($path)
{
    $files = Get-ChildItem -Path "$path.*.part" | Sort-Object -Property @{Expression={
        $shortName = [System.IO.Path]::GetFileNameWithoutExtension($_.Name)
        $extension = [System.IO.Path]::GetExtension($shortName)
        if ($extension -ne $null -and $extension -ne '')
        {
            $extension = $extension.Substring(1)
        }
        [System.Convert]::ToInt32($extension)
    }}
    $writer = [System.IO.File]::OpenWrite($path)
    foreach ($file in $files)
    {
        $bytes = [System.IO.File]::ReadAllBytes($file)
        $writer.Write($bytes, 0, $bytes.Length)
    }
    $writer.Close()
}

#join "C:\path\to\file"

#function split($path, $chunkSize=107374182)
function split($path, $chunkSize=24000000)
{
    $fileName = [System.IO.Path]::GetFileNameWithoutExtension($path)
    $directory = [System.IO.Path]::GetDirectoryName($path)
    $extension = [System.IO.Path]::GetExtension($path)

    $file = New-Object System.IO.FileInfo($path)
    $totalChunks = [int]($file.Length / $chunkSize) + 1
    $digitCount = [int][System.Math]::Log10($totalChunks) + 1

    $reader = [System.IO.File]::OpenRead($path)
    $count = 0
    $buffer = New-Object Byte[] $chunkSize
    $hasMore = $true
    while($hasMore)
    {
        $bytesRead = $reader.Read($buffer, 0, $buffer.Length)
        $chunkFileName = "$directory\$fileName$extension.{0:D$digitCount}.part"
        $chunkFileName = $chunkFileName -f $count
        $output = $buffer
        if ($bytesRead -ne $buffer.Length)
        {
            $hasMore = $false
            $output = New-Object Byte[] $bytesRead
            [System.Array]::Copy($buffer, $output, $bytesRead)
        }
        [System.IO.File]::WriteAllBytes($chunkFileName, $output)
        ++$count
    }

    $reader.Close()
}

#split "C:\path\to\file"

$inDir = 'C:\DataLogsTran'
<#
#$zipDir = 'D:\AppData\NRZ_Extract\temp'
#$outDir = '\\wtpchr06320v1\ddc\DC_OUTPUT\EOM\NRZ\UAT'
#$outDir = '\\wtpchr06320v1\hr06320s03$\CTDIV\IR2\SHARED\Common\UNITS\TECHNOLOGY\SFSDOC_4307'
#$outDir = '\\wtpchr06320v1\ddc\DC_OUTPUT\EOM\NRZ\PROD'
$logFile = "$inDir\LogFile.txt"

$date = (Get-Date).ToString("yyyyMMdd")
#Remove-Item -Recurse -Path "$inDir\*.csv"
Get-ChildItem "$inDir\*" -Include *.csv -Recurse | Remove-Item
Get-ChildItem "$inDir\*" -Include *.zip -Recurse | Remove-Item
Compress-Archive -Path "$inDir\*" -DestinationPath "$inDir\$date.zip"
#Copy-Item "$zipDir\DOC_SFSDOC_4307_$date.zip" -Destination $outDir

#if(!$Error) {
#    Remove-Item -Path "$inDir\*.csv"
#}

"Date: $(Get-Date), Error if any: $Error" | Out-File $logFile -Append -Encoding ASCII
Write-Host "Data stored in $inDir"
#>

split "C:\DataLogsTran\20180809.zip"
