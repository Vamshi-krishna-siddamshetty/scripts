Add-Type -AssemblyName System.IO.Compression.FileSystem
$CompressionLevel = [System.IO.Compression.CompressionLevel]::Optimal

$SourceFiles = Get-Childitem C:\*\*.bak
$ZipFile = 'C:\*\filename.zip'
$Zip = [System.IO.Compression.ZipFile]::Open($ZipFile,'Create')

ForEach ($SourceFile in $SourceFiles)
  {
    $SourcePath = $SourceFile.Fullname
    $SourceName = $SourceFile.Name
    $null = [System.IO.Compression.ZipFileExtensions]::
             CreateEntryFromFile($Zip,$SourcePath,$SourceName,1)
  }

$Zip.Dispose()
