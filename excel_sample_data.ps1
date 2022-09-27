# use TLS1.2 with HTTPS:
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# creates folder if it does not yet exist:
filter Assert-FolderExists
{
  $exists = Test-Path -Path $_ -PathType Container
  if (!$exists) { 
    Write-Warning "$_ did not exist. Folder created."
    $null = New-Item -Path $_ -ItemType Directory 
  }
}

# download, unblock and extract zip files
filter Download-Zip($Path)
{
  # download to temp file:
  $temp = "$env:temp\temp.zip"
  Invoke-WebRequest -Uri $_ -OutFile $temp
  # unblock:
  Unblock-File -Path $temp
  # extract archive content:
  Expand-Archive -Path $temp -DestinationPath $Path -Force
  
  # report
  $zip = [System.IO.Compression.ZipFile]::OpenRead($temp)
  $zip.Entries | ForEach-Object { Write-Warning "Download: $_" }
  $zip.Dispose()
  
  # remove temp file:
  Remove-Item -Path $temp
}

# test whether filename is valid:
function Test-ValidFileName($FileName)
{
  $FileName.IndexOfAny([System.IO.Path]::GetInvalidFileNameChars()) -eq -1
}

# download and unblock file:
filter Download-File($Path, $FileName)
{
  # does the url specify a filename?
  if ([string]::IsNullOrWhiteSpace($FileName))
  {
    # take filename from url:
    $FileName = $_.Split('/')[-1]
    # remove url parameters:
    $FileName = $FileName.Split('?')[0]
    # test for valid file name:
    $isValid = Test-ValidFileName -FileName $FileName
    if (!$isValid)
    {
      throw "Url contains no valid file name. $FileName is not valid. Use parameter -FileName to specify a valid filename."
    }
  }
  
  $filePath = Join-Path -Path $Path -ChildPath $FileName
  Invoke-WebRequest -Uri $_ -OutFile $filePath
  # unblock:
  Unblock-File -Path $Path
  
  Write-Warning "Download: $FileName"
}

# create local folder for downloaded files:
($OutPath = "$env:temp\excelsampledata") | Assert-FolderExists

# download various excel sample files:
'https://www.contextures.com/SampleData.zip' | Download-Zip -Path $OutPath
'https://go.microsoft.com/fwlink/?LinkID=521962' | Download-File -Path $OutPath -FileName financial.xlsx
'http://www.principlesofeconometrics.com/excel/theories.xls' | Download-File -Path $OutPath 
'http://www.principlesofeconometrics.com/excel/food.xls' | Download-File -Path $OutPath 
'https://www.who.int/healthinfo/statistics/whostat2005_mortality.xls?ua=1' | Download-File -Path $OutPath 
'https://www.who.int/healthinfo/statistics/whostat2005_demographics.xls?ua=1' | Download-File -Path $OutPath 