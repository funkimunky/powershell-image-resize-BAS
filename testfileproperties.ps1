Add-Type -AssemblyName System.Drawing

$ImageList = Get-ChildItem "C:\Temp\images"

$LandscapePath = "C:\Temp\Landscape"
$PortraitPath = "C:\Temp\Portrait"

foreach ($ImageFile in $ImageList)
{
    #Get the image information
    $image = New-Object System.Drawing.Bitmap $ImageFile.Fullname
    #Get the image attributes
    $ImageHeight = $image.Height
    $ImageWidth = $image.Width
    #Close the image
    $image.Dispose()
    If($ImageWidth -gt $ImageHeight)
    {
        Move-Item -LiteralPath $ImageFile.FullName -Destination $LandscapePath -Force
    }
    Else
    {
        Move-Item -LiteralPath $ImageFile.FullName -Destination $PortraitPath -Force
    }
}
