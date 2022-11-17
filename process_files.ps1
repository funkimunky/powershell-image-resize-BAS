Add-Type -AssemblyName System.Drawing
Import-Module -Name $PSScriptRoot/modules/ImportExcel -force

$OrigionalTotal = 0
$FinalTotal = 0

function Get-Size-Item-mb
{
    param([string]$pth)
    $size = "{0:n2}" -f ((Get-Item -path $pth | measure-object -property length -sum).sum /1mb)
    Return [float]$size
}

function Get-Size-Item-Kb
{
    param([string]$pth)
    $size = "{0:n2}" -f ((Get-Item -path $pth | measure-object -property length -sum).sum /1kb) + " kb"
    Return $size
}


Function Resize-Image() {    
    [CmdLetBinding(
        SupportsShouldProcess = $true, 
        PositionalBinding = $false,
        ConfirmImpact = "Low",
        DefaultParameterSetName = "Absolute"
    )]
    Param 
    (
        [Parameter(Mandatory = $True)]
        [ValidateScript({$_ | ForEach-Object { Test-Path $_ } })][String[]]$ImagePath,
        [Parameter(Mandatory = $False)][Switch]$MaintainRatio,
        [Parameter(Mandatory = $False, ParameterSetName = "Longerside")][Int]$Longerside,
        [Parameter(Mandatory = $False, ParameterSetName = "Absolute")][Int]$Height,
        [Parameter(Mandatory = $False, ParameterSetName = "Absolute")][Int]$Width,        
        [Parameter(Mandatory = $False, ParameterSetName = "Percent")][Double]$Percentage,
        [Parameter(Mandatory = $False)][System.Drawing.Drawing2D.SmoothingMode]$SmoothingMode = "HighQuality",
        [Parameter(Mandatory = $False)][System.Drawing.Drawing2D.InterpolationMode]$InterpolationMode = "HighQualityBicubic",
        [Parameter(Mandatory = $False)][System.Drawing.Drawing2D.PixelOffsetMode]$PixelOffsetMode = "HighQuality",
        [Parameter(Mandatory = $False)][String]$NameModifier = "resized",
        [Parameter(Mandatory = $False)][System.Management.Automation.SwitchParameter]$OverWrite
    )
    
    Begin 
    {
        $Global:OrigionalTotal

        If ($Width -and $Height -and $MaintainRatio) {
            Throw "Absolute Width and Height cannot be given with the MaintainRatio parameter."
        }
 
        If (($Width -xor $Height) -and (-not $MaintainRatio)) {
            Throw "MaintainRatio must be set with incomplete size parameters (Missing height or width without MaintainRatio)"
        }
 
        If ($Percentage -and $MaintainRatio) {
            Write-Warning "The MaintainRatio flag while using the Percentage parameter does nothing"
        }

        If ($Longerside -and $Width -or $Longerside -and $height) {
            Throw "Should only be longer side in pixels"
        }

        If ($Percentage -and $Longerside -or $MaintainRatio -and $Longerside) {
            Throw "Percentage or maintain ratio cannot be used with longerside pixels flag"
        }

        
    }
    Process 
    {
        
        try 
        {
            ForEach ($Image in $ImagePath) {                
                $Path = (Resolve-Path $Image).Path
                $Dot = $Path.LastIndexOf(".")
                
    
                switch ($OverWrite.IsPresent) {
                    $true {
                            # Overite images
                            $OutputPath = $Path.Substring(0, $Dot) + $Path.Substring($Dot, $Path.Length - $Dot)
                        }
                    $false {
                            # rename images
                            $OutputPath = $Path.Substring(0, $Dot) + "_" + $NameModifier + $Path.Substring($Dot, $Path.Length - $Dot)
                        }
                                      
                }

                $Global:OrigionalTotal += Get-Size-Item-mb($Image)

                $OldImage = New-Object -TypeName System.Drawing.Bitmap -ArgumentList $Path
                # Grab these for use in calculations below. 
                $OldHeight = $OldImage.Height
                $OldWidth = $OldImage.Width

                If ($MaintainRatio) {                
                    If ($Height) {
                        $Width = $OldWidth / $OldHeight * $Height
                    }
                    If ($Width) {
                        $Height = $OldHeight / $OldWidth * $Width
                    }
                }

                If ($Percentage) {
                    $Product = ($Percentage / 100)
                    $Height = $OldHeight * $Product
                    $Width = $OldWidth * $Product
                }

                If ($Longerside) {
                    If ($OldWidth -gt $OldHeight) {
                        $ratio = $OldHeight / $OldWidth
                        $width = $Longerside
                        $height = $ratio * $Longerside

                    }
                    If ($OldWidth -lt $OldHeight) {
                        $ratio = $OldWidth / $OldHeight
                        $height = $Longerside
                        $width = $ratio * $Longerside
                    }
                    If ($OldWidth -eq $OldHeight) {
                        $Width = $Longerside
                        $Height = $Longerside
                    }
                }



                $Bitmap = New-Object -TypeName System.Drawing.Bitmap -ArgumentList $Width, $Height
                $NewImage = [System.Drawing.Graphics]::FromImage($Bitmap)
        
                #Retrieving the best quality possible
                $NewImage.SmoothingMode = $SmoothingMode
                $NewImage.InterpolationMode = $InterpolationMode
                $NewImage.PixelOffsetMode = $PixelOffsetMode
                $NewImage.DrawImage($OldImage, $(New-Object -TypeName System.Drawing.Rectangle -ArgumentList 0, 0, $Width, $Height))
               
    
                $OldImage.Dispose()                

                If ($PSCmdlet.ShouldProcess("Resized image based on $Path", "save to $OutputPath")) {
                    $Bitmap.Save($OutputPath)                   
                    Compress-Image -type "jpg" -path $OutputPath        
                }
    
                $Bitmap.Dispose()
                $NewImage.Dispose()            
        }
                    }
        catch 
        {
            Throw "$($_.Exception.Message)"
        }
        
    }
}

function Compress-Image() {
    [CmdLetBinding(
        SupportsShouldProcess = $true, 
        PositionalBinding = $false,
        ConfirmImpact = "Low",
        DefaultParameterSetName = "Absolute"
    )]
    param(
        [Parameter(Mandatory = $True)][System.String]$type,
        [Parameter(Mandatory = $True)][string]$path,
        [Parameter(Mandatory = $False)][Switch]$report
    )

    $params = switch ($type) {
        "jpg" { "-compress jpeg -quality 82" }
        "gif" { "-fuzz 10% -layers Optimize" }
        "png" { "-depth 24 -define png:compression-filter=2 -define png:compression-level=9 -define png:compression-strategy=1" }
    }

    if ($report) {
        # Write-Output ""
        # Write-Output "Listing $type files that would be included for compression with params: $params"
    } else {
        # Write-Output ""
        # Write-Output "Compressing $type files with parameters: $params"
    }
    
    Get-Item $path -Include "*.$type" | 
        Where-Object {
            $_.Length/1kb -gt $minSize
        } | 
        Sort-Object -Descending length |
        ForEach-Object {
            $file = "'" + $_.FullName + "'"
        
            if ($report) {
                # $fSize = Get-Size-Kb($file)
                # Write-Output "$file - $fSize"
            } else {
                if ($verbose) {
                    # Write-Output "Compressing $file"
                    # $fileStartSize = Get-Size-Kb($file)
                }
        
                # compress image
                if ($report -eq $False) {
                    Invoke-Expression "magick $file $params $file"
                }

                if ($verbose) {
                    # $fileEndSize = Get-Size-Kb($file)
                    # Write-Output "Reduced from $fileStartSize to $fileEndSize"
                }

                $Global:FinalTotal += Get-Size-Item-mb($path)                
            }
        }
}


Function Get-pathfile{
    [CmdletBinding()]
    param (      
        [Parameter(Position = 0, Mandatory=$true)]    
        [ValidateNotNullOrEmpty()]
        [string]$IncludeExcludePath
    )

    # add module as administrator with 
    # Install-Module -Name ImportExcel -Force
    # path with excel files
    # (assuming you downloaded the sample data as instructed before)
    Set-Location -Path $IncludeExcludePath
    # Get-Help Import-Excel
    # $excel_obj = Import-Excel -Path .\financial.xlsx | Where-Object 'Month Number' -eq 12
    $paths_include = Import-Excel -Path .\files\path_list.xlsx -WorkSheetname 'include'
    $paths_exclude = Import-Excel -Path .\files\path_list.xlsx -WorkSheetname 'exclude'
    
    $return_hash = @{}
    $include_array = [System.Collections.ArrayList]::new()
    $exclude_array = [System.Collections.ArrayList]::new()
    
    foreach ($row in $paths_include)
    {
        if($row.type -eq 'path'){
            [void]$include_array.Add($row.value)
        }   
    }
    
    foreach ($row in $paths_exclude)
    {
        if($row.type -eq 'path'){
            [void]$exclude_array.Add($row.value)
        }   
    }
    
    [void]$return_hash.Add("include",$include_array)
    [void]$return_hash.Add("exclude",$exclude_array)

    
    return $return_hash
}

Function Get-Imagepaths{
    [cmdletbinding()]
    param( 
        [Parameter(Position = 0, Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [hashtable]$ExcelPaths 
        ) 

    $exclude_list = $ExcelPaths['exclude']
    $include_list = $ExcelPaths['include']

    $recursive_paths = [System.Collections.ArrayList]::new()

    foreach($path in $include_list){
        $systempath = Get-Item -Path $path
        $recursive_paths.Add($systempath.FullName) | Out-Null # https://stackoverflow.com/questions/10286164/function-return-value-in-powershell

        Get-ChildItem -Path $include_list -Directory -Recurse 
        | ForEach-Object{
            $allowed = $true
            foreach ($exclude in $exclude_list) { 
                if (($_.Parent -ilike $exclude) -Or ($_ -ilike $exclude)) {
                    $allowed = $false
                    break
                }
            }
            if ($allowed) {
                $recursive_paths.Add($_.FullName)
            }
        } | Out-Null # https://stackoverflow.com/questions/7325900/powershell-2-array-size-differs-between-return-value-and-caller

    }
    
    return $recursive_paths
}

Function Get-imagelist{
    [cmdletbinding()]
    param (
        [Parameter(Position = 0, Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [System.Array]$paths,
        [Int]$Width,
        [Int]$Height,
        [Int]$BatchAmount = 5000
    )
    $ImageList = [System.Collections.ArrayList]::new()
    $arrpathlist = [System.Collections.ArrayList]$paths

    $counter = 1
    :outer
    foreach($path in $arrpathlist){
        Get-ChildItem -Path $path -Filter *.jpg |         
        ForEach-Object {
            $t = [System.Drawing.Image]::FromFile($_.FullName)             
            if ($t.Width -gt $Width -or $t.Height -gt $Height ) {
                $ImageList.Add($_) 
                $t.Dispose()     
                if($counter -eq $BatchAmount){
                    [string]$outputStr = 'batch limit of {0} reached' -f $BatchAmount
                    Write-Host $outputStr -ForegroundColor Magenta | Out-Null
                    break outer #breaking named loop https://stackoverflow.com/questions/36025696/break-out-of-inner-loop-only-in-nested-loop
                }                       
            }else{
                $t.Dispose() #need to close connection to bitmap so it can be overwritten  
            }
            $counter++
        } | 
        Out-Null         
    } 

    return [System.Collections.ArrayList]$ImageList

}

$longerSide = 3000
Write-Output "Started processing $(Get-Date -Format u)"
$ExcelPaths = Get-pathfile -IncludeExcludePath $PSScriptRoot
$paths = Get-Imagepaths -ExcelPaths $ExcelPaths
$image_list = Get-imagelist -paths $paths -Width $longerSide -Height $longerSide -batch 5000
Resize-Image -ImagePath $image_list -Longerside $longerSide -OverWrite -InterpolationMode Default -SmoothingMode Default -PixelOffsetMode Default
Write-Output "end processing $(Get-Date -Format u)"
Write-Output "Origional storage used $Global:OrigionalTotal MB : Storage used after compression $Global:FinalTotal MB)"
