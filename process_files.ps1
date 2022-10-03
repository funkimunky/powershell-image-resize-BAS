Add-Type -AssemblyName System.Drawing
Import-Module ./Resize-Image/Resize-Image -force

Import-Module -Name $PSScriptRoot/modules/ImportExcel -force
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

Function Get-Paths{
    [cmdletbinding()]
    param( 
        [hashtable]$ExcelPaths 
        ) 

    $exclude_list = $ExcelPaths['exclude']
    $include_list = $ExcelPaths['include']

    $recursive_paths = [System.Collections.ArrayList]::new()

    foreach($path in $include_list){
        $path = Get-Item -Path $path
        $recursive_paths.Add($path.FullName) | Out-Null # https://stackoverflow.com/questions/10286164/function-return-value-in-powershell

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

Function Get-images{
    [cmdletbinding()]
    param (
        [System.Array]$ImagePaths,
        [Int]$Width,
        [Int]$Height,
        [Int]$BatchAmount = 10
    )
    $ImageList = [System.Collections.ArrayList]::new()
    $arrpathlist = [System.Collections.ArrayList]$ImagePaths

    $counter = 1
    :outer
    foreach($path in $arrpathlist){
        Get-ChildItem -Path $path -Filter *.jpg |         
        ForEach-Object {
            $t = [System.Drawing.Image]::FromFile($_.FullName)             
            if ($t.Width -gt $Width -and $t.Height -gt $Height ) {
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

Function Process_Images{
    [cmdletbinding()]
    param (
        [System.Collections.ArrayList]$ImageList,
        [string]$OverWrite,
        [string]$WhatIf = $null
    )
     
    foreach($Image in $ImageList){
        # Resize-Image -width 400 -MaintainRatio -ImagePath $Image -OverWrite $OverWrite -WhatIf
        if($WhatIf){
            Resize-Image -ImagePath $Image -Longerside 1000 -OverWrite $OverWrite -WhatIf
        }else {
            Resize-Image -ImagePath $Image -Longerside 1000 -OverWrite $OverWrite -WhatIf
        }
        
        
    }

}


$ExcelPaths = Get-inclusions_exclusions -IncludeExcludePath $PSScriptRoot
$paths = Get-Paths -ExcelPaths $ExcelPaths
$image_list = Get-images -ImagePaths $paths -Width 1000 -Height 1000 -batch 20
Process_Images -ImageList $image_list -OverWrite y -WhatIf