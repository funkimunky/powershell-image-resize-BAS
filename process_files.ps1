Add-Type -AssemblyName System.Drawing
Function Get-inclusions_exclusions(){
    # add module as administrator with 
    # Install-Module -Name ImportExcel -Force
    # path with excel files
    # (assuming you downloaded the sample data as instructed before)
    Set-Location -Path "C:\Temp\excelsampledata\"
    # Get-Help Import-Excel
    # $excel_obj = Import-Excel -Path .\financial.xlsx | Where-Object 'Month Number' -eq 12
    $paths_include = Import-Excel -Path .\path_list.xlsx -WorkSheetname 'include'
    $paths_exclude = Import-Excel -Path .\path_list.xlsx -WorkSheetname 'exclude'
    
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
    param( $path_hash ) 

    $exclude_list = $path_hash['exclude']
    $include_list = $path_hash['include']

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

Function Get-images(){
    param (
        $thispathlist
    )
    $ImageList = [System.Collections.ArrayList]::new()
    $arrpathlist = [System.Collections.ArrayList]$thispathlist
    foreach($path in $arrpathlist){
        # $images = Get-ChildItem -Path $path -Filter *.png | Where-Object { $_.Width -gt 400 -and $_.Height -gt 400 }
        Get-ChildItem -Path $path -Filter *.png|ForEach-Object{$ImageList.Add($_)}|Out-Null   
    }   

    return $ImageList

}


$path_hash = Get-inclusions_exclusions
$paths = Get-Paths($path_hash)
$image_list = Get-images($paths)
# Get-Paths($path_hash)
$test