Add-Type -AssemblyName System.Drawing
Function Get-Excel(){
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

Function Get-images(){
    param (
        $path_hash
    )

    $exclude_list = $path_hash['exclude']
    $include_list = $path_hash['include']
    # $ImageList = Get-ChildItem -Path $include_list -Filter *.png -Recurse -Exclude $exclude_list | Where-Object { $_.Width -gt 400 -and $_.Height -gt 400 }
 
    $ImageList = Get-ChildItem -Path $include_list -Directory -Recurse 
    # | ForEach-Object{
    #     $allowed = $true
    #     foreach ($exclude in $exclude_list) { 
    #         $mynname = $_
    #         $parentname = $_.Parent
    #         if (($_.Parent -ilike $exclude) -Or ($_ -ilike $exclude)) { 
    #             $allowed = $false
    #             break
    #         }
    #     }
    #     if ($allowed) {
    #         $_
    #     }
    # } | Get-ChildItem -Filter *.png # | ForEach-Object { [System.Drawing.Image]::FromFile($_.FullName) } | Where-Object { $_.Width -gt 400 -and $_.Height -gt 400 }

   

    return $ImageList

}

$path_hash = Get-Excel
$image_list = Get-images($path_hash)