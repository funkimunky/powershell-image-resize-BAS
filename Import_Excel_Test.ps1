
Function Get_Excel(){
    # add module as administrator with 
    # Install-Module -Name ImportExcel -Force
    # path with excel files
    # (assuming you downloaded the sample data as instructed before)
    Set-Location -Path "C:\Temp\excelsampledata\"
    # Get-Help Import-Excel
    # $excel_obj = Import-Excel -Path .\financial.xlsx | Where-Object 'Month Number' -eq 12
    $paths_include = Import-Excel -Path .\path_list.xlsx -WorkSheetname 'include'
    $paths_exclude = Import-Excel -Path .\path_list.xlsx -WorkSheetname 'exclude'

    $return_array =  [System.Collections.ArrayList]::new()
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

    [void]$return_array.Add($include_array)
    [void]$return_array.Add($exclude_array)

    return $return_array
}
