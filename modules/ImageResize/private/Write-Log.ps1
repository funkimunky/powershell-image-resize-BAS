Function Write-Log{
    <#
    .SYNOPSIS
        Writes string to log file
    .DESCRIPTION
        Writes string to log file       
    .EXAMPLE
        Write-Log("log content") -Verbose
        Explanation of the function or its result. You can include multiple examples with additional .EXAMPLE lines
    #>    
    
    [cmdletbinding()]
    param(
        [string] $Text
    )
    $fileDate = Get-Date -Format FileDate
    $fileDateTime = Get-Date -Format ddMMyy-HH:mm
    $exportPath = "C:\temp\process{0}.txt" -f $fileDate
    $logText = "{0}`t{1}" -f $fileDateTime, $Text

    $logText | Tee-Object -FilePath $exportPath -Append
}