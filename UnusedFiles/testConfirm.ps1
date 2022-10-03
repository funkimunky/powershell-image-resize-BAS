function ShouldTestCallee {
    [cmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Low')]
    param($test)
    Begin {        
    }
    Process {
        if ($PSCmdlet.ShouldProcess($env:COMPUTERNAME, "Confirm?")) {
            Write-Host "test"
        }
    }
    end {}
} 

function ShouldTestCaller {     
    [cmdletBinding(SupportsShouldProcess=$true)]     
    param($test,$list) 
    
    try 
    {
        foreach($number in $list){
            ShouldTestCallee 
    } 
    }
    catch 
    {
        Throw "$($_.Exception.Message)"
    } 
    
}  
    
$list = 0..3
$ConfirmPreference = 'High' 
ShouldTestCaller -list $list -Confirm
# ShouldTestCaller -Confirm