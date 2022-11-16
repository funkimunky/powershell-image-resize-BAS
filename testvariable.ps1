$x = 1

function changeit {
    $x += 5
    Write-Output "x in function scope now $x"
    Write-Output "Changing $x. Was: $x"
    # $script:x = 2
    $x = 1
    Write-Output "New value: $x"
}

Write-Output "$x has value $x"
changeit
Write-Output "$x has value $x"