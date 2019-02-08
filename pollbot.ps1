## Author: 0xVox
## A vote on who was going to greggs was had, I was waiting for a build, I really didn't want to go - this happened.

param(
    [Parameter(Mandatory=$true)][string]$option,
    [Parameter(Mandatory=$true)][int]$numberOfVotes,
    [Parameter(Mandatory=$true)][string]$pollUrl
)

$ie = New-Object -ComObject 'internetExplorer.Application'
$ie.Visible = $true

foreach($num in 1..$numberOfVotes){
    Write-Output "Voting $num'th time"
    $ie.Navigate($pollUrl)
    while ($ie.Busy -eq $true){Start-Sleep -seconds 1;}

    $radioButton = $ie.Document.IHTMLDocument3_GetElementByID("field-options-$option")
    $radioButton.checked = $true

    $link=$ie.Document.IHTMLDocument3_getElementsByTagName("button") | where-object {$_.type -eq "submit"}
    $link.click()

    while ($ie.Busy -eq $true){Start-Sleep -seconds 1;}
}

$ie.Quit()