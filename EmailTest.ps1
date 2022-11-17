# EmailTest.ps1

#Parameters for the email:

$To = "<To>"
$CC = "<CC>"
$Subject = "<Subject>"
$Body = "<Body>"


function CreateMailByOutlook {
    $Outlook = New-Object -ComObject Outlook.Application
    $Mail = $Outlook.CreateItem(0)
    $Mail.To = $To
    $Mail.CC = $CC
    $Mail.Subject = $Subject
    $Mail.Body = $Body
    $Inspector = $Mail.GetInspector
    $Inspector.Display() 
}

function Main {
    CreateMailByOutlook
}

