
Disable-Mailbox -Identity olahzoltanteszt
Get-Mailbox -Identity olahzoltanteszt
Get-QADuser -SamAccountName olahzoltanteszt
enable-Mailbox -Identity olahzoltanteszt -Database "k-mail3\nrml11\nrml11"

get-user | where-object{$_.RecipientType –eq "User"} | Enable-Mailbox –Database "servername\Mailbox Database" | get-mailbox | select name,windowsemailaddress,database

Get-QADuser -SamAccountName olahzoltanteszt | New-Mailbox -Database "k-mail3\nrml11\nrml11" -Name $_.DisplayName -Password $password -UserPrincipalName $_.UserPrincipalName

Get-QADuser -SamAccountName olahzoltanteszt | Enable-Mailbox -Database "k-mail3\nrml11\nrml11"

$password = ConvertTo-SecureString 'Pa$$w0rd' -AsPlainText -Force


$a = Get-QADuser -SamAccountName olahzoltanteszt 
$displayname = $a.DisplayName
$upn = $a.UserPrincipalName

New-Mailbox -Database "k-mail3\nrml11\nrml11" -Name $displayname -Password $password -UserPrincipalName $upn
