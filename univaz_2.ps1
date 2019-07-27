$nevek = New-Object system.collections.arraylist

Get-User -OrganizationalUnit "kfki.corp/IQSYS/Felhasznalok/Normal" | ForEach-Object { [void] $nevek.add($_.get_SamAccountName())}


$nevek | ForEach-Object {Get-Mailbox -Identity $_} | foreach-object `
{ 
$b = [Microsoft.Exchange.Data.ProxyAddressCollection] $_.emailaddresses
	#több email cim közül a megfeleő addrassstringböl kivágjuk az univaz-t, majd a tömbhöz fűzzük.
	for ($i=0; $i -ile $b.count; $i++)
		{ 
			if ($b[$i].Addressstring -imatch "HUIQ")
				{
					$univaz =$b[$i].Addressstring -replace ("HUIQ","") -replace ("@iqsys.hu","")
					$_.samaccountname
					$univaz
					
				}
		}
}
