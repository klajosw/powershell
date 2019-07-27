$proba = New-Object system.collections.arraylist

Get-User -OrganizationalUnit "kfki.corp/IQSYS/Felhasznalok/Alvallalkozok" | foreach-object `
{ 
	$proba1 = Get-QADUser  $_.get_SamAccountName() -IncludeAllProperties -SerializeValues -usedefaultexcludedproperties $true -ExcludedProperties usercertificate,msexchumrecipientdialplanlink,msexchumpinchecksum,msrtcsip-userpolicy,msexchumtemplatelink
	[void] $proba.add($proba1)
	
}

#out-file -inputobject $proba -encoding unicode  -FilePath c:\temp\gabocs1.csv