param($fajl)
$sorok=get-content $fajl
$neg=@()
$poz=@()
for ($i=0; $i -lt $sorok.length; $i++) 
{
	if([int]$sorok[$i] -gt 0) 
	{
	$poz=$poz + $sorok[$i]
	}
	elseif([int]$sorok[$i] -eq 0) # a null�t a pozit�vakhoz �ratom. ( dimat gyakra hivatkozva:) )  
	{
	$poz=$poz + $sorok[$i]
	}
	else 
	{
	$neg=$neg + $sorok[$i]
	}
}


echo "Pozitiv sz�mok"
for ($i=0; $i -lt $poz.length; $i++) 
	{
	echo $poz[$i]
	}
echo "Negativ sz�mok" 
for ($i=0; $i -lt $neg.length; $i++) 
	{
	echo $neg[$i]
	}