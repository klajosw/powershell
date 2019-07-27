param($fajl)
$sorok=get-content $fajl
$jo_kod_db=0
for ($i=0; $i -lt $sorok.length; $i++) 
{
	if($sorok[$i] -match "^[a-z]{3}-[0-9][A-B]$") 
	{
	$jo_kod_db++;
	}
}

echo "a megadott fájlban "$jo_kod_db" db helyed kód van"
	
	