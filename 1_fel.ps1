param($a,$b)
#els� feladat
#[string] $a = Read-Host �Hogy h�vnak?�
#echo "szia $a"
#[string] $b = Read-Host "H�ny �ves vagy?"
$ev=2011
$eletkor=$ev-$b
echo "Szia $a, $eletkor-ben sz�lett�l"


if ($b -ge 20) 
	{
	echo "Te m�r nagykor� vagy!"
	}
	else
	{
	"M�g tini vagy"
	}

