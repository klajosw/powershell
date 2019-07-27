param($a,$b)
#elsõ feladat
#[string] $a = Read-Host “Hogy hívnak?”
#echo "szia $a"
#[string] $b = Read-Host "Hány éves vagy?"
$ev=2011
$eletkor=$ev-$b
echo "Szia $a, $eletkor-ben születtél"


if ($b -ge 20) 
	{
	echo "Te már nagykorú vagy!"
	}
	else
	{
	"Még tini vagy"
	}

