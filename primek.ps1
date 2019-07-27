[int]$szam=read-host "adj meg egy szamot!"
$print=@()
$i=2
while ($szam -gt 1) 
	{	
		if (($szam % $i) -eq 0) 
		{
		$szam=$szam/$i
		$print+=$i
		}
		else 
		{
		$i++
		}
	}
foreach ($elem in $print)
	{
	$elem
	}