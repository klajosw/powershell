param([int]$a, [int]$b) 

if ($a -le 0 -and $b -le 0)
{
echo "adjon meg k�t pozitiv eg�sz sz�mot param�terk�nt"
} else {

if([int]$a -gt 0 -and [int]$b -gt 0) {
	if([int]$a -eq [int]$b) {
		$kerulet=4*$a
		$terulet=$a*$a
		echo "ez egy n�gyzet" }
	else {
		$kerulet=2*$a+2*$b
		$terulet=$a*$b
		echo "ez egy t�glalap"
		}
	echo "ker�let: " $kerulet
	echo "terulet: " $terulet 
	}
else {

echo "minden param�ternek pozitivnak kell lennie"
}
}