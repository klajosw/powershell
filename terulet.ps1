param($a, $b) 

if([int]$a -gt 0 -and [int]$b -gt 0) {
	if([int]$a -eq [int]$b) {
		$kerulet=4*$a
		$terulet=$a*$a
		echo "ez egy négyzet" }
	else {
		$kerulet=2*$a+2*$b
		$terulet=$a*$b
		echo "ez egy téglalap"
		}
	echo "kerület: " $kerulet
	echo "terulet: " $terulet 
	}
else {

echo "minden paraméternek pozitivnak kell lennie"
}