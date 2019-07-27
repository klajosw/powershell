# Másodfokú egyenlet megoldóképlet számolás
# a*x^2 + b*x + c = 0
write “Másodfokú egyenlet megoldóképlet számolás”

# a változó bekérése
[int] $a = Read-Host “Add meg az a értékét:”
# b változó bekérése
[int] $b = Read-Host “Add meg a b értékét:”
# c változó bekérése
[int] $c = Read-Host “Add meg a c értékét:”

# determináns értékétõl függ a valós gyök(ök) létezése
# d = b^2 – 4*a*c

$d = ($b*$b) – (4*$a*$c)

# d < 0 -> nincs gyök
# d = 0 -> egy valós gyök
# x = -b / (2*a)
# d > 0 -> 2 valós gyök
# x1 = (-b + d) / (2*a)
# x2 = (-b – d) / (2*a)

if($d -lt 0)
{
“Nincs valós gyök”
}
elseIf ($d -eq 0)
{
$x = (-$b) / (2*$a);
Write-Host “Egy darab gyök van”
write $x
}
else
{
$x1 = (-$b + [System.Math]::Sqrt($d) ) / (2*$a);
$x2 = (-$b – [System.Math]::Sqrt($d) ) / (2*$a);
“Két valós gyök van”
“Elsõ:”
$x1
“Második:”
$x2
}