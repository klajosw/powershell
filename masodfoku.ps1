# M�sodfok� egyenlet megold�k�plet sz�mol�s
# a*x^2 + b*x + c = 0
write �M�sodfok� egyenlet megold�k�plet sz�mol�s�

# a v�ltoz� bek�r�se
[int] $a = Read-Host �Add meg az a �rt�k�t:�
# b v�ltoz� bek�r�se
[int] $b = Read-Host �Add meg a b �rt�k�t:�
# c v�ltoz� bek�r�se
[int] $c = Read-Host �Add meg a c �rt�k�t:�

# determin�ns �rt�k�t�l f�gg a val�s gy�k(�k) l�tez�se
# d = b^2 � 4*a*c

$d = ($b*$b) � (4*$a*$c)

# d < 0 -> nincs gy�k
# d = 0 -> egy val�s gy�k
# x = -b / (2*a)
# d > 0 -> 2 val�s gy�k
# x1 = (-b + d) / (2*a)
# x2 = (-b � d) / (2*a)

if($d -lt 0)
{
�Nincs val�s gy�k�
}
elseIf ($d -eq 0)
{
$x = (-$b) / (2*$a);
Write-Host �Egy darab gy�k van�
write $x
}
else
{
$x1 = (-$b + [System.Math]::Sqrt($d) ) / (2*$a);
$x2 = (-$b � [System.Math]::Sqrt($d) ) / (2*$a);
�K�t val�s gy�k van�
�Els�:�
$x1
�M�sodik:�
$x2
}