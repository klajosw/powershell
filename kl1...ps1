$test='
sz�veg
#a13
valami
#14
m�g valami
#c15
v�ge
';

$x=[regex]::split($test,"#"); 

$x[2];

 Foreach($z in $x) { 
      $z="EDI_DC40"+$z; 
      Out-File -filepath $j -inputobject $z -Encoding default; $j++;
   } 
##############
Get-Content("File.txt") | Out-String | %{ 
   $j=1; 
   $x=[regex]::split($_,"EDI_DC40"); 
   Foreach($z in $x) { 
      $z="EDI_DC40"+$z; 
      Out-File -filepath $j -inputobject $z -Encoding default; $j++;
   } 
};

###########
"A z�r�jelek k�zti (sz�veget) keresem" -match "\([^)]*\)"

"Minta az elej�n, v�g�n is minta" -match "minta"
$matches[0]


##### T�bbsz�ros tal�lat
$minta = [regex] "\w+"
$eredmeny = $minta.matches("Ez a sz�veg t�bb sz�b�l �ll")
$eredmeny[1].Value
foreach($elem in $eredm�ny){$elem.value}

###
 ("Ebben {van}, minden-f�le (elv�laszt�) [jel]").split()
Ebben
{van},
minden-f�le
(elv�laszt�)
[jel]

###
$test='
sz�veg
#a13
valami
#14
m�g valami
#c15
v�ge
';

$tomb = $test -split "#";
$tomb[2]

###
$test='sz�veg#1#1#2#c3#c3';
$tomb = $test -split "#";
$tomb = $tomb | select -uniq
$tomb ## [2]


###
$input_path = �c:\ps\emails.txt�
$output_file = �c:\ps\extracted_addresses.txt�
$regex = �\b[A-Za-z0-9._%-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,4}\b�
select-string -Path $input_path -Pattern $regex -AllMatches | % { $_.Matches } | % { $_.Value } > $output_file

###
$input_path = �c:\ps\ip_addresses.txt�
$output_file = �c:\ps\extracted_ip_addresses.txt�
$regex = �\b\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}\b�
select-string -Path $input_path -Pattern $regex -AllMatches | % { $_.Matches } | % { $_.Value } > $output_file










