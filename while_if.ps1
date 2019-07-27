$szavak="alma", "korte", "banan", "szilva", "szolo","barack"
$i=0;
while ($i -lt $szavak.length) {
 if ($szavak[$i] -match "^B"){
        write-host ($szavak[$i], " B-vel kezdodik") 
 }
 $i++
 }
