write-host "Szotar kezelese"
$szotar=@{kutya="dog";macska="cat";eger="mouse"}
write-host "Szavak:"
foreach ($magyar in $szotar.Keys){
        write-host ($magyar," ",$szotar[$magyar])
}
