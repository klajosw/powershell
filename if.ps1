#---------------------------------------------------
#
#  ne legyen tartalom.txt fajl a futasa elott
#---------------------------------------------------
if (Test-Path tartalom.txt) {"Van tartalom.txt fajl"}
else 
{"Nincs tartalom.txt fajl"}

New-Item -path tartalom.txt  -Type File
Write-Host "Most hoztam letre"

if (Test-Path tartalom.txt) {"Van tartalom.txt f�jl"} 
else 
{"Nincs tartalom.txt fajl"}
