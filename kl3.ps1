# " íéáűúőóüö_ÍÉÁŰÚŐÓÜÖ " > charset.sql
#  get-content -Path $args[0] -encoding utf8 | out-file $tempfile -encoding Unicode
# get-content "d:\prg\Reporting\Reports\kl\kl_rep_dsl.sql" | out-file "d:\prg\Reporting\Reports\kl\kl_rep_dsl2.sql" -encoding Unicode
. d:\prg\Reporting\Modules\GodotHelper\GodotHelper.ps1 
#. d:\prg\Reporting\Modules\GodotHelper\Send-Mail-Deferred.ps1
##  d:\prg\kl_ps\fs_riport\kl.ps1 

cd \prg\kl_ps\forgalom
Start-Log "naplo.log"
Set-LogLevel Debug
Write-Log "Indult0"

Db-Connect2 "dwhprod" "kecskemetil" "Ildiko_02"   ## diamond
# Db-Connect2 "dwhprod" "kecskemetil" "Ildiko_02"   ## cldb

$count =0


Write-Log "Indult_00"

$sql = Load-File ("d:\prg\kl_ps\forgalom\kl2.sql")
Write $sql

Write-Log "Indult_11"
$results = Execute-Query $sql
if ($results.IsEmpty) { 
  Write "Figyi, nincs eredmény a tömbbe !!!" 
  Write-Log "Figyi, nincs eredmény a tömbbe !!!"
} else {
  Write "Kész"
}

Create-Workbook 
Write-Log "Vég_11"
Add-Worksheet $results "201202"
Write-Log "Indult_12"

$sql = Load-File ("d:\prg\kl_ps\forgalom\kl3.sql")
Write $sql

$results = Execute-Query $sql
if ($results.IsEmpty) { 
  Write "Figyi, nincs eredmény a tömbbe !!!" 
  Write-Log "Figyi, nincs eredmény a tömbbe !!!"
} else {
  Write "Kész"
}

Write-Log "Vég_12"
Add-Worksheet $results "201203"
Write-Log "Indult_13"

$sql = Load-File ("d:\prg\kl_ps\forgalom\kl4.sql")
Write $sql


$results = Execute-Query $sql
if ($results.IsEmpty) { 
  Write "Figyi, nincs eredmény a tömbbe !!!" 
  Write-Log "Figyi, nincs eredmény a tömbbe !!!"
} else {
  Write "Kész"
}

Write-Log "Vég_13"
Add-Worksheet $results "201204"
Write-Log "Indult_14"

$sql = Load-File ("d:\prg\kl_ps\forgalom\kl5.sql")

$results = Execute-Query $sql
if ($results.IsEmpty) { 
  Write "Figyi, nincs eredmény a tömbbe !!!" 
  Write-Log "Figyi, nincs eredmény a tömbbe !!!"
} else {
  Write "Kész"
}

Write-Log "Vég_14"
Add-Worksheet $results "201205"

#--------------------------------

$klfile_nev = "KL" ;

##Save-Workbook $kltam[$count]+"kl_eredmeny.xlsx"
$klfile_nev = $klfile_nev + "kesz.xlsx"
Save-Workbook( $klfile_nev)

#   Send-Mail-Deferred "kecskemeti.lajos@t-systems.hu" "kecskemeti.lajos@telekom.hu" "TAM heti riport" "Üzenettörzs" $klfile_nev+".xlsx"
$emailto = "kecskemeti.lajos@kfkizrt.hu"
$subject = "TAM heti riport"
$body = "TAM heti riport (Fodor Sándor)"
Send-Mail "kecskemeti.lajos@telekom.hu" $emailto $subject $body $klfile_nev
   
   Write "Levél ment"
   Write-Log "Levél ment"



Dispose-Db
$results.Dispose()

End-Log

