# " íéáűúőóüö_ÍÉÁŰÚŐÓÜÖ " > charset.sql
#  get-content -Path $args[0] -encoding utf8 | out-file $tempfile -encoding Unicode
# get-content "d:\prg\Reporting\Reports\kl\kl_rep_dsl.sql" | out-file "d:\prg\Reporting\Reports\kl\kl_rep_dsl2.sql" -encoding Unicode
. d:\prg\Reporting\Modules\GodotHelper\GodotHelper.ps1 
##  d:\prg\kl_ps\fs_riport\kl.ps1 
cd \prg\kl_ps\fs_riport
Start-Log "naplo.log"
Set-LogLevel Debug
Write-Log "Indult0"

Db-Connect2 "cldb" "kecskemetil" "Ildiko_01" 

# $results = Execute-Query "select * from kecskemetil.MF_SZOLG_HA_2011 kp  where 1=1 and kp.t_period =  (select max(t_period) from kecskemetil.MF_SZOLG_HA_2011 )" 

$sql = Load-File ("d:\prg\kl_ps\fs_riport\kl1.sql")
$results = Execute-Query $sql
if ($results.IsEmpty) { 
  Write "Figyi, nincs eredmény a tömbbe !!!" 
  Write-Log "Figyi, nincs eredmény a tömbbe !!!"
} else {
  Write "Kész"
  Write-Log "Kész :)"
}

Write-Log "Indult1"
Create-Workbook 

Add-Worksheet $results "SIM"

$sql = Load-File ("d:\prg\kl_ps\fs_riport\kl2.sql")
 
$results = Execute-Query $sql
if ($results.IsEmpty) { 
  Write "Figyi, nincs eredmény a tömbbe !!!" 
  Write-Log "Figyi, nincs eredmény a tömbbe !!!"
} else {
  Write "Kész"
  Write-Log "Kész :)"
}

Add-Worksheet $results "SZOLG"

Save-Workbook "UJ2_eredmeny.xlsx"

Dispose-Db
$results.Dispose()

End-Log
