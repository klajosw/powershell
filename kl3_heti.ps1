# " íéáűúőóüö_ÍÉÁŰÚŐÓÜÖ " > charset.sql
## ---------------------------- Szepeshelyi István napi  -----------------------------------
#  get-content -Path $args[0] -encoding utf8 | out-file $tempfile -encoding Unicode
# get-content "d:\prg\Reporting\Reports\kl\kl_rep_dsl.sql" | out-file "d:\prg\Reporting\Reports\kl\kl_rep_dsl2.sql" -encoding Unicode
### . d:\prg\Reporting\Modules\GodotHelper\GodotHelper.ps1 
##. d:\amo\Reporting\Modules\GodotHelper\GodotHelper.ps1
. d:\amo\Users\kecskemet1l314\kl_ps\GodotHelper.ps1
#. d:\prg\Reporting\Modules\GodotHelper\Send-Mail-Deferred.ps1
##  d:\prg\kl_ps\fs_riport\kl.ps1 

d:
cd \amo\Users\kecskemet1l314\kl_ps\axcrypt_he

Start-Log "naplo.log"
Set-LogLevel Debug
Write-Log "Indult0"
$kltam_email = Get-Content "d:\amo\Users\kecskemet1l314\kl_ps\axcrypt_he\tam_email.txt"
$kllevel = [string]::Join([Environment]::NewLine, (Get-Content "d:\amo\Users\kecskemet1l314\kl_ps\axcrypt_he\level.txt"))

$kl_jsz_ = Get-Content "d:\amo\Users\kecskemet1l314\kl_ps\jsz.txt"
$kl_jsz = $kl_jsz_[0].split(" ")    ### dwhprod
##$kl_jsz = $kl_jsz_[1].split(" ")  ### cldbpr
Db-Connect2 $kl_jsz[0] $kl_jsz[1] $kl_jsz[2]
## Db-Connect2 "cldbpr" "kecskemetil" "Ildiko_03" 

$count =0

Write-Log "Indult_00"

$sql = Load-File ("d:\amo\Users\kecskemet1l314\kl_ps\axcrypt_he\kl11.sql")
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
Add-Worksheet $results "FIX utolsó hét"
#-----------------------------------------
$sql = Load-File ("d:\amo\Users\kecskemet1l314\kl_ps\axcrypt_he\kl22.sql")
Write-Log "Indult_22"
$results = Execute-Query $sql
if ($results.IsEmpty) { 
  Write "Figyi, nincs eredmény a tömbbe !!!" 
  Write-Log "Figyi, nincs eredmény a tömbbe !!!"
} else {
  Write "Kész"
}

Write-Log "Vég_22"
Add-Worksheet $results "MOBIL utolsó hét"
#--------------------------------


Write-Log "Indult_88"
$klfile_nev = "ue_rip" ;

##Save-Workbook $kltam[$count]+"kl_eredmeny.xlsx"
$klfile_nev = $klfile_nev + ".xlsx"
Save-Workbook( $klfile_nev)

##   $kltam_email_ki = "kecskemeti.lajos@t-systems.hu, Szepeshelyi.Istvan@t-systems.hu"
##   $kltam_email_ki = "kecskemeti.lajos@freemail.hu"
   $kltam_email_ki = "kecskemeti.lajos@telekom.hu"
   $datumom = Get-Date -format "yyyy_MM_d"

   $emailfrom = "kecskemeti.lajos@kfkizrt.hu"
   $subject = "Adattárház heti riport [T-Systems ügyfélelégedettségi minta - [ $datumom ] "
   $body = "  "
   $body = $body + $kllevel ;
   Send-Mail $emailfrom $kltam_email_ki $subject $body $klfile_nev 
   
   Write "Levél ment"
   Write-Log "Levél ment"


Dispose-Db
$results.Dispose()

End-Log


New-Item ".\xls_$datumom" -type directory
##Move-Item *.xlsx ".\xls_$datumom"  -force
Copy-Item *.xlsx ".\xls_$datumom"  -force

start indit.bat

##"c:\Program Files\Axantum\AxCrypt\AxCrypt.exe" `-b 2 `-e `-k "Mt107" `-z ue_rip.xlsx
##"C:\Program Files\Axon Data\AxCrypt\1.6.4.4-3\AxCrypt.exe" -b 2 -e -k "Mt107" -z ue_rip.xlsx

## várakozzunk
Start-Sleep -s 3

$klfile_nev = "ue_rip-xlsx.axx"
## $kltam_email_ki = "kecskemeti.lajos@t-systems.hu"
$subject = "Adattárház heti riport [T-Systems ügyfélelégedettségi minta - [ $datumom ] (axx) "
# $subject = "Ügyfél elégedetség napi riport (axx) "
$kltam_email_ki = "kecskemeti.lajos@t-systems.hu, Szepeshelyi.Istvan@t-systems.hu"
## $kltam_email_ki = "kecskemeti.lajos@t-systems.hu, Szepeshelyi.Istvan@t-systems.hu, iccaminta@szociograf.hu"
Send-Mail $emailfrom $kltam_email_ki $subject $body $klfile_nev 

Move-Item *.axx ".\xls_$datumom"  -force


