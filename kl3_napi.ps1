# " íéáűúőóüö_ÍÉÁŰÚŐÓÜÖ " > charset.sql
## ---------------------------- Szepeshelyi István napi  -----------------------------------
#  get-content -Path $args[0] -encoding utf8 | out-file $tempfile -encoding Unicode
# get-content "d:\prg\Reporting\Reports\kl\kl_rep_dsl.sql" | out-file "d:\prg\Reporting\Reports\kl\kl_rep_dsl2.sql" -encoding Unicode
### . d:\prg\Reporting\Modules\GodotHelper\GodotHelper.ps1 
##  . d:\amo\Reporting\Modules\GodotHelper\GodotHelper.ps1
#   . d:\prg\Reporting\Modules\GodotHelper\Send-Mail-Deferred.ps1
##  d:\prg\kl_ps\fs_riport\kl.ps1 
##  . d:\amo\Users\kecskemet1l314\kl_ps\GodotHelper.ps1

### Kill
# Get-Process | Where-Object -FilterScript {$_.name -eq 'AxCrypt' } | %{ kill $_.Id}

. d:\amo\Users\kecskemet1l314\kl_ps\GodotHelper.ps1

d:
cd \amo\Users\kecskemet1l314\kl_ps\axcrypt

Start-Log "naplo.log"
Set-LogLevel Debug
Write-Log "Indult0"
$kltam_email = Get-Content "d:\amo\Users\kecskemet1l314\kl_ps\axcrypt\tam_email.txt"
$kllevel = [string]::Join([Environment]::NewLine, (Get-Content "d:\amo\Users\kecskemet1l314\kl_ps\axcrypt\level.txt"))

$kl_jsz_ = Get-Content "d:\amo\Users\kecskemet1l314\kl_ps\jsz.txt"
$kl_jsz = $kl_jsz_[0].split(" ")    ### dwhprod
##$kl_jsz = $kl_jsz_[1].split(" ")  ### cldbpr
Db-Connect2 $kl_jsz[0] $kl_jsz[1] $kl_jsz[2]
## Db-Connect2 "cldbpr" "kecskemetil" "Ildiko_03" 

$count =0

Write-Log "Indult_00"

$sql = Load-File ("d:\amo\Users\kecskemet1l314\kl_ps\axcrypt\kl11.sql")
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
Add-Worksheet $results "FIX utolsó két nap"
#-----------------------------------------
$sql = Load-File ("d:\amo\Users\kecskemet1l314\kl_ps\axcrypt\kl22.sql")
Write-Log "Indult_22"
$results = Execute-Query $sql
if ($results.IsEmpty) { 
  Write "Figyi, nincs eredmény a tömbbe !!!" 
  Write-Log "Figyi, nincs eredmény a tömbbe !!!"
} else {
  Write "Kész"
}

Write-Log "Vég_22"
Add-Worksheet $results "MOBIL utolsó két nap"
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
   $subject = "Adattárház napi riport [T-Systems ügyfélelégedettségi minta - [ $datumom ] "
   $body = "  "
   $body = $body + $kllevel ;
   Send-Mail $emailfrom $kltam_email_ki $subject $body $klfile_nev 
   
   Write "Levél ment"
   Write-Log "Levél ment"


Dispose-Db
$results.Dispose()

## ellenőrzés és ha nem létezik létrehoz
$path_kl = ".\xls_$datumom"
while(!(test-path $path_kl)){new-item -ItemType Directory -Path $Path_kl}

## New-Item ".\xls_$datumom" -type directory

##Move-Item *.xlsx ".\xls_$datumom"  -force
Copy-Item *.xlsx ".\xls_$datumom"  -force

## start indit.bat

Write-Log "Fájl titkosítása kezd"
## & "C:\Program Files\Axantum\AxCrypt\AxCrypt.exe" -b 2 -e -k "Mt107" -z ue_rip.xlsx >> naplo.log
##"C:\Program Files\Axon Data\AxCrypt\1.6.4.4-3\AxCrypt.exe" -b 2 -e -k "Mt107" -z ue_rip.xlsx


& "c:\Program Files\Axantum\AxCrypt\AxCrypt.exe" -b 2 -e -k "Mt107" -z ue_rip.xlsx 

### & "c:\Program Files\Axantum\AxCrypt\AxCrypt.exe" -b 2 -e -k "Mt107" -z ue_rip.xlsx 

Write-Log "Fájl titkosítása veg"

## várakozzunk
Start-Sleep -s 8

$klfile_nev = "ue_rip-xlsx.axx"

## csak teszt
$kltam_email_ki = "kecskemeti.lajos@t-systems.hu"


$subject = "Adattárház napi riport [T-Systems ügyfélelégedettségi minta - [ $datumom ] (axx) "
# $subject = "Ügyfél elégedetség napi riport (axx) "
##$kltam_email_ki = "kecskemeti.lajos@t-systems.hu, Szepeshelyi.Istvan@t-systems.hu"

## végleges :
$kltam_email_ki = "kecskemeti.lajos@t-systems.hu, Szepeshelyi.Istvan@t-systems.hu, iccaminta@szociograf.hu"

Send-Mail $emailfrom $kltam_email_ki $subject $body $klfile_nev 

Move-Item *.axx ".\xls_$datumom"  -force


Get-Process | Where-Object -FilterScript {$_.name -eq 'AxCrypt' } | %{ kill $_.Id}

