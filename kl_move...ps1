# " íéáűúőóüö_ÍÉÁŰÚŐÓÜÖ " > charset.sql
#  get-content -Path $args[0] -encoding utf8 | out-file $tempfile -encoding Unicode
# get-content "d:\prg\Reporting\Reports\kl\kl_rep_dsl.sql" | out-file "d:\prg\Reporting\Reports\kl\kl_rep_dsl2.sql" -encoding Unicode
. d:\prg\Reporting\Modules\GodotHelper\GodotHelper.ps1 
##  d:\prg\kl_ps\fs_riport\kl.ps1 
d:
cd \prg\kl_ps\fs_riport

$datumom = Get-Date -format "yyyy_MM_d"

New-Item ".\xls_$datumom" -type directory

Move-Item *.xlsx ".\xls_$datumom"  -force