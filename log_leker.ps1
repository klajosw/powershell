# AD (Active Directory) / LDAP userek lek�rdez�se ps (Powershell script) seg�ts�g�vel
# Megh�v�sa : powershell.exe .\ldap.ps1 
#-------------------------------------------------------------------------------------
#Alap k�nyvt�r be�ll�t�s
CD d:\prg\kl_ps\jelenleti

# Futtat�si d�tum kik�r�se (filen�v v�ge ezt tartalmazza) ha esetleg l�tezik m�r a file fel�l�rja
$date = ( get-date ).ToString('yyyyMMdd')
$file = New-Item -type file "d:\prg\kl_ps\jelenleti\kl_munkaido_$date.log" -Force

#Feldolgoz�s ind�t�s le�ll�t�s adatok kiv�tele a logb�l az eredm�ny filebe �r�ny�t�s�val
Get-WinEvent -FilterHashtable @{logname='system';id=89,109;StartTime="2012/09/01";EndTime="2015/01/01"} | Format-Table   -Autosize | out-file $file
