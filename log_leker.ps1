# AD (Active Directory) / LDAP userek lekérdezése ps (Powershell script) segítségével
# Meghívása : powershell.exe .\ldap.ps1 
#-------------------------------------------------------------------------------------
#Alap könyvtár beállítás
CD d:\prg\kl_ps\jelenleti

# Futtatási dátum kikérése (filenév vége ezt tartalmazza) ha esetleg létezik már a file felülírja
$date = ( get-date ).ToString('yyyyMMdd')
$file = New-Item -type file "d:\prg\kl_ps\jelenleti\kl_munkaido_$date.log" -Force

#Feldolgozás indítás leállítás adatok kivétele a logból az eredmény filebe írányításával
Get-WinEvent -FilterHashtable @{logname='system';id=89,109;StartTime="2012/09/01";EndTime="2015/01/01"} | Format-Table   -Autosize | out-file $file
