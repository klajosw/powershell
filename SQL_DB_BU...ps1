#============================================================ 
# Backup a Database using PowerShell and SQL Server SMO 
# Script below creates a full backup 
# Donabel Santos 
#============================================================ 
 
 
#specify database to backup 
#ideally this will be an argument you pass in when you run 
#this script, but let's simplify for now 
$dbToBackup = "OPRAKTAR"
 
 
#clear screen 
cls
 
 
#load assemblies 
#note need to load SqlServer.SmoExtended to use SMO backup in SQL Server 2008 
#otherwise may get this error 
#Cannot find type [Microsoft.SqlServer.Management.Smo.Backup]: make sure 
#the assembly containing this type is loaded. 
 
 
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SqlServer.SMO") | Out-Null
#Need SmoExtended for smo.backup 
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SqlServer.SmoExtended") | Out-Null
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SqlServer.ConnectionInfo") | Out-Null
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SqlServer.SmoEnum") | Out-Null
 
#create a new server object 
$server = New-Object ("Microsoft.SqlServer.Management.Smo.Server") "(local)"
$backupDirectory = $server.Settings.BackupDirectory 
 
#display default backup directory 
"Default Backup Directory: " + $backupDirectory
 
 
$db = $server.Databases[$dbToBackup] 
$dbName = $db.Name 
 
 
$timestamp = Get-Date -format yyyyMMddHHmmss 
$smoBackup = New-Object ("Microsoft.SqlServer.Management.Smo.Backup") 
 
 
#BackupActionType specifies the type of backup. 
#Options are Database, Files, Log 
#This belongs in Microsoft.SqlServer.SmoExtended assembly 
 
 
$smoBackup.Action = "Database"
$smoBackup.BackupSetDescription = "Full Backup of " + $dbName
$smoBackup.BackupSetName = $dbName + " Backup"
$smoBackup.Database = $dbName
$smoBackup.MediaDescription = "Disk"
$smoBackup.Devices.AddDevice($backupDirectory + "\" + $dbName + "_" + $timestamp + ".bak", "File") 
$smoBackup.SqlBackup($server) 
 
 
#let's confirm, let's list list all backup files 
$directory = Get-ChildItem $backupDirectory
 
 
#list only files that end in .bak, assuming this is your convention for all backup files 
$backupFilesList = $directory | where {$_.extension -eq ".bak"} 
$backupFilesList | Format-Table Name, LastWriteTime


$ma = Get-Date			
$ma = $ma.Date
dir "D:\MSSQL10_50.MSSQLSERVER\MSSQL\Backup\*.bak"  | ForEach-Object `
{
#file létrehozás dátumának meghatározása
		$file = $_
		$create =[datetime] $file.CreationTime
		$create = $create.Date
		$name = $file.name.Replace(".kfkirestore","")
		$hatra = ($ma - $create).days
		#email a felhasználonak naponta a 7.napig
		if ($hatra -gt 10){
			$path = "D:\MSSQL10_50.MSSQLSERVER\MSSQL\Backup\" + $file
			Remove-Item $path -Recurse
		}
}


