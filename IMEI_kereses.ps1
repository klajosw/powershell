$excel = New-Object -COM Excel.Application
$ci = [System.Globalization.CultureInfo]'en-US'
#Excell file tárolásának helye, és a file neve, és ha már létezik a könyvtárba, akkor törlése
$src="d:\IMEI.xls"
if ( [System.IO.File]::Exists($src) )
{
  remove-item -force $src
}
$book = $excel.Workbooks.PSBase.GetType().InvokeMember('Add', [Reflection.BindingFlags]::InvokeMethod, $null,$excel.Workbooks, $null, $ci)
$excel.Visible = $True
$sheet = 1
$sh = $book.sheets.item($sheet)
		
$sh.Cells.item(1,1).formulaLocal = "Neve"		#oszlop 1
$sh.Cells.item(1,2).formulaLocal = "DeviceID"			#oszlop 2
$sh.Cells.item(1,3).formulaLocal = "DeviceIMEI"		#oszlop 3
$i = 2
Get-mailbox -OrganizationalUnit  "kfki.corp/kfki/felhasznalok/normal"| foreach-object `
{ 
	Get-ActiveSyncDeviceStatistics -Mailbox $_ |  foreach-object `
	{ 
		$a = $_.identity
		$name = $a.SmtpAddress -replace ("@kfkizrt.hu","")	
		"Neve: " + $name
		"DeviceID: " + $_.DeviceID
		"DeviceIMEI: " + $_.DeviceIMEI
		$sh.Cells.item($i,1).formulaLocal = $name		#oszlop 1
		$sh.Cells.item($i,2).formulaLocal = $_.DeviceID			#oszlop 2
		$sh.Cells.item($i,3).formulaLocal = $_.DeviceIMEI		#oszlop 3
		$i++
	}
}


#excel elmentése $src helyre és néven "56" formátumban - excel formátum, majd futo excel folyamat bezárása
[void]$book.PSBase.GetType().InvokeMember('SaveAs', [Reflection.BindingFlags]::InvokeMethod, $null, $book, ($src,"56"), $ci)
$excel.quit() 
if (ps excel) { kill -name excel}
$null = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)

