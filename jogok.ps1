$excel = New-Object -COM Excel.Application
$ci = [System.Globalization.CultureInfo]'en-US'
#-----------------------------------------forrás excel file elérési utvonala---------------------------------------------------------------
$src="I:\munka\Feladat_7_jogok\KFKI_uj.xls"
$book = $excel.Workbooks.PSBase.GetType().InvokeMember('Open', [Reflection.BindingFlags]::InvokeMethod, $null,$excel.Workbooks, $src, $ci)
#$excel.Visible = $True
$sheet = 1
$sh = $book.sheets.item($sheet)

#-------------------A tömb ahova a cella tartalmakat beolvassuk--------------------------
# itt adhatjuk meg hogy mettöl meddig olvassuk. --$row-- sortól --$rto-- sorig, és --$col-- oszloptól --$cto-- oszlopig számmal megadva.
# a tömb méretét szintén ehhez igazítva: rendre: hány sor, hány oszlop
$adat = New-Object 'string[,]' 1,29
$row = 2
$rto = 2
$col = 14
$cto = 42
$sor = $rto-$row+1
$oszlop = $cto-$col+1
for ($i=0; $i -lt $sor; $i++){
	for($j=0; $j -lt $oszlop; $j++){
		#a cellák beolvasása
		$adat[$i,$j] = [string] $sh.Cells.item($row,$col).formulaLocal
		$row
		$col
		#tesztként kiírás
		$sh.Cells.item($row,$col).formulaLocal
		$col++
	}
	$row++
	$col = 2
}
#kilépés, excel folyamat leállítása, és com object release
$excel.quit() 
if (ps excel) { kill -name excel}
$null = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)

#Atolso cella adat beállítása: Belépés datum date format legyen.
$temp = Get-Date "01/01/1900"
$temp2 = ([long]$adat[0,28]) - 2
$temp = $temp.addDays($temp2)
$temp = $temp.toShortDateString()
$adat[0,28] =  $temp

# teszt kiírás
$adat | ft

#név kisbetűsítése
$veznev = $adat[0,5].tolower()
$kereszt = $adat[0,3].tolower()

#ékezetes betűk eltávolítása
$veznev = $veznev.Replace("á","a")
$veznev = $veznev.Replace("é","e")
$veznev = $veznev.Replace("í","i")
$veznev = $veznev.Replace("ó","o")
$veznev = $veznev.Replace("ö","o")
$veznev = $veznev.Replace("ő","o")
$veznev = $veznev.Replace("ú","u")
$veznev = $veznev.Replace("ü","u")
$veznev = $veznev.Replace("ű","u")
$veznev = $veznev.Replace(" ","")

$kereszt = $kereszt.Replace("á","a")
$kereszt = $kereszt.Replace("é","e")
$kereszt = $kereszt.Replace("í","i")
$kereszt = $kereszt.Replace("ó","o")
$kereszt = $kereszt.Replace("ö","o")
$kereszt = $kereszt.Replace("ő","o")
$kereszt = $kereszt.Replace("ú","u")
$kereszt = $kereszt.Replace("ü","u")
$kereszt = $kereszt.Replace("ű","u")


