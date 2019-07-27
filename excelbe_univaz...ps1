#--------------------------------------------------------------------------------------------------------------------------------------------
#22. oszlopbanvan a samaccountname.(V-oszlop); 2.sorban a fejléc, adatok a 3.sortol.
#337. oszlopban van az Univaz kod. (LY-oszlop); 2.sorban a fejléc, adatok a 3.sortol legyenek.
#összes név amire kell: 569: 3-tól 571-ig
#---FONTOS: Ha számot szeretnék excelbe irni, és 0 van az elején, akkor levágja: állitsuk elöbb át a mezöket szöveg formátumúra manuálisan---
#--------------------------------------------------------------------------------------------------------------------------------------------

$excel = New-Object -COM Excel.Application
$ci = [System.Globalization.CultureInfo]'en-US'
#-----------------------------------------forrás excel file elérési utvonala-----------------------------------------------------------------
$src="I:\munka\Feladat_12_IQSYSusers\AD_User_List.xlsx"
$book = $excel.Workbooks.PSBase.GetType().InvokeMember('Open', [Reflection.BindingFlags]::InvokeMethod, $null,$excel.Workbooks, $src, $ci)
#$excel.Visible = $True
$sheet = 1
$sh = $book.sheets.item($sheet)

#2 statikus ostzlop valtozo kell, a kettö adat számára. beolvas elsöböl, kiír a másodikba. Egy szabad változó tárolja az akt adatot.
#$col1 = 22	  #sumaccauntname
$col1 = 28   #mailnickname
$col2 = 337

$sh.Cells.item(2,337).formulaLocal = "Univaz-kód"
#a ciklus a sorokat viszi: 3tól 571 ig. 
for ($i=3; $i -le 571; $i++){
	$samname= $sh.Cells.item($i,$col1).formulaLocal			#$samname be kiolvassulk az excelböl:jelenleg a mailnickname-t, ami alapján keressük az univazt
	#csak akkor lépünk be a keresési részbe, ha van tényleges név..szoval ha nem üres mezöt olvastunk ki az excelböl
	if($samname -ine ""){
		$samname			#név kiírás csak tesztre
		#Maga az Univaz kinyerés:
		Get-Mailbox -Identity $samname | foreach-object `
		{ 
			$b = [Microsoft.Exchange.Data.ProxyAddressCollection] $_.emailaddresses
			for ($j=0; $j -ile $b.count; $j++)
			{ 
				if ($b[$j].Addressstring -imatch "HUIQ")
				{
					[string] $univaz =[string] $b[$j].Addressstring -replace ("HUIQ","") -replace ("@iqsys.hu","")	
					$sh.Cells.item($i,$col2).formulaLocal = [string] $univaz
				}
			}
		}
	}
}
#excel változásainak elmentése, majd kilépés, bezárás
[void]$book.PSBase.GetType().InvokeMember('Save', [Reflection.BindingFlags]::InvokeMethod, $null, $book, $null, $ci)
$excel.quit() 
if (ps excel) { kill -name excel}
$null = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)



