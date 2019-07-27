#-----------------------------------------------------------------------------------------------------------------------------------------------
#---------------------------------------------------------Script Kódvázának Leírása:------------------------------------------------------------
#-----------------------------------------------------------------------------------------------------------------------------------------------
#A script kódja több fõ részbõl áll, ami látványosan el van választva. Legelõl szerepel két GUI rajzoltató függvény, elsõ a popUp ablak
#generáló fgv, a második a Utasítást tartalmazó ablakot generáló fgv. Ezután következik a Main/fõ rész, amiben elõsször a szükséges adatok
# kinyerése szerepel, majd utánna azok axcel táblába történõ beírása, majd a file mentése, és végül a file elküldése csatolmányként emailben 
# Csefkó Anikónak
#-----------------------------------------------------------------------------------------------------------------------------------------------
#-----------------------------------------------------------------------------------------------------------------------------------------------
#------------------------------------------------------------PopUp ablak Függvény---------------------------------------------------------------
#-----------------------------------------------------------------------------------------------------------------------------------------------
function GeneratePopUp {
#region Import the Assemblies
[reflection.assembly]::loadwithpartialname("System.Windows.Forms") | Out-Null
[reflection.assembly]::loadwithpartialname("System.Drawing") | Out-Null
#endregion

#region Generated Form Objects
$form1 = New-Object System.Windows.Forms.Form
$label1 = New-Object System.Windows.Forms.Label
$PopUpOk = New-Object System.Windows.Forms.Button
$InitialFormWindowState = New-Object System.Windows.Forms.FormWindowState
#endregion Generated Form Objects

#Az 'OK' gomb megnyomása esetén ez fut le: Bezárja a felugro ablakot
$B_PopUpOK= 
{
	$form1.close()
}
$OnLoadForm_StateCorrection=
{
	$form1.WindowState = $InitialFormWindowState
}
#region Generated Form Code
$form1.Name = 'form1'
$form1.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 246
$System_Drawing_Size.Height = 95
$form1.ClientSize = $System_Drawing_Size
$form1.FormBorderStyle = 5

$label1.TabIndex = 1
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 203
$System_Drawing_Size.Height = 53
$label1.Size = $System_Drawing_Size
$label1.Text = 'NEM végezted el a beállítást, Vagy rossz oszlopban tetted meg! Kérlek ellenõrizd!'
$label1.Font = New-Object System.Drawing.Font("Microsoft Sans Serif",9,0,3,0)

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 22
$System_Drawing_Point.Y = 4
$label1.Location = $System_Drawing_Point
$label1.DataBindings.DefaultDataSourceUpdateMode = 0
$label1.Name = 'label1'

$form1.Controls.Add($label1)

$PopUpOk.TabIndex = 0
$PopUpOk.Name = 'PopUpOk'
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 75
$System_Drawing_Size.Height = 23
$PopUpOk.Size = $System_Drawing_Size
$PopUpOk.UseVisualStyleBackColor = $True

$PopUpOk.Text = 'OK'

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 76
$System_Drawing_Point.Y = 60
$PopUpOk.Location = $System_Drawing_Point
$PopUpOk.DataBindings.DefaultDataSourceUpdateMode = 0
$PopUpOk.add_Click($B_PopUpOK)

$form1.Controls.Add($PopUpOk)

#endregion Generated Form Code
$InitialFormWindowState = $form1.WindowState
$form1.add_Load($OnLoadForm_StateCorrection)
$form1.ShowDialog()| Out-Null

} 

#-----------------------------------------------------------------------------------------------------------------------------------------------
#------------------------------------------------------------Utasítás ablak Függvény---------------------------------------------------------------
#-----------------------------------------------------------------------------------------------------------------------------------------------

function GenerateForm {
#region Import the Assemblies
[reflection.assembly]::loadwithpartialname("System.Windows.Forms") | Out-Null
[reflection.assembly]::loadwithpartialname("System.Drawing") | Out-Null
#endregion

#region Generated Form Objects
$Urlap = New-Object System.Windows.Forms.Form
$label6 = New-Object System.Windows.Forms.Label
$label5 = New-Object System.Windows.Forms.Label
$label4 = New-Object System.Windows.Forms.Label
$Cim = New-Object System.Windows.Forms.Label
$leiras = New-Object System.Windows.Forms.Label
$Kilep = New-Object System.Windows.Forms.Button
$OK = New-Object System.Windows.Forms.Button
$InitialFormWindowState = New-Object System.Windows.Forms.FormWindowState
#endregion Generated Form Objects

#'OK' gombra kattintás esemény definiálása: 
$B_OK= 
{
	#megvizsgálja hogy be lett e állítva a cella formátuma text tipusúra
	$tesztszam = "0003"
	$sh.Cells.item(1,3).formulaLocal = $tesztszam
	$teszt = $sh.Cells.item(1,3).formulaLocal
	$teszt
	$tesztszam
	#Ha nem, akkor egy felugró ablakban figyelmeztet. Meghivja  afelugró ablak rajzoló függvényét: GeneratePopUp
	if ($tesztszam -ne $teszt){
		$sh.Cells.item(1,3).formulaLocal = ""
		GeneratePopUp
	}
	#Ha igen, akkor bezárja az ûrlapot, elrejti az excel táblát, és beállítja a $exitChk értékét ugy, hogy a hívás utánni rész lefusson
	if ($tesztszam -eq $teszt){
		$sh.Cells.item(1,3).formulaLocal = ""
		$Urlap.close()
		$excel.Visible = $false
		$exitChk = 0
	}
}
#futás emgszakítása gombra kattintás esemény definiálása: bezárja az excelt, és a folyamatát, majd bezárja
# az ûrlapot is..és beállítja a $exitChk változót 1 re, aminek a fgv hívás utánni rész lefutásában van szerepe
$Exit= 
{
	$excel.quit() 
	if (ps excel) { kill -name excel}
	$null = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
	$exitChk = 1
	$Urlap.close()	
}
$OnLoadForm_StateCorrection=
{
	$Urlap.WindowState = $InitialFormWindowState
}
#ûrlap elemek legenerálása: elemek tulajdonságai
#region Generated Form Code
$Urlap.RightToLeft = 0
$Urlap.Text = 'FONTOS!!! '
$Urlap.Name = 'Urlap'
$Urlap.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 369
$System_Drawing_Size.Height = 266
$Urlap.ClientSize = $System_Drawing_Size

$label6.TabIndex = 6
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 221
$System_Drawing_Size.Height = 37
$label6.Size = $System_Drawing_Size
$label6.Text = 'NE ZÁRJA BE az excel táblát. Az beállítások után kattintson itt az OK gombra.'

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 61
$System_Drawing_Point.Y = 165
$label6.Location = $System_Drawing_Point
$label6.DataBindings.DefaultDataSourceUpdateMode = 0
$label6.Name = 'label6'

$Urlap.Controls.Add($label6)

$label5.TabIndex = 5
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 327
$System_Drawing_Size.Height = 56
$label5.Size = $System_Drawing_Size
$label5.Text = '- A "Number" fülön a "Category" mezõben válassza ki a "Text" formátumot, majd OKézza le!'
$label5.Font = New-Object System.Drawing.Font("Microsoft Sans Serif",9,0,3,1)

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 22
$System_Drawing_Point.Y = 119
$label5.Location = $System_Drawing_Point
$label5.DataBindings.DefaultDataSourceUpdateMode = 0
$label5.Name = 'label5'

$Urlap.Controls.Add($label5)

$label4.TabIndex = 4
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 327
$System_Drawing_Size.Height = 30
$label4.Size = $System_Drawing_Size
$label4.Text = '- Kattintson egér jobb gombjával a kijelölt területre majd válassza a "Format Cells..." alpontot!'
$label4.Font = New-Object System.Drawing.Font("Microsoft Sans Serif",9,0,3,1)

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 22
$System_Drawing_Point.Y = 78
$label4.Location = $System_Drawing_Point
$label4.DataBindings.DefaultDataSourceUpdateMode = 0
$label4.Name = 'label4'

$Urlap.Controls.Add($label4)

$Cim.TabIndex = 3
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 279
$System_Drawing_Size.Height = 23
$Cim.Size = $System_Drawing_Size
$Cim.Text = 'Kérem végezze el a szükséges teendõket: '
$Cim.Font = New-Object System.Drawing.Font("Microsoft Sans Serif",9,1,3,1)

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 13
$System_Drawing_Point.Y = 13
$Cim.Location = $System_Drawing_Point
$Cim.DataBindings.DefaultDataSourceUpdateMode = 0
$Cim.Name = 'Cim'

$Urlap.Controls.Add($Cim)

$leiras.TabIndex = 2
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 327
$System_Drawing_Size.Height = 33
$leiras.Size = $System_Drawing_Size
$leiras.FlatStyle = 1
$leiras.Text = '- A Script által megnyitott excel táblában kérem jelölje ki a 3. oszlopot teljesen!'
$leiras.Font = New-Object System.Drawing.Font("Microsoft Sans Serif",9,0,3,1)

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 22
$System_Drawing_Point.Y = 36
$leiras.Location = $System_Drawing_Point
$leiras.DataBindings.DefaultDataSourceUpdateMode = 0
$leiras.Name = 'leiras'
$leiras.add_Click($handler_label1_Click)

$Urlap.Controls.Add($leiras)

$Kilep.TabIndex = 1
$Kilep.Name = 'Kilep'
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 122
$System_Drawing_Size.Height = 23
$Kilep.Size = $System_Drawing_Size
$Kilep.UseVisualStyleBackColor = $True

$Kilep.Text = 'Futás Megszakítás'

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 227
$System_Drawing_Point.Y = 217
$Kilep.Location = $System_Drawing_Point
$Kilep.DataBindings.DefaultDataSourceUpdateMode = 0
$Kilep.add_Click($Exit)

$Urlap.Controls.Add($Kilep)

$OK.TabIndex = 0
$OK.Name = 'OK'
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 122
$System_Drawing_Size.Height = 23
$OK.Size = $System_Drawing_Size
$OK.UseVisualStyleBackColor = $True

$OK.Text = 'OK'

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 22
$System_Drawing_Point.Y = 217
$OK.Location = $System_Drawing_Point
$OK.DataBindings.DefaultDataSourceUpdateMode = 0
$OK.add_Click($B_OK)

$Urlap.Controls.Add($OK)

#endregion Generated Form Code

$InitialFormWindowState = $Urlap.WindowState
$Urlap.add_Load($OnLoadForm_StateCorrection)
$Urlap.ShowDialog()| Out-Null
return $exitChk
} 

#-----------------------------------------------------------------------------------------------------------------------------------------------
#-----------------------------------------------------------------------------------------------------------------------------------------------
#------------------------------------------------------------MAIN Program-----------------------------------------------------------------------
#-----------------------------------------------------------------------------------------------------------------------------------------------
#-----------------------------------------------------------------------------------------------------------------------------------------------
$excel = New-Object -COM Excel.Application
$ci = [System.Globalization.CultureInfo]'en-US'
#Excell file tárolásának helye, és a file neve, és ha már létezik a könyvtárba, akkor törlése
$src="C:\temp\KFKI_Normal_AD.xls"
if ( [System.IO.File]::Exists($src) )
{
  remove-item -force $src
}
$book = $excel.Workbooks.PSBase.GetType().InvokeMember('Add', [Reflection.BindingFlags]::InvokeMethod, $null,$excel.Workbooks, $null, $ci)
$excel.Visible = $True
$sheet = 1
$sh = $book.sheets.item($sheet)

#Az utasítás leíró GUI rajzoltaó fgv meghívása, és visszatérési értékének átadása egy változónak
$exitChk = GenerateForm
if ($exitChk -eq 0) {
#init:
$j = 0					#felhasználok száma lesz benne. Ciklushoz kell majd ami az excelbe ír.
$univazChk = 0			#akinek nincs univaz kodja annál is kerüljön a tömbbe valami -- > ennek medoldásához kell
#Egyszerübb megközelítés végett oldjuk meg 4 tömbel, ugyanis több dimnziós tömböt, csak fix hosszúságúat definiálhatunk,
#egydimenziósat viszont nem muszály fixen
#A négy adat tárolására szolgáló 4 tömb:
$DisplayN = New-Object system.collections.arraylist		
$SamA = New-Object system.collections.arraylist		
$UnivazKod = New-Object system.collections.arraylist		
$Department = New-Object system.collections.arraylist
$Title = New-Object system.collections.arraylist		
$Manager = New-Object system.collections.arraylist

#------------------------------Adat gyüjtés: megfelelõ adatok tömbökbe gyüjtése, tárolása----------------------------------------------
Get-Mailbox -OrganizationalUnit "kfki.corp/kfki/felhasznalok/normal" | foreach-object `
{ 
 [void] $DisplayN.Add($_.displayname)
	[void] $SamA.Add($_.samaccountname)
	$b = [Microsoft.Exchange.Data.ProxyAddressCollection] $_.emailaddresses
	for ($i=0; $i -ile $b.count; $i++)
		{ 
			if ($b[$i].Addressstring -imatch "HUKF")
				{
					$univaz =$b[$i].Addressstring -replace ("HUKF","") -replace ("@kfkizrt.hu","")
					
					if ($univazChk -eq 0)  
							{
					$univazlista=$univaz
					}
					else
					{
						$univazlista=$univazlista + " és még " + $univaz
					}
					$univazChk = 1
				}
		}
	
		if ($univazChk -eq 1)
		{
				[void] $UnivazKod.Add($univazlista)
				
		}
	
	if ($univazChk -eq 0)
		{
			[void] $UnivazKod.Add("Nincs univaz kód")
		}
	$univazChk = 0
	$univazlista=""
	$j++
}
#4. adat, a department-et külön részben tudjuk kinyerni: (a get-mailbox által visszaadott értékekben nincsen)
$j = 0
Get-Mailbox -OrganizationalUnit "kfki.corp/kfki/felhasznalok/normal" | Get-User | foreach-object `
{ 

	[void] $Department.Add($_.Department)
	[void] $Title.Add($_.Title)
	[void] $Manager.Add($_.manager.name)
	$j++	
}
#--------------------------------------------------------------------------------------------------------------------------------------------
#----------------------------------------------------Excel feltöltése------------------------------------------------------------------------
#--------------------------------------------------------------------------------------------------------------------------------------------

#adatok kiírása az excel file ba:
for ($i=1; $i -le $j; $i++){
	$sh.Cells.item($i,1).formulaLocal = $DisplayN[$i-1]		#oszlop 1
	$sh.Cells.item($i,2).formulaLocal = $SamA[$i-1]			#oszlop 2
	$sh.Cells.item($i,3).formulaLocal = $UnivazKod[$i-1]		#oszlop 3
	$sh.Cells.item($i,4).formulaLocal = $Department[$i-1]		#oszlop 4
	$sh.Cells.item($i,5).formulaLocal = $Title[$i-1]		#oszlop 5
	$sh.Cells.item($i,6).formulaLocal = $Manager[$i-1]		#oszlop 6

}
$j=$j+2

#gyakornokok lekérdezése, és excelbe töltése.
$sh.Cells.item($j,1).formulaLocal = "Gyakornokok:"
$j++
$sh.Cells.item($j,1).formulaLocal = "DisplayName"		
$sh.Cells.item($j,2).formulaLocal = "SamAccountName"	
$sh.Cells.item($j,4).formulaLocal = "Department"		
$sh.Cells.item($j,5).formulaLocal = "Title"
$sh.Cells.item($j,6).formulaLocal = "Manager"
Get-User -OrganizationalUnit "kfki.corp/kfki/felhasznalok/Gyakornokok" | foreach-object `
{ 

	$j++	
	
	$sh.Cells.item($j,1).formulaLocal = $_.displayname			#oszlop 1
	$sh.Cells.item($j,2).formulaLocal = $_.samaccountname			#oszlop 2
	$sh.Cells.item($j,4).formulaLocal = $_.Department			#oszlop 4
	$sh.Cells.item($j,5).formulaLocal = $_.Title			#oszlop 5
	$sh.Cells.item($j,6).formulaLocal = $_.Manager.name			#oszlop 5
	
}

#excel elmentése $src helyre és néven "56" formátumban - excel formátum, majd futo excel folyamat bezárása
[void]$book.PSBase.GetType().InvokeMember('SaveAs', [Reflection.BindingFlags]::InvokeMethod, $null, $book, ($src,"56"), $ci)
$excel.quit() 
if (ps excel) { kill -name excel}
$null = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)

#--------------------------------------------------------------------------------------------------------------------------------------------
#----------------------------------------------File Csatolás, és Email küldés Csefkonak------------------------------------------------------
#--------------------------------------------------------------------------------------------------------------------------------------------

$sender = "olah.zoltan@kfkizrt.hu"
#$recipient = "olah.zoltan@kfkizrt.hu"
#$recipient = "olah.zoltan@kfkizrt.hu,smalekker.szilvia@kfkizrt.hu"
$recipient = "olah.zoltan@kfkizrt.hu,csefko.aniko@kfkizrt.hu"

$server = "k-mail3hc"
$subject = "AD lista - " + [System.DateTime]::Now
$body = "Szia!

Küldöm csatolva újra az AD listát. 


Üdv, Zoli
"
$msg = new-object System.Net.Mail.MailMessage $sender, $recipient, $subject, $body
$attachment = new-object System.Net.Mail.Attachment $src
$msg.Attachments.Add($attachment)
$client = new-object System.Net.Mail.SmtpClient $server
$client.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials
$client.Send($msg)
}

