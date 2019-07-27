#-----------------------------------------------------------------------------------------------------------------------------------------------
#---------------------------------------------------------Script K�dv�z�nak Le�r�sa:------------------------------------------------------------
#-----------------------------------------------------------------------------------------------------------------------------------------------
#A script k�dja t�bb f� r�szb�l �ll, ami l�tv�nyosan el van v�lasztva. Legel�l szerepel k�t GUI rajzoltat� f�ggv�ny, els� a popUp ablak
#gener�l� fgv, a m�sodik a Utas�t�st tartalmaz� ablakot gener�l� fgv. Ezut�n k�vetkezik a Main/f� r�sz, amiben el�ssz�r a sz�ks�ges adatok
# kinyer�se szerepel, majd ut�nna azok axcel t�bl�ba t�rt�n� be�r�sa, majd a file ment�se, �s v�g�l a file elk�ld�se csatolm�nyk�nt emailben 
# Csefk� Anik�nak
#-----------------------------------------------------------------------------------------------------------------------------------------------
#-----------------------------------------------------------------------------------------------------------------------------------------------
#------------------------------------------------------------PopUp ablak F�ggv�ny---------------------------------------------------------------
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

#Az 'OK' gomb megnyom�sa eset�n ez fut le: Bez�rja a felugro ablakot
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
$label1.Text = 'NEM v�gezted el a be�ll�t�st, Vagy rossz oszlopban tetted meg! K�rlek ellen�rizd!'
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
#------------------------------------------------------------Utas�t�s ablak F�ggv�ny---------------------------------------------------------------
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

#'OK' gombra kattint�s esem�ny defini�l�sa: 
$B_OK= 
{
	#megvizsg�lja hogy be lett e �ll�tva a cella form�tuma text tipus�ra
	$tesztszam = "0003"
	$sh.Cells.item(1,3).formulaLocal = $tesztszam
	$teszt = $sh.Cells.item(1,3).formulaLocal
	$teszt
	$tesztszam
	#Ha nem, akkor egy felugr� ablakban figyelmeztet. Meghivja  afelugr� ablak rajzol� f�ggv�ny�t: GeneratePopUp
	if ($tesztszam -ne $teszt){
		$sh.Cells.item(1,3).formulaLocal = ""
		GeneratePopUp
	}
	#Ha igen, akkor bez�rja az �rlapot, elrejti az excel t�bl�t, �s be�ll�tja a $exitChk �rt�k�t ugy, hogy a h�v�s ut�nni r�sz lefusson
	if ($tesztszam -eq $teszt){
		$sh.Cells.item(1,3).formulaLocal = ""
		$Urlap.close()
		$excel.Visible = $false
		$exitChk = 0
	}
}
#fut�s emgszak�t�sa gombra kattint�s esem�ny defini�l�sa: bez�rja az excelt, �s a folyamat�t, majd bez�rja
# az �rlapot is..�s be�ll�tja a $exitChk v�ltoz�t 1 re, aminek a fgv h�v�s ut�nni r�sz lefut�s�ban van szerepe
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
#�rlap elemek legener�l�sa: elemek tulajdons�gai
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
$label6.Text = 'NE Z�RJA BE az excel t�bl�t. Az be�ll�t�sok ut�n kattintson itt az OK gombra.'

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
$label5.Text = '- A "Number" f�l�n a "Category" mez�ben v�lassza ki a "Text" form�tumot, majd OK�zza le!'
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
$label4.Text = '- Kattintson eg�r jobb gombj�val a kijel�lt ter�letre majd v�lassza a "Format Cells..." alpontot!'
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
$Cim.Text = 'K�rem v�gezze el a sz�ks�ges teend�ket: '
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
$leiras.Text = '- A Script �ltal megnyitott excel t�bl�ban k�rem jel�lje ki a 3. oszlopot teljesen!'
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

$Kilep.Text = 'Fut�s Megszak�t�s'

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
#Excell file t�rol�s�nak helye, �s a file neve, �s ha m�r l�tezik a k�nyvt�rba, akkor t�rl�se
$src="C:\temp\KFKI_Normal_AD.xls"
if ( [System.IO.File]::Exists($src) )
{
  remove-item -force $src
}
$book = $excel.Workbooks.PSBase.GetType().InvokeMember('Add', [Reflection.BindingFlags]::InvokeMethod, $null,$excel.Workbooks, $null, $ci)
$excel.Visible = $True
$sheet = 1
$sh = $book.sheets.item($sheet)

#Az utas�t�s le�r� GUI rajzolta� fgv megh�v�sa, �s visszat�r�si �rt�k�nek �tad�sa egy v�ltoz�nak
$exitChk = GenerateForm
if ($exitChk -eq 0) {
#init:
$j = 0					#felhaszn�lok sz�ma lesz benne. Ciklushoz kell majd ami az excelbe �r.
$univazChk = 0			#akinek nincs univaz kodja ann�l is ker�lj�n a t�mbbe valami -- > ennek medold�s�hoz kell
#Egyszer�bb megk�zel�t�s v�gett oldjuk meg 4 t�mbel, ugyanis t�bb dimnzi�s t�mb�t, csak fix hossz�s�g�at defini�lhatunk,
#egydimenzi�sat viszont nem musz�ly fixen
#A n�gy adat t�rol�s�ra szolg�l� 4 t�mb:
$DisplayN = New-Object system.collections.arraylist		
$SamA = New-Object system.collections.arraylist		
$UnivazKod = New-Object system.collections.arraylist		
$Department = New-Object system.collections.arraylist
$Title = New-Object system.collections.arraylist		
$Manager = New-Object system.collections.arraylist

#------------------------------Adat gy�jt�s: megfelel� adatok t�mb�kbe gy�jt�se, t�rol�sa----------------------------------------------
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
						$univazlista=$univazlista + " �s m�g " + $univaz
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
			[void] $UnivazKod.Add("Nincs univaz k�d")
		}
	$univazChk = 0
	$univazlista=""
	$j++
}
#4. adat, a department-et k�l�n r�szben tudjuk kinyerni: (a get-mailbox �ltal visszaadott �rt�kekben nincsen)
$j = 0
Get-Mailbox -OrganizationalUnit "kfki.corp/kfki/felhasznalok/normal" | Get-User | foreach-object `
{ 

	[void] $Department.Add($_.Department)
	[void] $Title.Add($_.Title)
	[void] $Manager.Add($_.manager.name)
	$j++	
}
#--------------------------------------------------------------------------------------------------------------------------------------------
#----------------------------------------------------Excel felt�lt�se------------------------------------------------------------------------
#--------------------------------------------------------------------------------------------------------------------------------------------

#adatok ki�r�sa az excel file ba:
for ($i=1; $i -le $j; $i++){
	$sh.Cells.item($i,1).formulaLocal = $DisplayN[$i-1]		#oszlop 1
	$sh.Cells.item($i,2).formulaLocal = $SamA[$i-1]			#oszlop 2
	$sh.Cells.item($i,3).formulaLocal = $UnivazKod[$i-1]		#oszlop 3
	$sh.Cells.item($i,4).formulaLocal = $Department[$i-1]		#oszlop 4
	$sh.Cells.item($i,5).formulaLocal = $Title[$i-1]		#oszlop 5
	$sh.Cells.item($i,6).formulaLocal = $Manager[$i-1]		#oszlop 6

}
$j=$j+2

#gyakornokok lek�rdez�se, �s excelbe t�lt�se.
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

#excel elment�se $src helyre �s n�ven "56" form�tumban - excel form�tum, majd futo excel folyamat bez�r�sa
[void]$book.PSBase.GetType().InvokeMember('SaveAs', [Reflection.BindingFlags]::InvokeMethod, $null, $book, ($src,"56"), $ci)
$excel.quit() 
if (ps excel) { kill -name excel}
$null = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)

#--------------------------------------------------------------------------------------------------------------------------------------------
#----------------------------------------------File Csatol�s, �s Email k�ld�s Csefkonak------------------------------------------------------
#--------------------------------------------------------------------------------------------------------------------------------------------

$sender = "olah.zoltan@kfkizrt.hu"
#$recipient = "olah.zoltan@kfkizrt.hu"
#$recipient = "olah.zoltan@kfkizrt.hu,smalekker.szilvia@kfkizrt.hu"
$recipient = "olah.zoltan@kfkizrt.hu,csefko.aniko@kfkizrt.hu"

$server = "k-mail3hc"
$subject = "AD lista - " + [System.DateTime]::Now
$body = "Szia!

K�ld�m csatolva �jra az AD list�t. 


�dv, Zoli
"
$msg = new-object System.Net.Mail.MailMessage $sender, $recipient, $subject, $body
$attachment = new-object System.Net.Mail.Attachment $src
$msg.Attachments.Add($attachment)
$client = new-object System.Net.Mail.SmtpClient $server
$client.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials
$client.Send($msg)
}

