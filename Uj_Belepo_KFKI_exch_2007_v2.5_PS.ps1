param($p1, $p2, $p3, $p4, $p5, $p6, $p7, $p8, $p9, $p10, $p11, $p12, $p13, $p14, $p15, $p16, $p17)

set-ExecutionPolicy remotesigned
Add-PSSnapin Quest.ActiveRoles.ADManagement
Add-PSSnapin Microsoft.Exchange.Management.PowerShell.Admin

echo "eeekkezdtem" >> c:\Temp\MAMA.log
#================================================================================================
#================================================================================================
#=================================  FÜGGVÉNYEK helye  ===========================================
#================================================================================================
# Itt szerepel a név ellenörző GUI kodja 
#		- futás elött ellenörizhetjük, valoban jo nevet irtunk e be az Excelbe
#================================================================================================

#Generated Form Function
function GenerateForm {
#region Import the Assemblies
[reflection.assembly]::loadwithpartialname("System.Windows.Forms") | Out-Null
[reflection.assembly]::loadwithpartialname("System.Drawing") | Out-Null
#endregion

#region Generated Form Objects
$form1 = New-Object System.Windows.Forms.Form
$L_szoveg2 = New-Object System.Windows.Forms.Label
$L_Name = New-Object System.Windows.Forms.Label
$L_szoveg = New-Object System.Windows.Forms.Label
$B_Exit = New-Object System.Windows.Forms.Button
$B_OK = New-Object System.Windows.Forms.Button
$InitialFormWindowState = New-Object System.Windows.Forms.FormWindowState
#endregion Generated Form Objects

#----------------------------------------------
#Generated Event Script Blocks
#----------------------------------------------
#Provide Custom Code for events specified in PrimalForms.
$Accept= 
{
	$form1.close()
	$mehet = 1
}

$Nemjo= 
{
	$form1.close()
	$mehet = 0
}

$OnLoadForm_StateCorrection=
{#Correct the initial state of the form to prevent the .Net maximized form issue
	$form1.WindowState = $InitialFormWindowState
}

#----------------------------------------------
#region Generated Form Code
$form1.Name = 'form1'
$form1.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 298
$System_Drawing_Size.Height = 176
$form1.ClientSize = $System_Drawing_Size

$L_szoveg2.TabIndex = 4
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 273
$System_Drawing_Size.Height = 33
$L_szoveg2.Size = $System_Drawing_Size
$L_szoveg2.Text = 'Biztosan őt szeretnéd beléptetni? Ha igen, klikkelj az "OK"-ra! Különben lépj ki és nézd meg az xls-t!'

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 13
$System_Drawing_Point.Y = 89
$L_szoveg2.Location = $System_Drawing_Point
$L_szoveg2.DataBindings.DefaultDataSourceUpdateMode = 0
$L_szoveg2.Name = 'L_szoveg2'
$L_szoveg2.add_Click($handler_label3_Click)

$form1.Controls.Add($L_szoveg2)

$L_Name.TabIndex = 3
$L_Name.BackColor = [System.Drawing.Color]::FromArgb(255,255,255,255)
$L_Name.TextAlign = 16
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 173
$System_Drawing_Size.Height = 23
$L_Name.Size = $System_Drawing_Size
$L_Name.text = $displayname

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 60
$System_Drawing_Point.Y = 56
$L_Name.Location = $System_Drawing_Point
$L_Name.DataBindings.DefaultDataSourceUpdateMode = 0
$L_Name.Name = 'L_Name'

$form1.Controls.Add($L_Name)

$L_szoveg.TabIndex = 2
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 285
$System_Drawing_Size.Height = 37
$L_szoveg.Size = $System_Drawing_Size
$L_szoveg.Text = 'Ezt az embert találtam a \\KFKI.CORP\KFKI\mukodes\HR\KFKI\KFKI_uj.xls-ben:'

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 12
$System_Drawing_Point.Y = 19
$L_szoveg.Location = $System_Drawing_Point
$L_szoveg.DataBindings.DefaultDataSourceUpdateMode = 0
$L_szoveg.Name = 'L_szoveg'

$form1.Controls.Add($L_szoveg)

$B_Exit.TabIndex = 1
$B_Exit.Name = 'B_Exit'
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 75
$System_Drawing_Size.Height = 23
$B_Exit.Size = $System_Drawing_Size
$B_Exit.UseVisualStyleBackColor = $True

$B_Exit.Text = 'Kilépés'

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 202
$System_Drawing_Point.Y = 134
$B_Exit.Location = $System_Drawing_Point
$B_Exit.DataBindings.DefaultDataSourceUpdateMode = 0
$B_Exit.add_Click($Nemjo)

$form1.Controls.Add($B_Exit)

$B_OK.TabIndex = 0
$B_OK.Name = 'B_OK'
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 75
$System_Drawing_Size.Height = 23
$B_OK.Size = $System_Drawing_Size
$B_OK.UseVisualStyleBackColor = $True

$B_OK.Text = 'OK'

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 23
$System_Drawing_Point.Y = 134
$B_OK.Location = $System_Drawing_Point
$B_OK.DataBindings.DefaultDataSourceUpdateMode = 0
$B_OK.add_Click($Accept)

$form1.Controls.Add($B_OK)

#endregion Generated Form Code

#Save the initial state of the form
$InitialFormWindowState = $form1.WindowState
#Init the OnLoad event to correct the initial state of the form
$form1.add_Load($OnLoadForm_StateCorrection)
#Show the Form
$form1.ShowDialog()| Out-Null

Return $mehet
} 

function remove_spaces($string){
	$string = $string.TrimEnd()
	$string = $string.TrimStart()
	return $String
}


#================================================================================================
#================================================================================================
#=====================================  MAIN Progam  ============================================
#================================================================================================
#================================================================================================


#log könytár létezésének ellenőrzése, ha nincs akkor létrehozás
If ((Test-Path "c:\Temp") -eq $false ){
	New-Item -Path "C:\" -Name "Temp" -type directory
}
#logolás
echo $datum > c:\Temp\kfki_newuser_v2.5.log
echo "                               Uj felhasználó felvételének elkezdése." >> c:\Temp\kfki_newuser_v2.5.log
echo "" >> c:\Temp\kfki_newuser_v2.5.log
echo "¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤" >> c:\Temp\kfki_newuser_v2.5.log
echo "0. pont végrahajtása: Adatok beolvasása az Űrlapról, majd alakítása, és származtatott adatok generálása" >> c:\Temp\kfki_newuser_v2.5.log




$mailboxdatabase = remove_spaces($p1)
$organizationalunit = remove_spaces($p2)
If ($p3 -eq "-"){
	$initials = "" #remove_spaces($p3)
}else {
	$initials = remove_spaces($p3)
}

$lastname = remove_spaces($p5)
$firstname = remove_spaces($p4)
$telephonenumber = "" #remove_spaces($p6)
$mobile = "" #remove_spaces($p7)
$title = remove_spaces($p8)
$department = remove_spaces($p9)
$flcsoport = remove_spaces($p10)
$manager = remove_spaces($p11)
$physicaldeliveryofficename = remove_spaces($p12)
$csoporttagsag2 = remove_spaces($p14)
$csoporttagsag3 = remove_spaces($p13)
$csoporttagsag4 = remove_spaces($p15)
$belepes_datuma = remove_spaces($p16)
$koltseg_hely = remove_spaces($p17)


#Belelpes datum beállítása: Date format legyen.
#$temp = Get-Date "01/01/1900"
#$temp2 = ([long]$belepes_datuma) - 2
#$temp = $temp.addDays($temp2)
#$temp = $temp.toShortDateString()
#$belepes_datuma =  $temp

If ($initials -eq "") {
	$displayname = $lastname + " " + $firstname
}
Else{
	$displayname = $initials + " " + $lastname + " " + $firstname
}

#------------------------------------- Ékezetek kivétele -----------------------------------------------
$samaccountnamebeta = $lastname + $firstname
$samaccountnamebeta2= $samaccountnamebeta.tolower()

$sam_without_ekezet1 = $samaccountnamebeta2.Replace("á","a")
$sam_without_ekezet2 = $sam_without_ekezet1.Replace("é","e")
$sam_without_ekezet3 = $sam_without_ekezet2.Replace("í","i")
$sam_without_ekezet4 = $sam_without_ekezet3.Replace("ó","o")
$sam_without_ekezet5 = $sam_without_ekezet4.Replace("ö","o")
$sam_without_ekezet6 = $sam_without_ekezet5.Replace("ő","o")
$sam_without_ekezet7 = $sam_without_ekezet6.Replace("ú","u")
$sam_without_ekezet8 = $sam_without_ekezet7.Replace("ü","u")
$sam_without_ekezet9 = $sam_without_ekezet8.Replace("ű","u")
$sam_without_ekezet10 = $sam_without_ekezet9.Replace(" ","")
$sam_without_ekezet = $sam_without_ekezet10
$samaccountname = $sam_without_ekezet

#--------------------------------------------------------------------------------------------------------

$company = "KFKI_ZRT"
#$homedirectory = "\\k-file1\" + $samaccountname + "$"
$homedrive = "i:"
$logonscript = "kfkizrt-gen.cmd"
$csoporttagsag1 = ""
if ($csoporttagsag3 -eq "X") { $csoporttagsag1="X" }
if ($csoporttagsag4 -eq "X") { $csoporttagsag1="X" }
#---------------------------------------------------------------------------------------------------------

#-----------------------------------Ha ezeket nem toltik ki az excel-ben akkor kinullazzuk----------------
if ($telephonenumber -eq "") { $telephonenumber = " " }
if ($mobile -eq "") { $mobile = " " }
#---------------------------------------------------------------------------------------------------------

#-----------------------------------Kiszedjuk az ekezeteket a distinghuised name-bol, mert a mailbox generalas csak igy megy
$displayname2 = $displayname.tolower()
$userdn_without_ekezet1 = $displayname2.Replace("á","a")
$userdn_without_ekezet2 = $userdn_without_ekezet1.Replace("é","e")
$userdn_without_ekezet3 = $userdn_without_ekezet2.Replace("í","i")
$userdn_without_ekezet4 = $userdn_without_ekezet3.Replace("ó","o")
$userdn_without_ekezet5 = $userdn_without_ekezet4.Replace("ö","o")
$userdn_without_ekezet6 = $userdn_without_ekezet5.Replace("ő","o")
$userdn_without_ekezet7 = $userdn_without_ekezet6.Replace("ú","u")
$userdn_without_ekezet8 = $userdn_without_ekezet7.Replace("ü","u")
$userdn_without_ekezet9 = $userdn_without_ekezet8.Replace("ű","u")
$userdn_without_ekezet10 = $userdn_without_ekezet9.Replace(" ","")
$userdn_without_ekezet = $userdn_without_ekezet10
#---------------------------------------------------------------------------------------------------------

#------------------------------------		Premier partner gyartasa		------------------------------
$lastname2 = $lastname.tolower()
$lastname_without_ekezet1 = $lastname2.Replace("á","a")
$lastname_without_ekezet2 = $lastname_without_ekezet1.Replace("é","e")
$lastname_without_ekezet3 = $lastname_without_ekezet2.Replace("í","i")
$lastname_without_ekezet4 = $lastname_without_ekezet3.Replace("ó","o")
$lastname_without_ekezet5 = $lastname_without_ekezet4.Replace("ö","o")
$lastname_without_ekezet6 = $lastname_without_ekezet5.Replace("ő","o")
$lastname_without_ekezet7 = $lastname_without_ekezet6.Replace("ú","u")
$lastname_without_ekezet8 = $lastname_without_ekezet7.Replace("ü","u")
$lastname_without_ekezet9 = $lastname_without_ekezet8.Replace("ű","u")
$lastname_without_ekezet10 = $lastname_without_ekezet9.Replace(" ","")
$lastname_without_ekezet = $lastname_without_ekezet10

$firstname2 = $firstname.tolower()
$firstname_without_ekezet1 = $firstname2.Replace("á","a")
$firstname_without_ekezet2 = $firstname_without_ekezet1.Replace("é","e")
$firstname_without_ekezet3 = $firstname_without_ekezet2.Replace("í","i")
$firstname_without_ekezet4 = $firstname_without_ekezet3.Replace("ó","o")
$firstname_without_ekezet5 = $firstname_without_ekezet4.Replace("ö","o")
$firstname_without_ekezet6 = $firstname_without_ekezet5.Replace("ő","o")
$firstname_without_ekezet7 = $firstname_without_ekezet6.Replace("ú","u")
$firstname_without_ekezet8 = $firstname_without_ekezet7.Replace("ü","u")
$firstname_without_ekezet9 = $firstname_without_ekezet8.Replace("ű","u")
$firstname_without_ekezet10 = $firstname_without_ekezet9.Replace(" ","")
$firstname_without_ekezet = $firstname_without_ekezet10

$initials2 = $initials.tolower()
$initials_without_ekezet1 = $initials2.Replace("á","a")
$initials_without_ekezet2 = $initials_without_ekezet1.Replace("é","e")
$initials_without_ekezet3 = $initials_without_ekezet2.Replace("í","i")
$initials_without_ekezet4 = $initials_without_ekezet3.Replace("ó","o")
$initials_without_ekezet5 = $initials_without_ekezet4.Replace("ö","o")
$initials_without_ekezet6 = $initials_without_ekezet5.Replace("ő","o")
$initials_without_ekezet7 = $initials_without_ekezet6.Replace("ú","u")
$initials_without_ekezet8 = $initials_without_ekezet7.Replace("ü","u")
$initials_without_ekezet9 = $initials_without_ekezet8.Replace("ű","u")
$initials_without_ekezet10 = $initials_without_ekezet9.Replace(" ","")
$initials_without_ekezet = $initials_without_ekezet10



if ($initials -eq "") {
	$pointemail= $lastname_without_ekezet10 + "." + $firstname_without_ekezet10 
}
Else{
	$pointemail= $initials_without_ekezet10 + "." + $lastname_without_ekezet10 + "." + $firstname_without_ekezet10
}

$organizationalunitudn = 	$organizationalunit
if ($organizationalunit -eq "PremierPartner") {
	$organizationalunitudn = "PremierPartner,ou=Alvallalkozok"
	$organizationalunit= "Alvallalkozok/PremierPartner"
	$title = "KFKI Premier Partner"
	$csoporttagsag5 = "X"
}
#---------------------------------------------------------------------------------------------------------------------


#logolás
echo "0-es pont sikeresen végrahajtva: OK" >> c:\Temp\kfki_newuser_v2.5.log
echo "¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤" >> c:\Temp\kfki_newuser_v2.5.log
echo "" >> c:\Temp\kfki_newuser_v2.5.log


#-------------------------------------Névellenőrző fgv meghívása------------------------------------------
#Ellenörizhetjük, az excelbe vajon jo nevet írtunk e be. Ha igen akkor tovább megy a program,
# ha nem, akkor a 'Kilép' gombra kattintva megszakithatjuk az egészet, majd javithatjuk az xls-t

$mehet = GenerateForm
If ($mehet -eq 0){
	break
}
#---------------------------------------------------------------------------------------------------------
#---------------------------------------------------------------------------------------------------------------------




#=====================================================================================================================
#        Újabb Törzs rész - GUI egybeintegrálva AD/Mailbox létrehozás, Csoport tagság, Folder jogok
#=====================================================================================================================
#Itt elindul egy GUI ami tájékoztat az egyes állapotokról, majd a végén lekéri az uj felhasználó adatait.
#Itt indul a Gui, majd a 3 förész külön tömbönként van jelülve a GUI kodon belül

#Generated Form Function
function GenerateForm2 {
#region Import the Assemblies
[reflection.assembly]::loadwithpartialname("System.Windows.Forms") | Out-Null
[reflection.assembly]::loadwithpartialname("System.Drawing") | Out-Null
#endregion

#region Generated Form Objects
$form1 = New-Object System.Windows.Forms.Form
$vScrollBar1 = New-Object System.Windows.Forms.VScrollBar
$label2 = New-Object System.Windows.Forms.Label
$label1 = New-Object System.Windows.Forms.Label
$InitialFormWindowState = New-Object System.Windows.Forms.FormWindowState
#endregion Generated Form Objects

#----------------------------------------------
#Generated Event Script Blocks
#----------------------------------------------
#Provide Custom Code for events specified in PrimalForms.
$szoveg= 
{

}

$OnLoadForm_StateCorrection=
{#Correct the initial state of the form to prevent the .Net maximized form issue
	$form1.WindowState = $InitialFormWindowState
}

#----------------------------------------------
#region Generated Form Code
$form1.Text = 'NewUser - beléptetés'
$form1.Name = 'form1'
$form1.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 450
$System_Drawing_Size.Height = 393#293
$form1.ClientSize = $System_Drawing_Size

$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 17
$System_Drawing_Size.Height = 313
$vScrollBar1.Size = $System_Drawing_Size
$vScrollBar1.TabIndex = 2
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 409
$System_Drawing_Point.Y = 58
$vScrollBar1.Location = $System_Drawing_Point
$vScrollBar1.Name = 'vScrollBar1'
$vScrollBar1.DataBindings.DefaultDataSourceUpdateMode = 0
$vScrollBar1.add_Scroll($szoveg)

$form1.Controls.Add($vScrollBar1)

$label2.TabIndex = 1
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 100
$System_Drawing_Size.Height = 20
$label2.Size = $System_Drawing_Size
$label2.Text = 'Futási lépések:'

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 23
$System_Drawing_Point.Y = 25
$label2.Location = $System_Drawing_Point
$label2.DataBindings.DefaultDataSourceUpdateMode = 0
$label2.Name = 'label2'
$label2.add_Click($handler_label2_Click)

$form1.Controls.Add($label2)

$label1.TabIndex = 0
$label1.BackColor = [System.Drawing.Color]::FromArgb(255,255,255,255)
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 404
$System_Drawing_Size.Height = 313
$label1.Size = $System_Drawing_Size
$label1.Text = ''

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 23
$System_Drawing_Point.Y = 58
$label1.Location = $System_Drawing_Point
$label1.DataBindings.DefaultDataSourceUpdateMode = 0
$label1.Name = 'label1'

$form1.Controls.Add($label1)

#¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤
#-------------------------------------------Innentől feldolgozás, és a lépések kiíratása--------------------------------------------
#¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤
$label1.text = "`n                                                 Uj_belepo_KFKI`n`n"
$label1.text = $label1.text + "Elkezdem a beléptetési folyamatot:`n - Felhasználó felvétele az AD-ba, Mailbox létrehozása"

#logolás
echo "¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤" >> c:\Temp\kfki_newuser_v2.5.log
echo "1. pont végrahajtása: AD/Mailbox létrehozás" >> c:\Temp\kfki_newuser_v2.5.log


#---------------------------------------------------------------------------------------------------------
#----------- 1. pont: Felhasznalo AD/Mailbox letrehozasa - uj belepo kfki exch 2007 v2.5 -----------------
#---------------------------------------------------------------------------------------------------------
#Ez volt régen külön PS Scriptben : uj belepo kfki exch 2007 v2.5
#

#Régen külön Ps file volt erre a részrel. Most ide be lett másolva. Ehhez  kell: Változó megfeleltetés
$displayname = $displayname
$alias = $samaccountname
$OU = "kfki.corp/KFKI/Felhasznalok/$organizationalunit"
$UPN = "$samaccountname@kfki.corp"
$SAN = $samaccountname
$FN = $firstname
$initials = $initials
$LN = $lastname
$Database_type = $mailboxdatabase
$managerdp = $manager
$mobil = $mobile
$phone = $telephonenumber
$title = $title
$department = $department
$company = $company
$office = $physicaldeliveryofficename
$pointemail = $pointemail
$koltseg_hely = $koltseg_hely
$belepes_datuma = $belepes_datuma


$email = $pointemail + "@kfkizrt.hu"

######################################################################################################

# logolas
$datum = get-date
echo $datum >> c:\Temp\kfki_newuser_v2.5.log

echo "Paraméterek: " $displayname $alias $OU $UPN $SAN $FN $initials $LN $Database_type $managerdp $mobil $phone $title $department $company $office $pointemail $koltseg_hely $belepes_datuma >> c:\Temp\kfki_newuser_v2.5.log
echo "Paraméterek vége!" >> c:\Temp\kfki_newuser_v2.5.log

# database_type alapjan megkeresni az adatbazist ahova lehet usert rakni

$vanhely = $false

# VIP adatbazisok lekerdezes, melyikbe rakjuk a usert?
if ($Database_type -ieq "vip")
{

	if ($vanhely -eq $false)
	{
	$database = "k-mail3\vip1\vip1"
	if (50 -gt (Get-Mailbox -Database $database).count)
		{
		$vanhely = $true
		}
	}

	if ($vanhely -eq $false)
	{
	$database = "k-mail3\vip3\vip3"
	if (50 -gt (Get-Mailbox -Database $database).count)
		{
		$vanhely = $true
		}
	}
}



# Normal adatbazisok lekerdezes, melyikbe rakjuk a usert?

if ($Database_type -ieq "normal")
{

	if ($vanhely -eq $false)
	{
	$database = "k-mail3\nrml2\nrml2"
	if (100 -gt (Get-Mailbox -Database $database).count)
		{
		$vanhely = $true
		}
	}

	if ($vanhely -eq $false)
	{
	$database = "k-mail3\nrml3\nrml3"
	if (100 -gt (Get-Mailbox -Database $database).count)
		{
		$vanhely = $true
		}
	}

	if ($vanhely -eq $false)
	{
	$database = "k-mail3\nrml4\nrml4"
	if (100 -gt (Get-Mailbox -Database $database).count)
		{
		$vanhely = $true
		}
	}

	if ($vanhely -eq $false)
	{
	$database = "k-mail3\nrml5\nrml5"
	if (100 -gt (Get-Mailbox -Database $database).count)
		{
		$vanhely = $true
		}
	}

	if ($vanhely -eq $false)
	{
	$database = "k-mail3\nrml6\nrml6"
	if (100 -gt (Get-Mailbox -Database $database).count)
		{
		$vanhely = $true
		}
	}

	if ($vanhely -eq $false)
	{
	$database = "k-mail3\nrml7\nrml7"
	if (100 -gt (Get-Mailbox -Database $database).count)
		{
		$vanhely = $true
		}
	}

	if ($vanhely -eq $false)
	{
	$database = "k-mail3\nrml8\nrml8"
	if (100 -gt (Get-Mailbox -Database $database).count)
		{
		$vanhely = $true
		}
	}

	if ($vanhely -eq $false)
	{
	$database = "k-mail3\nrml9\nrml9"
	if (100 -gt (Get-Mailbox -Database $database).count)
		{
		$vanhely = $true
		}
	}
	
	if ($vanhely -eq $false)
	{
	$database = "k-mail3\nrml11\nrml11"
	if (100 -gt (Get-Mailbox -Database $database).count)
		{
		$vanhely = $true
		}
	}
	
	if ($vanhely -eq $false)
	{
	$database = "k-mail3\nrml12\nrml12"
	if (100 -gt (Get-Mailbox -Database $database).count)
		{
		$vanhely = $true
		}
	}
	
	if ($vanhely -eq $false)
	{
	$database = "k-mail3\nrml13\nrml13"
	if (100 -gt (Get-Mailbox -Database $database).count)
		{
		$vanhely = $true
		}
	}
}
 # vege: database_type alapjan megkeresni az adatbazist ahova lehet usert rakni

echo "Szabad adatbázis:" $database >> c:\Temp\kfki_newuser_v2.5.log



######################################################################################################


#   New user create


$password = ConvertTo-SecureString 'Pa$$w0rd' -AsPlainText -Force

New-Mailbox -Name $displayname -Alias $alias -OrganizationalUnit $OU -UserPrincipalName $UPN -SamAccountName $SAN -FirstName $FN -Initials $initials -LastName $LN -Password $password -ResetPasswordOnNextLogon $false -Database $Database  >> c:\Temp\kfki_newuser_v2.5.log

# vege: New user create

$user = get-mailbox -Identity $SAN

echo "Létrehozott user:" $user >> c:\Temp\kfki_newuser_v2.5.log

######################################################################################################



# KFKI Premier Partner adatmodositasa #####################
IF ($title -eq "KFKI Premier Partner")
{
$email = $SAN + "@pp.kfkizrt.hu"
$pointemail = $pointemail + "@pp.kfkizrt.hu"


$user.EmailAddresses += $pointemail
Set-Mailbox $user -EmailAddressPolicyEnabled $false -EmailAddresses $user.EmailAddresses

$user.PrimarysmtpAddress = $pointemail
Set-Mailbox $user -PrimarySmtpAddress $user.PrimarysmtpAddress

}


# KFKI Premier Partner adatmodositasa #####################



# loginname email cim hozzaadasa


$user.EmailAddresses += $email
Set-Mailbox $user -EmailAddresses $user.EmailAddresses

# loginname email cim hozzaadasa


#   AD parameterek beallitasa

$manager = Get-Mailbox -Filter "Displayname -eq '$managerdp'"

$setuser = Get-User -Identity $SAN

$setuser.CountryOrRegion = 'Hungary'
$setuser.MobilePhone = $mobil
$setuser.Phone = $phone
$setuser.title = $title
$setuser.Department = $department
$setuser.Manager = $manager.Identity
$setuser.Company = $company
$setuser.Office = $office


Set-User $setuser -MobilePhone $setuser.MobilePhone -Title $setuser.title -Department $setuser.Department -Manager $setuser.Manager -Company $setuser.Company -Office $setuser.Office -CountryOrRegion $setuser.CountryOrRegion -Phone $setuser.Phone >> c:\Temp\kfki_newuser_v2.5.log

# vege:  AD parameterek beallitasa


#------------------------------------------------------------------------------------------------------------------------------------------




#------------------------------------------------------------------------------------------------------------------------------------------


# uj beleponek email kuldes

$From    = New-Object system.net.Mail.MailAddress "helpdesk@kfkizrt.hu", "KFKI Helpdesk"
$To      = new-object system.net.mail.MailAddress $email
$Message = New-Object system.Net.Mail.MailMessage $From, $To


echo $email >> c:\Temp\kfki_newuser_v2.5.log
 
 
# Populate message
$Message.Subject = "Tudnivalók"

$Message.IsBodyHTML = $true

$html = "
Kedves Kollega!<br>
<br>
Mielőtt megkezdenéd a cégen belüli tevékenységedet, az alábbi linken található IT működéssel kapcsolatos dokumentumokat kérlek olvasd el:<br>
<br>
http://wwwin/txtlstvw.aspx?LstID=1d4444eb-b0f9-4b5e-b43c-fdd44554d3d6<br>
<br>
http://intraapps.wwwin/ITHELP/eszkozok/KFKI_file_tarolas.ppt<br>
<br>
http://wwwin/WebHelp/Document%20Library/KFKI%20Intranet%20ismertető.ppt<br>
<br>
http://intraapps.wwwin/ITHELP/Biztonsag_security/Adatok_titkositasa.ppt<br
<br>
http://l-moodle1/moodle/file.php/1/manual/moodle_manual.pdf<br> 
<br>
http://wwwin/KFKI/IT/it.aspx<br> 
<br>
Ha bármilyen kérdésed/észrevételed van, fordulj bizalommal a HelpDesk -hez mely a hét minden napján nonstop rendelkezésedre áll:<br>
<br>
Belső telefonszám: 5555<br>
<br>
Külső telefonszám: 06 / 1 / 236-6702 | 06 / 80 / 408-080<br>
<br>
E-mail cím: helpdesk@kfkizrt.hu<br>
<br>
Üdvözlettel:<br>
KFKI Data Center csoport
"


$Message.Body = $html
 
 
# Create SMTP Client
$Server = "k-mail3hc"
$Client = New-Object System.Net.Mail.SmtpClient $server
$Client.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials
 
# send the message

$client.Send($Message);
# uj beleponek email kuldes vege


#------------------------------------------------------------------------------------------------------------------------------------------









#---------------------------------------------------------------------------------------------------------
#------------------------------uj belepo kfki exch 2007 v2.5 - VÉGE---------------------------------------
#---------------------------------------------------------------------------------------------------------


Set-QADUser -Identity $samaccountname -LogonScript $logonscript
Set-QADUser -Identity $samaccountname -HomeDrive $homedrive
#Set-QADUser -Identity $samaccountname -HomeDirectory $homedirectory
Set-Mailbox -Identity $samaccountname -CustomAttribute1 $flcsoport

#----------------------------------Felhasznalo AD/Mailbox letrehozasa VÉGE---------------------------------------------
#----------------------------------------------------------------------------------------------------------------------



#Véget ért az 1. pont, és ezt jelezzük a GUI-n:
$label1.text = $label1.text + "                  -                OK`n"
#Majd jön a következö, 2.pont
$label1.text = $label1.text + " - Csoporttagságok beállitása                                       "  

#logolás
echo "1-es pont sikeresen végrahajtva: OK" >> c:\Temp\kfki_newuser_v2.5.log
echo "¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤" >> c:\Temp\kfki_newuser_v2.5.log
echo "" >> c:\Temp\kfki_newuser_v2.5.log
echo "¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤" >> c:\Temp\kfki_newuser_v2.5.log
echo "2. pont égrahajtása: Csoporttagságok beállítása" >> c:\Temp\kfki_newuser_v2.5.log

#--------------------------- 2. pont: Csoporttagsag beallitas ---------------------------------------------
#----------------------------------------------------------------------------------------------------------
$userdn = Get-QADUser -Identity $samaccountname
if ($csoporttagsag1 -eq "X"){Add-QADGroupMember -Identity _KFKI -Member $userdn}
if ($csoporttagsag2 -eq "X"){Add-QADGroupMember -Identity KFKI_gyakornokok -Member $userdn}
if ($csoporttagsag3 -eq "X"){Add-QADGroupMember -Identity KFKI_Budapest_munkavallalok -Member $userdn}
if ($csoporttagsag4 -eq "X"){Add-QADGroupMember -Identity KFKI_Videk_munkavallalok -Member $userdn}
if ($csoporttagsag5 -eq "X"){Add-QADGroupMember -Identity KFKI_PremierPartners -Member $userdn}
#---------------------------------------------------------------------------------------------------------

#loggolás

echo "############################################################################" >> c:\Temp\kfki_newuser_v2.5.log
echo ("Mailbox létrehozása sikeres?:") >> c:\Temp\kfki_newuser_v2.5.log
Get-Mailbox -Identity $san | Select-Object Samaccountname,Displayname, emailaddresses, PrimarySmtpAddress, windowsemailaddress | fl >> c:\Temp\kfki_newuser_v2.5.log

echo "############################################################################" >> c:\Temp\kfki_newuser_v2.5.log
echo ("Ad adatok renben vannak?:") >> c:\Temp\kfki_newuser_v2.5.log
Get-Mailbox -Identity $san | Get-User | Select-Object Company, Department, Title, Office, Firstname, Initials, Lastname, Manager >> c:\Temp\kfki_newuser_v2.5.log
echo "############################################################################" >> c:\Temp\kfki_newuser_v2.5.log
echo ("Ad csoporttagságok renben vannak?:") >> c:\Temp\kfki_newuser_v2.5.log
Get-QADMemberOf $san >> c:\Temp\kfki_newuser_v2.5.log
echo "############################################################################" >> c:\Temp\kfki_newuser_v2.5.log
echo ("Tudnivalók email fogadásának a logja:") >> c:\Temp\kfki_newuser_v2.5.log
echo (get-messagetrackinglog -Recipients: $email -Sender "helpdesk@kfkizrt.hu" -Server "k-mail3hc1" -EventID "RECEIVE" -MessageSubject "Tudnivalók") >> c:\Temp\kfki_newuser_v2.5.log
echo (get-messagetrackinglog -Recipients: $email -Sender "helpdesk@kfkizrt.hu" -Server "k-mail3hc2" -EventID "RECEIVE" -MessageSubject "Tudnivalók") >> c:\Temp\kfki_newuser_v2.5.log
echo "############################################################################" >> c:\Temp\kfki_newuser_v2.5.log
$subject = "KFKI új belépő: " + $san
echo ("Bozsik Laciéknak küldött email fogadásának a logja: ") >> c:\Temp\kfki_newuser_v2.5.log
echo (get-messagetrackinglog  -Sender "KFKI_DATA_Center_uzemeltetes@kfkizrt.hu" -Server "k-mail3hc1" -EventID "RECEIVE" -MessageSubject $subject) >> c:\Temp\kfki_newuser_v2.5.log
echo (get-messagetrackinglog  -Sender "KFKI_DATA_Center_uzemeltetes@kfkizrt.hu" -Server "k-mail3hc2" -EventID "RECEIVE" -MessageSubject $subject) >> c:\Temp\kfki_newuser_v2.5.log

#Véget ért az 2. pont, és ezt jelezzük a GUI-n:
$label1.text = $label1.text + "                  -                OK`n"
#Majd jön a következö, 3.pont
$label1.text = $label1.text + " - Home folder jogok állítása                                         "

#logolás
echo "2-es pont sikeresen végrahajtva: OK" >> c:\Temp\kfki_newuser_v2.5.log
echo "¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤" >> c:\Temp\kfki_newuser_v2.5.log
echo "" >> c:\Temp\kfki_newuser_v2.5.log
echo "¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤" >> c:\Temp\kfki_newuser_v2.5.log
echo "3. pont végrahatása: User könyvtárainak jogosultásg beállítása" >> c:\Temp\kfki_newuser_v2.5.log


#-------------------------- 3. pont: Home folder jogosultsag allitas --------------------------------------
#----------------------------------------------------------------------------------------------------------
#                    samaccountname könyvtár létrehozása, és jogok beállítása
$directory1 = "\\K-file1\k-users"
New-Item -Path $directory1 -Name "$samaccountname" -type Directory
$directory = "\\K-file1\k-users\$samaccountname"
# alkönyvtárra öröklődés beállítása: szükséges flagek
$inherit = [system.security.accesscontrol.InheritanceFlags]"ContainerInherit, ObjectInherit"
$propagation = [system.security.accesscontrol.PropagationFlags]"None"
$acl = Get-Acl $directory
#jogok megadása
$accessrule = New-Object system.security.AccessControl.FileSystemAccessRule("KFKI\$samaccountname", "Modify", $inherit, $propagation, "Allow")
$acl.AddAccessRule($accessrule)
$accessrule = New-Object system.security.AccessControl.FileSystemAccessRule("_KFKI", "Read", $inherit, $propagation, "Allow")
$acl.AddAccessRule($accessrule)
#jogok könyvtárra alkalmazása
set-acl -aclobject $acl $directory

#                samaccountname\private könyvtár létrehozása, és jogok beállítása
$directory1 = "\\K-file1\k-users\$samaccountname"
New-Item -Path $directory1 -Name "private" -type Directory
$directory = "\\K-file1\k-users\$samaccountname\private"
$inherit = [system.security.accesscontrol.InheritanceFlags]"ContainerInherit, ObjectInherit"
$propagation = [system.security.accesscontrol.PropagationFlags]"None"
$acl = Get-Acl $directory
#jogok megadása
$acl.SetAccessRuleProtection($TRUE,$FALSE)
$acl.RemoveAccessRuleAll
$accessrule = New-Object system.security.AccessControl.FileSystemAccessRule("KFKI\lnx_account_useradmin", "FullControl", $inherit, $propagation, "Allow")
$acl.AddAccessRule($accessrule)
$accessrule = New-Object system.security.AccessControl.FileSystemAccessRule("KFKI\$samaccountname", "Modify", $inherit, $propagation, "Allow")
$acl.AddAccessRule($accessrule)
$accessrule = New-Object system.security.AccessControl.FileSystemAccessRule("BUILTIN\Administrators", "FullControl", $inherit, $propagation, "Allow")
$acl.AddAccessRule($accessrule)
$accessrule = New-Object system.security.AccessControl.FileSystemAccessRule("SYSTEM", "FullControl", $inherit, $propagation, "Allow")
$acl.AddAccessRule($accessrule)
#jogok könyvtárra alkalmazása
set-acl -aclobject $acl $directory

#----------------------------------------------------------------------------------------------------------


#Véget ért az 3. pont, és ezt jelezzük a GUI-n:
$label1.text = $label1.text + "                  -                OK`n"
$label1.text = $label1.text + "`nÚj Felhasználó létrehozása, és a szükséges dolgok beállítása véget ért."

#Adatok lekérdezése
$a = Get-QADUser -Identity $samaccountname
$c = get-qaduser $a.manager
$felettes = $c.DisplayName

$label1.text = $label1.text + "`nAdatai Lekérdezve:"
$label1.text = $label1.text + "`n   ¤ Név:           $($a.DisplayName)"
$label1.text = $label1.text + "`n   ¤ Osztály:      $($a.Department)"
$label1.text = $label1.text + "`n   ¤ Vezető:      $felettes"
$label1.text = $label1.text + "`n   ¤ e-mail:        $($a.Email)"
$label1.text = $label1.text + "`n`n   `# Csoport tagságok: "
$a.MemberOf | foreach {
$label1.text = $label1.text + "`n            $_"
}
#-------------------------------------------------------------------------------------------------------------------------------------
#-------------------------------------------------------------------------------------------------------------------------------------

# Bozsik Lacieknak email kuldes
$From    = New-Object system.net.Mail.MailAddress "KFKI_DATA_Center_uzemeltetes@kfkizrt.hu", "KFKI Data Center Üzemeltetés - Új belépő felvétele"
$To      = new-object system.net.mail.mailaddress "KFKI_Uj_Belepo_Script@kfki.corp"
$Message = New-Object system.Net.Mail.MailMessage $From, $To

 
# Populate message
$Message.Subject = "KFKI új belépő: " + $san;
$Message.IsBodyHTML = $true

$html1 = "
Login név		: <b>" + $SAN + "</b><br>
E-mail cím		: <b>" + $a.Email + "</b><br>
Belépés dátuma	: <b>" + $belepes_datuma + "</b><br>
Költség hely	: <b>" + $koltseg_hely + "</b><br>
Beosztás		: <b>" + $title + "</b><br>
Teljes név		: <b>" + $displayname + "</b><br>

<br>
Üdvözlettel:<br>
KFKI Data Center csoport
"


$Message.Body = $html1

# Create SMTP Client
$Server = "k-mail3hc"
$Client = New-Object System.Net.Mail.SmtpClient $server
$Client.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials
 

# send the message
$Client.Send($Message);
#------------------------------------------------
#------------------------------------------------


#endregion Generated Form Code

#Save the initial state of the form
$InitialFormWindowState = $form1.WindowState
#Init the OnLoad event to correct the initial state of the form
$form1.add_Load($OnLoadForm_StateCorrection)
#Show the Form
$form1.ShowDialog()| Out-Null
} 
#Vége a GUI függvénynek.. Igy már csak meg kell hívnunk hogy lefusson!
GenerateForm2

#logolás
echo "3-es pont sikeresen végrahajtva: OK" >> c:\Temp\kfki_newuser_v2.5.log
echo "¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤" >> c:\Temp\kfki_newuser_v2.5.log
echo "" >> c:\Temp\kfki_newuser_v2.5.log