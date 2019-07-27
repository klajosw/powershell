#----------------------------------------------------------------------------------------------------------------------------------------
#------------------------------------------------------Script leírás---------------------------------------------------------------------
#----------------------------------------------------------------------------------------------------------------------------------------
#Egy gép újratelepítése esetén a felhasználó adatait ideiglenesen egy file-szerverre mentsük. A felhasznáók állományai egy magadott 
#mappában a hozzájuk tartozó incidens számukkal elnevezett mappákban helyezkednek el. Miután az újrateplepítés befejeződött az FL 
#szintén egy script futtatása alapján egy  .kfkirestore  kitetrjesztésű file-t helyez el a megfelelő könyvtárba (file: $SamacountName.kfkirestore).
#----------------------------------------------------------------------------------------------------------------------------------------
#A script feladata: 
# - Ellenőrzi ha a .kfkirestore file nem került 14napon belül a könyvtárba és egy emailt küld a DM-nek erről, aki utánnajár.
# - Ha a file ott van, ellenőrzi mióta van ott
#		- Minden nap értesítést küld a felhasználónak hogy 9 nap mulva törlődnek az állmányok
#		- 7.napon küld az FL-nek egy listát a könyvtárakról amik 2nap mulva törlődnek. A felhasználót is értesíti
#		- 9. napon felhasználómnak értesítés hogy törlődtek az állományai, majd törli a megelelő könyvtárat.
#----------------------------------------------------------------------------------------------------------------------------------------
#Script futás megszakadás lehetséges okai:
#	- Megváltozott a könyvtárak elérési útvonala
#	- A scriptet futtató személynek nincs jogosultsága a file-szerverhez
#	- Az email küldés elakad
#	- A script futtatójának nincs joga 'helpdesk' nevében emailt küldeni e
#----------------------------------------------------------------------------------------------------------------------------------------
#----------------------------------------------------------------------------------------------------------------------------------------

#email küldö fügvény. Paraméterként megkapja hog ykitöl, kinek és mit küldjön
function email($from, $to, [string] $html) {
	$Message = New-Object system.Net.Mail.MailMessage $From, $To
	$Message.Subject = "Értesítés"
	$Message.IsBodyHTML = $true
	$Message.Body = $html
 	$Server = "k-mail3hc"
	$Client = New-Object System.Net.Mail.SmtpClient $server
	$Client.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials
	$client.Send($Message);
}
#a központi mappa amiben az incidens szám névvel rendelkezö mappák vannak
$path = "\\k-ub1\User_backups"
$dirs = New-Object system.collections.arraylist				#közponi mappán egy mélységen belüli összes mappa(incidens szám nevűek)- tömb
#aktuális dátum eltárolása, és az óra, perc nullázása
$ma = Get-Date			
$ma = $ma.Date
$14nap = New-Object system.collections.arraylist			#14 napja nincs file a könyvtárba: ezek nevei ebbe vannak összegyüjtve
$7nap = New-Object system.collections.arraylist
$9nap = New-Object system.collections.arraylist

#központi könyvtárbol egy mélységben az incidens szám könyvtárak összegyüjtése tömbbe
dir $path  | ForEach-Object `
{
	if ($_.psIsContainer -eq $true) {
		[void] $dirs.Add( $_.name)
	}
}
#Törzs: file létezés ellenörzs, és értesítő emailek a felhasználóknak. Könyvtáranként futnak a vizsgálatok
$dirs | ForEach-Object `
{
	#Ha nincs .kfkirestore file akkor a mappa nevét eltárolja, feltéve ha több mint 14 napja nincs, különben tovább megy
	if((Test-Path ($path + "\" + $_ + "\*.kfkirestore")) -ceq $false){
		$folder = [System.IO.DirectoryInfo] ($path + "\" + $_)
		$f_create =[datetime] $folder.CreationTime
		$f_create = $f_create.Date
		if (($ma - $f_create)-cge 14){
			[void] $14nap.add($_)
		}
	}
	else{
		#file létrehozás dátumának meghatározása
		$file = dir ($path + "\" + $_ + "\*.kfkirestore")
		$create =[datetime] $file.CreationTime
		$create = $create.Date
		$name = $file.name.Replace(".kfkirestore","")
		$hatra = ($ma - $create).days
		#email a felhasználonak naponta a 7.napig
		if ($hatra -ne 7){
			#kitől, kinek, mit, majd az email küldö függvény meghivása ezekkel a paraméterekkel
			$From    = New-Object system.net.Mail.MailAddress "helpdesk@kfkizrt.hu", "KFKI Helpdesk"
			#$To      = new-object system.net.mail.MailAddress "$name@kfki.corp"
			$To      = new-object system.net.mail.MailAddress "molnarakos@kfki.corp"
			$html = "
			Kedves Kollega!<br>
			<br>
			TESZT MAIL<>
			Még " + (9 - $hatra) + " napod van arra, hogy átmásold és letöröld a könyvtárad. Az idö leteltével automatikusan törlődik!<br>
			<br>
			"	
			email $from $to $html
			
			#kitől, kinek, mit, majd az email küldö függvény meghivása ezekkel a paraméterekkel
			$From    = New-Object system.net.Mail.MailAddress "helpdesk@kfkizrt.hu", "KFKI Helpdesk"
			#$To      = new-object system.net.mail.MailAddress "$name@kfki.corp"
			$To      = new-object system.net.mail.MailAddress "kokaipeter@kfki.corp"
			$html = "
			Kedves Kollega!<br>
			<br>
			TESZT MAIL<br>
			Még " + (9 - $hatra) + " napod van arra, hogy átmásold és letöröld a könyvtárad. Az idö leteltével automatikusan törlődik!<br>
			<br>
			"	
			email $from $to $html
		}
		#email a felhasználonak a 7.napon hogy 2 nap mulva törlés
		if ($hatra -eq 7){
			[void] $7nap.add($_)
		}
		if(($hatra -eq 7) -or ($hatra -eq 8)){
			#értesítés hogy 2 napon belül törölve lesznek
			$From    = New-Object system.net.Mail.MailAddress "helpdesk@kfkizrt.hu", "KFKI Helpdesk"
			$To      = new-object system.net.mail.MailAddress "$name@kfki.corp"
			$html = "
			Kedves Kollega!<br>
			<br>
			A könytárak " + (9 - $hatra) + " nap mulva törölve lesznek.<br>
			<br>
			"
			email $from $to $html
		}
		#email a felhasználonak hogy törölve vannak az állományai, majd törlés
		if($hatra -eq 9){
			#értesítés a törlésről
			$From    = New-Object system.net.Mail.MailAddress "helpdesk@kfkizrt.hu", "KFKI Helpdesk"
			$To      = new-object system.net.mail.MailAddress "$name@kfki.corp"
			$html = "
			Kedves Kollega!<br>
			<br>
			A könytárak töröltük!<br>
			Még lehetősége van az adatainak visszanyerésére a mentésből ha ír a következő cimre: <br>
			További szép napot!<br>
			<br>
			"
			email $from $to $html
			
			#a megfelelő mappa törlése almappákkal együtt
			$path2 = $path + "\" + $_
			Remove-Item $path2 -Recurse
		}
	}
}
#email a FL nek  a mappákrol amelyek két nap mulva törlésre kerülnek
if ($7nap -cne $null){
	#email a fl-nek
	$From    = New-Object system.net.Mail.MailAddress "helpdesk@kfkizrt.hu", "KFKI Helpdesk"
	$To      = new-object system.net.mail.MailAddress "KFKI_IT_uzemeltetes_frontline_BO_kozpont_ITFL@kfki.hu"
	$html = "
	Kedves Kollega!<br>
	<br>
	A következő könyvtárak lesznek  törölve 2 nap múlva:`n $14nap<br>
	<br>
	"	
	email $from $to $html
}
else{
	"Háá de jo...üres  a 7 napos cucc"
}
#email a DM nek a mappákról amikben több mint 14 napja nincsen .kfkirestore file
if ($14nap -cne $null){
	#email a dm-nek
	$From    = New-Object system.net.Mail.MailAddress "helpdesk@kfkizrt.hu", "KFKI Helpdesk"
	$To      = new-object system.net.mail.MailAddress "LNX_Desktop_Management_Team@kfki.hu"
	$html = "
	Kedves Kollega!<br>
	<br>
	A következő könyvtárakban még nincs `".kfkirestore`" kiterjesztésű file már 20napja:`n $14nap `n<br>
	<br>
	"	
	email $from $to $html
}
else{
	"Háá de jo...üres  a 14 napos cucc"
}

