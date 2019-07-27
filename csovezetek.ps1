#Get-ChildItem | Out-File konyvtar.txt                                #csõvezeték 
#Get-Date | Get-Member                                  #objektumok feltérképezése
#(Get-Date).year                                                # objektum tagadatának elérése
#Get-ChildItem konyvtar.txt | Get-Member                        # egy fájl adatait tartalmazó objektum tagjai
#(Get-ChildItem konyvtar.txt).FullName                          #fájl neve útvonallal
#Get-ChildItem | Where {$_.Length -gt 1250}                      #100-nál hosszabb fájlokat listázza ki #Get-ChildItem | Where {$_.Extension –eq „.txt”}                         #txt fájlokat listázza
#Get-Childitem| Sort-Object -Property LastTimeWrite -Descending         #a könyvtárat csökkenõ sorrendben utolsó írás 
#Get-Process                                            # futó processzeket listázza
#Get-Process | Format-List -Property Name, Id                   # a processzeket listázza, csak az ID-t és nevet 
#Get-Process |Format-table -AutoSize                            #táblázatos formában írja ki, 
#Get-Process| Where-Object {$_.ProcessName -lt "kozepe"}        #szûrés
#Get-Process| Sort-Object -Property CPU| select-Object -Property Name, Cpu -Last 5 | Format-Table -AutoSize 
	#a processzeket állítsuk sorba a CPU szerint, szûrjük le a name és cpu tulajdonságra és tábla formába írjuk ki. 
