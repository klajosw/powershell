#Get-ChildItem | Out-File konyvtar.txt                                #cs�vezet�k 
#Get-Date | Get-Member                                  #objektumok felt�rk�pez�se
#(Get-Date).year                                                # objektum tagadat�nak el�r�se
#Get-ChildItem konyvtar.txt | Get-Member                        # egy f�jl adatait tartalmaz� objektum tagjai
#(Get-ChildItem konyvtar.txt).FullName                          #f�jl neve �tvonallal
#Get-ChildItem | Where {$_.Length -gt 1250}                      #100-n�l hosszabb f�jlokat list�zza ki #Get-ChildItem | Where {$_.Extension �eq �.txt�}                         #txt f�jlokat list�zza
#Get-Childitem| Sort-Object -Property LastTimeWrite -Descending         #a k�nyvt�rat cs�kken� sorrendben utols� �r�s 
#Get-Process                                            # fut� processzeket list�zza
#Get-Process | Format-List -Property Name, Id                   # a processzeket list�zza, csak az ID-t �s nevet 
#Get-Process |Format-table -AutoSize                            #t�bl�zatos form�ban �rja ki, 
#Get-Process| Where-Object {$_.ProcessName -lt "kozepe"}        #sz�r�s
#Get-Process| Sort-Object -Property CPU| select-Object -Property Name, Cpu -Last 5 | Format-Table -AutoSize 
	#a processzeket �ll�tsuk sorba a CPU szerint, sz�rj�k le a name �s cpu tulajdons�gra �s t�bla form�ba �rjuk ki. 
