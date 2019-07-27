#-------------------------------------------------------------------------------------------------------
#   a vonalak közötti parancsok egyszerre futtatandóak
#-----------------------------------------------------------------------------------------------------
#$HOME                           # user home könyvtára c:\document and setting\ user, 
#$PSHOME                         #Powershell home könyvtára c:\Windows\system32\WindowsPowerShell\v1.0
#Set-Location $HOME              #
				#-----------------------Saját változók:
#Get-Variable                    #változók listája
#----------------------------------------------------------------------------------------------------
#Set-Variable -Name x -Value 2  #x változó, 2-s értékkel
#Get-Variable
#-------------------------------------------------------------------------------------------------------
#$x=3                           #x változó 3-s értékkel
#$x                             #kiírja az értéket
#----------------------------------------------------------------------------------------------
#$sz="hello"+" vilag"                    # a string, + konkatenáció
#$sz                     
#-------------------------------------------------------------------------------------------
# "hello$x"                      #„között a változó tartalma kerül be a szövegbe – hello3
# 'hello$x'                              # ’’-között a szöveg, a változó azonosítóját tárolja --- hello$x
#---------------------------------------------------------------------------------------------
#$ma=Get-Date                   #objektumokat is tárolhat!
#$ma.Year                       #az objektum Year adattagja
#----------------------------------------------------------------------
#$t=1,2,3,4,5                   #tomb hasznalata
#write $t
#-------------------------------------------------------------------------------------
#Remove-Variable –Name x                # azonosító eltávolítása
#Get-Variable
#----------------------------------------------------------------------------------------
#$tombbe=Get-Content tartalom.txt
#$tombbe                                 # a teljes fájl tartalmát beolvassa
#$tombbe |get-member             # a tömbbel kapcsolatos elérhetõ adatok, metódusok
#----------------------------------------------------------------------------------------------------
