#-------------------------------------------------------------------------------------------------------
#   a vonalak k�z�tti parancsok egyszerre futtatand�ak
#-----------------------------------------------------------------------------------------------------
#$HOME                           # user home k�nyvt�ra c:\document and setting\ user, 
#$PSHOME                         #Powershell home k�nyvt�ra c:\Windows\system32\WindowsPowerShell\v1.0
#Set-Location $HOME              #
				#-----------------------Saj�t v�ltoz�k:
#Get-Variable                    #v�ltoz�k list�ja
#----------------------------------------------------------------------------------------------------
#Set-Variable -Name x -Value 2  #x v�ltoz�, 2-s �rt�kkel
#Get-Variable
#-------------------------------------------------------------------------------------------------------
#$x=3                           #x v�ltoz� 3-s �rt�kkel
#$x                             #ki�rja az �rt�ket
#----------------------------------------------------------------------------------------------
#$sz="hello"+" vilag"                    # a string, + konkaten�ci�
#$sz                     
#-------------------------------------------------------------------------------------------
# "hello$x"                      #�k�z�tt a v�ltoz� tartalma ker�l be a sz�vegbe � hello3
# 'hello$x'                              # ��-k�z�tt a sz�veg, a v�ltoz� azonos�t�j�t t�rolja --- hello$x
#---------------------------------------------------------------------------------------------
#$ma=Get-Date                   #objektumokat is t�rolhat!
#$ma.Year                       #az objektum Year adattagja
#----------------------------------------------------------------------
#$t=1,2,3,4,5                   #tomb hasznalata
#write $t
#-------------------------------------------------------------------------------------
#Remove-Variable �Name x                # azonos�t� elt�vol�t�sa
#Get-Variable
#----------------------------------------------------------------------------------------
#$tombbe=Get-Content tartalom.txt
#$tombbe                                 # a teljes f�jl tartalm�t beolvassa
#$tombbe |get-member             # a t�mbbel kapcsolatos el�rhet� adatok, met�dusok
#----------------------------------------------------------------------------------------------------
