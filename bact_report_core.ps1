<#
.SYNOPSIS
    BachReport cmdlet-ek
.DESCRIPTION
    További segítség: Get-Help <cmdlet-név>
    A cmdletek listáját lásd a Related Links alatt.
.LINK
    Db-Connect
    Db-Connect2
    Execute-Query
    Execute-Query-To-Csv
    Execute-Scalar
    Execute-Non-Query
    Dispose-Db
    Dispose-Result
    Create-Workbook
    Add-Worksheet
    Add-Comment
    Save-Workbook
    Set-Worksheet-Styles
    Set-Workbook-Properties
    Send-Mail
    Start-Log
    Write-Log
    End-Log
    Clear-Log
    Set-LogLevel
    Get-LogLevel
    Get-Log-Encoding
    Set-Log-Encoding
    Load-File
    Make-Pivot
    Xlsx-To-Csv
    Save-To-Csv
    Export-To-Html
.EXAMPLE
    . .\BachReport.ps1
.NOTES
    Fájlnév:       BachReport.ps1
    Verzió:        3.0.0
    Dátum:         2013-02-20
    Készítette:    Gárdonyi László Andráds
    Email:         gardonyi.laszlo@t-systems.hu
    Telefon:       +36304443326
    Követelmények: PowerShell Version 2.0, .NET 4.0
    Függőségek:    BachReport.dll (GAC)
                   EPPlus.dll (GAC)
                   CvsHelper.dll (GAC)
                   MySql.Data.dll (GAC)
#>
function BachReport { Get-Help BachReport }

Set-Alias -Name Bach-Report -Value BachReport

<#
.SYNOPSIS
    Betölti a BachReport.dll-t a PowerShell környezetbe.
.DESCRIPTION
    Betölti a BachReport.dll-t a PowerShell környezetbe.
    Meghívása a szkript futtatásakor automatikusan megtörténik, nem szükséges újból futtatni.
.PARAMETER BaseDir
    A BachReport.dll elérési útvonala
    Alapértelmezett érték: aktuális könyvtár
.EXAMPLE
    . .\BachReport.ps1
##
function Import-BachReport
{
  param ([string] $BaseDir = (Get-Location).Path)
  $path = [System.IO.Path]::GetFullPath([System.IO.Path]::Combine($baseDir, "BachReport.dll"))
  if (!(Test-Path $path -PathType Leaf))
  {
    throw "Nem találom a BachReport.dll fájl!"
  }
  [Reflection.Assembly]::LoadFile($path) | out-null
}

# BachReport.dll betöltése
if (![string]::IsNullOrWhiteSpace($args[0]) -and (Test-Path $args[0] -PathType Container))
{
  Import-BachReport $args[0]
} else {
  Import-BachReport
}
#>

try
{
    # BachReport betöltése
    [Reflection.Assembly]::Load("BachReport, Version=3.0.0.0, Culture=neutral, PublicKeyToken=abd5f8293fa96958, processorArchitecture=MSIL") | out-null
}
catch
{
    Write-Host "A 3.0-ás verziójú BachReport.dll nincs telepítve a rendszeren vagy nincs engedélyezve a PowerShell számára a .NET 4.0-ás szerelvények használata. A telepítéshez indítsd el az install_epplus.cmd és az install_BachReport.cmd fájlokat adminisztrátori parancssorból. A .NET 4.0 engedélyezéséhez indítsd el az Enable .NET 4.0 for PowerShell.cmd nevű fájlt szintén adminisztrátori parancssorból."
}

<#
.SYNOPSIS
    Elindítja a naplózást.
.DESCRIPTION
    Elindítja a naplózást a megadott fájlba.
    A naplózást az End-Log cmdlet-tel kell bezárni.
    A Start-Log újbóli meghívása a naplózás lezárása nélkül hibát generál.
.PARAMETER FileName
    A naplófájl elérési útja (opcionális, ha a Console kapcsoló meg van adva)
.PARAMETER Clear
    A naplófájl ürítése a naplózás megkezdése előtt (opcionális)
    Alapértelmezésként ki van kapcsolva
.PARAMETER Console
    A naplóüzenetek kiírása a konzolra (opcionális)
    Alapértelmezésként ki van kapcsolva
.LINK
    End-Log
    Write-Log
    Clear-Log
    Set-LogLevel
    Get-LogLevel
    Get-Log-Encoding
    Set-Log-Encoding
.EXAMPLE
    Start-Log naplo.log -Clear
.NOTES
    Begin-Log
#>
function Start-Log
{
  param ([string] $FileName,
         [switch] $Clear = $false,
         [switch] $Console = $false)
  Consolidate-WorkingDir
  [BachRep.Log]::Start($FileName, $Clear, $Console)
}

Set-Alias -Name Begin-Log -Value Start-Log

<#
.SYNOPSIS
    Leállítja a naplózást.
.DESCRIPTION
    Leállítja a naplózást.
.LINK
    Start-Log
    Write-Log
    Set-LogLevel
    Get-LogLevel
    Get-Log-Encoding
    Set-Log-Encoding
.EXAMPLE
    End-Log
.NOTES
    Aliasok:
    Stop-Log
    Close-Log
#>
function End-Log
{
  [BachRep.Log]::End()
}

Set-Alias -Name Stop-Log -Value End-Log
Set-Alias -Name Close-Log -Value End-Log

<#
.SYNOPSIS
    Kiüríti az aktuális naplófájlt.
.DESCRIPTION
    Kiüríti az aktuális naplófájlt.
.LINK
    Start-Log
    End-Log
    Write-Log
    Set-LogLevel
    Get-LogLevel
    Get-Log-Encoding
    Set-Log-Encoding
.EXAMPLE
    Clear-Log
#>
function Clear-Log
{
  [BachReport.Log]::Clear()
}

<#
.SYNOPSIS
    Beír egy sort a naplóba.
.DESCRIPTION
    Beír egy sort a naplófájlba.
    Ha nincs elindítva a naplózás, akkor nem történik semmi.
.PARAMETER Message
    A naplózandó üzenet (kötelező)
.PARAMETER Source
    Az üzenet forrása (opcionális)
.PARAMETER Level
    Az üzenet naplózási szintje (opcionális). Lehetséges értékek: Normal, Verbose, Debug
    Alapértelmezett: Normal
    Lásd még: Set-LogLevel
.LINK
    Start-Log
    End-Log
    Clear-Log
    Set-LogLevel
    Get-LogLevel
    Get-Log-Encoding
    Set-Log-Encoding
.EXAMPLE
    Write-Log "MZ/X" "Köbüki" Verbose
#>
function Write-Log
{
  param ([parameter(Mandatory=$true,HelpMessage="A naplózandó üzenet")] [string] $Message,
         [string] $Source = "",
         [BachReport.LogLevel] $Level = [BachReport.LogLevel]::Normal)
  [BachReport.Log]::Write($Message, $Source, $Level)
}


<#
.SYNOPSIS
    Lekérdezi az érvényes naplózási szintet.
.DESCRIPTION
    Lekérdezi az érvényes naplózási szintet. Csak a beállított naplózási szintnek megfelelő üzenetek kerülnek bele a naplófájlba.
    A naplózási szintek: Normal, Verbose, Debug.
    Az egyes szintek jobbról balra tartalmazzák egymást.
    Az alapértelmezett naplózási szint: Normal.
.LINK
    Start-Log
    End-Log
    Clear-Log
    Write-Log
    Set-LogLevel
    Get-Log-Encoding
    Set-Log-Encoding
.EXAMPLE
    Get-LogLevel
    Normal
.EXAMPLE
    Set-LogLevel Verbose
    Get-LogLevel
    Verbose
.EXAMPLE
    Set-LogLevel Verbose,SQL
    Get-LogLevel
    Verbose, SQL
#>
function Get-LogLevel
{
  [BachReport.Log]::Level
}

Set-Alias -Name Get-Log-Level -Value Get-LogLevel


<#
.SYNOPSIS
    Beállítja a naplózási szintet.
.DESCRIPTION
    Beállítja a naplózási szintet. Csak a beállított naplózási szintnek megfelelő üzenetek kerülnek bele a naplófájlba.
    A naplózási szintek: Normal, Verbose, Debug, Error, SQL.
    Az egyes szintek kombinálhatók vesszővel elválasztva. A Debug szint az összes többi szintet tartalmazza.
    Az alapértelmezett naplózási szint: Normal.
    A naplózási szint csak globálisan állítható, és átállítás után érvényes.
    A BachReport cmdlet-jei önállóan is képesek naplózni a tevékenységüket, nem szükséges explicit naplózni az eljáráshívásokat.
.PARAMETER Level
    Az új naplózási szint (kötelező)
.LINK
    Start-Log
    End-Log
    Clear-Log
    Write-Log
    Get-LogLevel
    Get-Log-Encoding
    Set-Log-Encoding
.EXAMPLE
    Set-LogLevel Verbose
    Write-Log "Normál üzenet" #Bekerül a naplófájlba
    Write-Log "Verbose üzenet" -Level Verbose #Bekerül a naplófájlba
    Write-Log "Debug üzenet" -Level Debug #Nem kerül be a naplófájlba
.EXAMPLE
    Set-LogLevel SQL,Error #Kombinált naplózási szint
.NOTES
    Aliasok:
    Set-Log-Level
#>
function Set-LogLevel
{
  param ([parameter(Mandatory=$true, HelpMessage="Az új naplózási szint")] [BachReport.LogLevel] $Level)
  [BachReport.Log]::Level = $Level
}

Set-Alias -Name Set-Log-Level -Value Set-LogLevel


<#
.SYNOPSIS
    Lekérdezi a naplófájl kódolását.
.DESCRIPTION
    Lekérdezi a naplófájl kódolását.
    Lehetséges kódolások: ASCII, UTF-7, UTF-8, UTF-32, Unicode.
.LINK
    Start-Log
    End-Log
    Clear-Log
    Write-Log
    Get-LogLevel
    Set-LogLevel
    Set-Log-Encoding
    http://msdn.microsoft.com/en-us/library/system.text.encoding.aspx
.EXAMPLE
    Set-Log-Encoding Unicode
    Get-Log-Encoding
    Unicode
#>
function Get-Log-Encoding
{
  [BachReport.Log]::Encoding.EncodingName
}


<#
.SYNOPSIS
    Beállítja a naplófájl kódolását.
.DESCRIPTION
    Beállítja a naplófájl kódolását. A beállítás után a fájlba kiírt üzenetek a megadott kódolással kerülnek kiírásra.
    Lehetséges kódolások: ASCII, UTF-7, UTF-8, UTF-32, Unicode.
.PARAMETER Encoding
    A naplófájl új kódolása (kötelező)
    Lehetséges kódolások: ASCII, UTF-7, UTF-8, UTF-32, Unicode.
.LINK
    Start-Log
    End-Log
    Clear-Log
    Write-Log
    Get-LogLevel
    Set-LogLevel
    Get-Log-Encoding
    http://msdn.microsoft.com/en-us/library/system.text.encoding.aspx
.EXAMPLE
    Set-Log-Encoding Unicode
#>
function Set-Log-Encoding
{
  param ([parameter(Mandatory=$true, HelpMessage="A naplófájl kódolása")] [string] $Encoding)
  [BachReport.Log]::Encoding = [System.Text.Encoding]::GetEncoding($Encoding)
}


#Legutóbbi adatbázis-kapcsolat
Set-Variable -Name _lastDbConn -Scope Script

<#
.SYNOPSIS
    Adatbázis-kapcsolatot inicializál
.DESCRIPTION
    Adatbázis-kapcsolat inicializálása a megadott kapcsolatleíró sztringgel.
.PARAMETER ConnStr
    Kapcsolatleíró sztring (kötelező)
    Formátum: felhasználó/jelszó@adatbázis (MySql esetén nem használható)
    Alternatív formátum: http://www.connectionstrings.com
.PARAMETER Oracle
    Kapcsolódás Oracle adatbázishoz (alapértelmezett, nem kötelező megadni)
.PARAMETER SqlServer
    Kapcsolódás Sql Server adatbázishoz
.PARAMETER MySql
    Kapcsolódás MySql adatbázishoz
.OUTPUTS
    BachReport.IDb. Adatbázis-kapcsolat objektum
.LINK
    Db-Connection2
    Execute-Query
    Execute-Query-To-Csv
    Execute-Scalar
    Execute-Non-Query
    Dispose-Db
.EXAMPLE
    Db-Connect "user/pass@db"
    Adatbázis-kapcsolat: user@db
.EXAMPLE
    Db-Connect "user/pass@db" -SqlServer
    Adatbázis-kapcsolat: user@db
.EXAMPLE
    $db = Db-Connect "user/pass@db" #kapcsolat tárolása a $db változóban
.NOTES
    Aliasok:
    Db-Conn
    DbConn
#>
function Db-Connect
{
  param ([parameter(Mandatory=$true,HelpMessage="Kapcsolatleíró sztring")] [string] $ConnStr,
         [switch] $Oracle,
         [switch] $SqlServer,
         [switch] $MySql)
  if (($SqlServer -and $MySql -and $Oracle) -or ($MySql -and $SqlServer) -or ($MySql -and $Oracle) -or ($SqlServer -and $Oracle)) { throw "Az Oracle, SqlServer, MySql kapcsolók közül egyszerre csak egy használható." }
  if ($SqlServer)
  {
    $db = New-Object BachReport.Sql($ConnStr)
  }
  elseif ($MySql)
  {
    $db = New-Object BachReport.MySql($ConnStr)
  }
  else
  {
    $db = New-Object BachReport.Oracle($ConnStr)
  }
  Set-Variable -Name _lastDbConn -Value $db -Scope Script
  return $db
}

Set-Alias -Name Db-Conn -Value Db-Connect
Set-Alias -Name DbConn -Value Db-Connect

<#
.SYNOPSIS
    Adatbázis-kapcsolatot inicializál
.DESCRIPTION
    Adatbázis-kapcsolat inicializálása a megadott név, jelszó és adatbázis paraméterekkel.
.PARAMETER DataSource
    Az adatbázis azonosítója (kötelező)
.PARAMETER UserId
    A felhasználó azonosítója (kötelező)
.PARAMETER Password
    A felhasználó jelszava (kötelező)
.PARAMETER Server
    MySql kapcsolat esetén a szerver címe
.PARAMETER Port
    MySql kapcsolat esetén a szerver port száma (opcionális, alapértelmezett: 3306)
.PARAMETER Oracle
    Kapcsolódás Oracle adatbázishoz (alapértelmezett, nem kötelező megadni)
.PARAMETER SqlServer
    Kapcsolódás Sql Server adatbázishoz
.PARAMETER MySql
    Kapcsolódás MySql adatbázishoz
.OUTPUTS
    BachReport.IDb. Adatbázis-kapcsolat objektum
.LINK
    Db-Connection
    Execute-Query
    Execute-Query-To-Csv
    Execute-Scalar
    Execute-Non-Query
    Dispose-Db
.EXAMPLE
    Db-Connect "db" "user" "pass"
    Adatbázis-kapcsolat: user@db
.EXAMPLE
    Db-Connect "db" "user" "pass" "server" <port> -MySql
    Adatbázis-kapcsolat: user@db [server:port]
.EXAMPLE
    $db = Db-Connect "db" "user" "pass" #kapcsolat tárolása a $db változóban
.EXAMPLE
    $db = Db-Connect "db" "user" "pass" #kapcsolat tárolása a $db változóban
.NOTES
    Aliasok:
    Db-Conn2
    DbConn2
#>
function Db-Connect2
{
  param ([parameter(Mandatory=$true,HelpMessage="Az adatbázis azonosítója")] [string] $DataSource,
         [parameter(Mandatory=$true,HelpMessage="A felhasználó azonosítója")] [string] $UserId,
         [string] $Password = "",
         [string] $Server = "",
         [int]    $Port = 3306,
         [switch] $Oracle,
         [switch] $SqlServer,
         [switch] $MySql)
  if (($SqlServer -and $MySql -and $Oracle) -or ($MySql -and $SqlServer) -or ($MySql -and $Oracle) -or ($SqlServer -and $Oracle)) { throw "Az Oracle, SqlServer, MySql kapcsolók közül egyszerre csak egy használható." }
  if ($SqlServer)
  {
    $db = New-Object BachReport.Sql($DataSource, $UserId, $Password)
  }
  elseif ($MySql)
  {
    $db = New-Object BachReport.MySql($DataSource, $UserId, $Password, $Server, $Port)
  }
  else
  {
    $db = New-Object BachReport.Oracle($DataSource, $UserId, $Password)
  }
  Set-Variable -Name _lastDbConn -Value $db -Scope Script
  return $db
}

Set-Alias -Name Db-Conn2 -Value Db-Connect2
Set-Alias -Name DbConn2 -Value Db-Connect2

<#
.SYNOPSIS
    Lekérdezés futtatása egy megnyitott adatbázis-kapcsolaton
.DESCRIPTION
    Lekérdezés futtatása egy megnyitott adatbázis-kapcsolaton
.PARAMETER Sql
    Futtatandó lekérdezése (kötelező)
.PARAMETER DbConn
    Az adatbázis-kapcsolat, amelyen a lekérdezést le kell futtatni.
    Ha nincs megadva, akkor a legutóbbi Db-Connection vagy Db-Connection2 által megnyitott kapcsolat lesz az alapértelmezett
.OUTPUTS
    BachReport.Results. A lekérdezés eredményeit tartalmazó objektum
.LINK
    Db-Connection
    Db-Connection2
    Execute-Query-To-Csv
    Execute-Non-Query
    Execute-Scalar
    Dispose-Db
    Dispose-Results
.EXAMPLE
    Db-Connect "user/pass@db"
    $results = Execute-Query "select * from dual"
.EXAMPLE
    $db = Db-Connect2 "db" "user" "pass"
    $results = Execute-Query "select * from dual" $db
.EXAMPLE
    Db-Connect "db" "user" "pass"
    # Szkript betöltése fájlból
    $results = Execute-Query (Load-File "fájlnév.kit")
.NOTES
    Aliasok:
    Execute-SQL
    Db-Query
    Db-SQL
    Invoke-SQL
    Invoke-Query
#>
function Execute-Query
{
  param ([parameter(Mandatory=$true,HelpMessage="Futtatandó lekérdezés")] [string] $Sql,
         [BachReport.IDb] $DbConn = $_lastDbConn)
  if (!$DbConn) { throw "Nincs adatbázis-kapcsolat" }
  $DbConn.Query($Sql)
}

Set-Alias -Name Execute-SQL -Value Execute-Query
Set-Alias -Name Db-Query -Value Execute-Query
Set-Alias -Name Db-SQL -Value Execute-Query
Set-Alias -Name Invoke-Query -Value Execute-Query
Set-Alias -Name Invoke-SQL -Value Execute-Query

<#
.SYNOPSIS
    Lekérdezés futtatása egy megnyitott adatbázis-kapcsolaton és az eredmények mentése CSV fájlba
.DESCRIPTION
    Lekérdezés futtatása egy megnyitott adatbázis-kapcsolaton és az eredmények mentése CSV fájlba
.PARAMETER Sql
    Futtatandó lekérdezése (kötelező)
.PARAMETER DbConn
    Az adatbázis-kapcsolat, amelyen a lekérdezést le kell futtatni.
    Ha nincs megadva, akkor a legutóbbi Db-Connection vagy Db-Connection2 által megnyitott kapcsolat lesz az alapértelmezett
.PARAMETER OutputFile
    Kimeneti fájl neve
.PARAMETER Delimiter
    Értékelválasztó. Alapértelmezett: ;
.PARAMETER Quote
    Értékeket körülvevő karakterek. Alapértelmezett: "
.PARAMETER Encoding
    A kimeneti fájl kódolása. Alapértelmezett: UTF-8.
    Lehetséges kódolások: ASCII, UTF-7, UTF-8, UTF-32, Unicode.
.PARAMETER Append
    A kimeneti fájl folytatólagos írása (nem felülírja, hanem folytatja a kimeneti fájlt, ha már létezik)
.OUTPUTS
    System.Int32. A kiírt sorok száma. Hiba esetén -1.
.LINK
    Db-Connection
    Db-Connection2
    Execute-Non-Query
    Execute-Scalar
    Dispose-Db
    Dispose-Results
    http://msdn.microsoft.com/en-us/library/system.text.encoding.aspx
.EXAMPLE
    Db-Connect "user/pass@db"
    $results = Execute-Query-To-Csv "select * from dual" kimenet.csv -Delimiter "`t"
.EXAMPLE
    $db = Db-Connect2 "db" "user" "pass"
    $results = Execute-Query-To-Csv  "select * from dual" $db kimenet.csv
.EXAMPLE
    Db-Connect "db" "user" "pass"
    # Szkript betöltése fájlból
    Execute-Query-To-Csv (Load-File "fájlnév.kit") kimenet.csv -Delimiter "`t"
.NOTES
    Aliasok:
    Execute-SQL-To-Csv
    Db-Query-To-Csv
    Db-SQL-To-Csv
    Invoke-SQL-To-Csv
    Invoke-Query-To-Csv
    Query-To-Csv
    QueryToCsv
#>
function Execute-Query-To-Csv
{
  param ([parameter(Mandatory=$true,HelpMessage="Futtatandó lekérdezés")] [string] $Sql,
         [BachReport.IDb] $DbConn = $_lastDbConn,
         [parameter(Mandatory=$true,HelpMessage="Kimeneti fájlnév")] [string] $OutputFile,
         [string] [Alias('Separator')] $Delimiter = ";",
         [char]   $Quote = '"',
         [string] $Encoding = "UTF-8",
         [switch] $Append)
  if (!$DbConn) { throw "Nincs adatbázis-kapcsolat" }
  Consolidate-WorkingDir
  $DbConn.QueryToCsv($Sql, $OutputFile, $Delimiter, $Quote, [System.Text.Encoding]::GetEncoding($Encoding), $Append)
}

Set-Alias -Name Execute-SQL-To-Csv -Value Execute-Query-To-Csv
Set-Alias -Name Db-Query-To-Csv -Value Execute-Query-To-Csv
Set-Alias -Name Db-SQL-To-Csv -Value Execute-Query-To-Csv
Set-Alias -Name Invoke-Query-To-Csv -Value Execute-Query-To-Csv
Set-Alias -Name Invoke-SQL-To-Csv -Value Execute-Query-To-Csv
Set-Alias -Name Query-To-Csv -Value Execute-Query-To-Csv
Set-Alias -Name QueryToCsv -Value Execute-Query-To-Csv


<#
.SYNOPSIS
    Skalár lekérdezés futtatása egy megnyitott adatbázis-kapcsolaton
.DESCRIPTION
    Skalár lekérdezés futtatása egy megnyitott adatbázis-kapcsolaton
.PARAMETER Sql
    Futtatandó lekérdezése (kötelező)
.PARAMETER DbConn
    Az adatbázis-kapcsolat, amelyen a lekérdezést le kell futtatni.
    Ha nincs megadva, akkor a legutóbbi Db-Connection vagy Db-Connection2 által megnyitott kapcsolat lesz az alapértelmezett
.OUTPUTS
    System.Object. A lekérdezés eredményét tartalmazó objektum
.LINK
    Db-Connection
    Db-Connection2
    Execute-Non-Query
    Execute-Query
    Execute-Query-To-Csv
    Dispose-Db
    Dispose-Results
.EXAMPLE
    Db-Connect "user/pass@db"
    Execute-Query "select * from dual"
    X
.EXAMPLE
    $db = Db-Connect2 "db" "user" "pass"
    Execute-Scalar "select count(*) from dual" $db
    1
.EXAMPLE
    $db = Db-Connect "db" "user" "pass"
    # Szkript betöltése fájlból
    $results = Execute-Scalar (Load-File "fájlnév.kit")
.NOTES
    Aliasok:
    Invoke-Scalar
    Db-Scalar
#>
function Execute-Scalar
{
  param ([parameter(Mandatory=$true,HelpMessage="Futtatandó lekérdezés")] [string] $Sql,
         [BachReport.IDb] $DbConn = $_lastDbConn)
  if (!$DbConn) { throw "Nincs adatbázis-kapcsolat" }
  $DbConn.Scalar($Sql)
}

Set-Alias -Name Db-Scalar -Value Execute-Scalar
Set-Alias -Name Invoke-Scalar -Value Execute-Scalar


<#
.SYNOPSIS
    DML lekérdezés futtatása egy megnyitott adatbázis-kapcsolaton
.DESCRIPTION
    DML lekérdezés futtatása egy megnyitott adatbázis-kapcsolaton
.PARAMETER Sql
    Futtatandó lekérdezése (kötelező)
.PARAMETER DbConn
    Az adatbázis-kapcsolat, amelyen a lekérdezést le kell futtatni.
    Ha nincs megadva, akkor a legutóbbi Db-Connection vagy Db-Connection2 által megnyitott kapcsolat lesz az alapértelmezett
.OUTPUTS
    System.Int32. A lekérdezés által beszúrt vagy módosított sorok száma.
.LINK
    Db-Connection
    Db-Connection2
    Execute-Query
    Execute-Query-To-Csv
    Execute-Scalar
    Dispose-Db
    Dispose-Results
.EXAMPLE
    Db-Connect "user/pass@db"
    $results = Execute-Non-Query "insert into universe (answer) values (42)"
.EXAMPLE
    $db = Db-Connect2 "db" "user" "pass"
    $results = Execute-Query "update set universe status=null where name='Earth'" $db
.NOTES
  Aliasok:
  Execute-NonQuery
  Execute-DML
  Execute-DDL
  Db-Non-Query
  Db-NonQuery
  Db-DML
  Db-DDL
  Invoke-Non-Query
  Invoke-NonQuery
  Invoke-DML
  Invoke-DDL
#>
function Execute-Non-Query
{
  param ([parameter(Mandatory=$true,HelpMessage="Futtatandó lekérdezés")] [string] $Sql,
         [BachReport.IDb] $DbConn = $_lastDbConn)
  if (!$DbConn) { throw "Nincs adatbázis-kapcsolat" }
  $DbConn.NonQuery($Sql)
}

Set-Alias -Name Execute-NonQuery -Value Execute-Non-Query
Set-Alias -Name Execute-DML -Value Execute-Non-Query
Set-Alias -Name Execute-DDL -Value Execute-Non-Query
Set-Alias -Name Db-Non-Query -Value Execute-Non-Query
Set-Alias -Name Db-NonQuery -Value Execute-Non-Query
Set-Alias -Name Db-DML -Value Execute-Non-Query
Set-Alias -Name Db-DDL -Value Execute-Non-Query
Set-Alias -Name Invoke-Non-Query -Value Execute-Non-Query
Set-Alias -Name Invoke-NonQuery -Value Execute-Non-Query
Set-Alias -Name Invoke-DML -Value Execute-Non-Query
Set-Alias -Name Invoke-DDL -Value Execute-Non-Query

<#
.SYNOPSIS
    Bontja a megnyitott adatbázis-kapcsolatot
.DESCRIPTION
    Bontja a megnyitott adatbázis-kapcsolatot
.PARAMETER DbConn
    A bontandó adatbázis-kapcsolat
    Ha nincs megadva, akkor a legutóbbi Db-Connection vagy Db-Connection2 által megnyitott kapcsolat lesz az alapértelmezett
.LINK
    Db-Connection
    Db-Connection2
    Execute-Query
    Execute-Query-To-Csv
    Execute-Scalar
    Execute-Non-Query
    Dispose-Results
.EXAMPLE
    Dispose-Db $db
.NOTES
    Aliasok:
    Db-Dispose
#>
function Dispose-Db
{
  param ([BachReport.IDb] $DbConn = $_lastDbConn)
  if ($DbConn)
  {
    if ($DbConn -eq $_lastDbConn)
    {
      Clear-Variable -Name _lastDbConn -Scope Script
    }
    $DbConn.Dispose()
  }
}

Set-Alias -Name Db-Dispose -Value Dispose-Db

<#
.SYNOPSIS
    Felszabadítja az eredmény objektum által lefoglalt memóriát
.DESCRIPTION
    Felszabadítja a BachReport.Results típusú eredmény objektum által lefoglalt memóriát
.PARAMETER Results
    A felszabadítandó BachReport.Results típusú eredmény objektum
.LINK
    Db-Connection
    Db-Connection2
    Execute-Query
    Execute-Query-To-Csv
    Execute-Scalar
    Execute-Non-Query
    Dispose-Db
.EXAMPLE
    Dispose-Results $result
.NOTES
    Aliasok:
    Dispose-Result
#>
function Dispose-Results
{
  param([parameter(Mandatory=$true,HelpMessage="A felszabadítandó BachReport.Results típusú objektum")] [BachReport.Results] $Results)
  if ($Results) { $Results.Dispose() }
}

Set-Alias -Name Dispose-Result -Value Dispose-Results

#Legutóbbi Excel-munkafüzet
Set-Variable -Name _lastWorkbook -Scope Script

<#
.SYNOPSIS
    Új Excel-munkafüzet inicializálása
.DESCRIPTION
    Új Excel-munkafüzet inicializálása
.PARAMETER Template
    A megadott fájl megnyitása sablonként (opcionális)
.OUTPUTS
    BachReport.Excel. Excel-munkafüzet objektum
.LINK
    Add-Worksheet
    Add-Comment
    Set-Worksheet-Styles
    Set-Workbook-Properties
    Save-Workbook
.EXAMPLE
    Create-Workbook
.EXAMPLE
    Create-Workbook "sablon.xlsx"
.NOTES
    Aliasok:
    Create-Excel
    Init-Excel
    Init-Workbook
#>
function Create-Workbook
{
  param([string] $Template)
  Consolidate-WorkingDir
  $workbook = New-Object BachReport.Excel($Template)
  Set-Variable -Name _lastWorkbook -Value $workbook -Scope Script
  return $workbook
}

Set-Alias -Name Create-Excel -Value Create-Workbook
Set-Alias -Name Init-Excel -Value Create-Workbook
Set-Alias -Name Init-Workbook -Value Create-Workbook

<#
.SYNOPSIS
    Új munkalap beszúrása egy Excel-munkafüzetbe
.DESCRIPTION
    Új munkalap beszúrása egy Excel-munkafüzetbe
.PARAMETER Results
    A munkalapként beszúrandó eredményhalmaz (kötelező)
.PARAMETER Name
    A beszúrandó munkalap neve.
    Ha már létezik a munkalap, mert már létezett a megnyitott sablonfájlban vagy korábban be lett szúrva, akkor az felülírásra kerül.
.PARAMETER Workbook
    Az Excel-munkafüzet, amelybe a munkalapot be kell szúrni (opcionális)
    Ha nincs megadva, akkor a legutóbbi Create-Workbook által létrehozott munkafüzet lesz az alapértelmezett
.LINK
    Create-Workbook
    Add-Comment
    Set-Worksheet-Styles
    Set-Workbook-Properties
    Save-Workbook
.EXAMPLE
    Add-Worksheet $results "Eredmények"
.EXAMPLE
    Add-Worksheet $results "Eredmények" $workbook
.NOTES
    Aliasok:
    Add-Result
    Add-Results
#>
function Add-Worksheet
{
  param([parameter(Mandatory=$true,HelpMessage="A beszúrandó eredményhalmaz")] [BachReport.Results] $Results,
        [parameter(Mandatory=$true,HelpMessage="A beszúrandó munkalap neve")] [string] $Name,
        [BachReport.Excel] $Workbook = $_lastWorkbook)
  if (!$Workbook) { throw "Nincs megadva munkafüzet" }
  $Workbook.AddWorksheet($Results, $Name)
}

Set-Alias -Name Add-Result -Value Add-Worksheet
Set-Alias -Name Add-Results -Value Add-Worksheet


<#
.SYNOPSIS
    Megjegyzés beszúrása egy Excel cellába
.DESCRIPTION
    Megjegyzés beszúrása egy kiválasztott munkalap kiválasztott cellájába
.PARAMETER Worksheet
    A kiválasztott munkalap neve
    Ha nem létezik a munkalap, akkor a megjegyzés nem kerül beszúrásra
.PARAMETER Text
    A megjegyzés szövege
.PARAMETER Author
    A megjegyzés szerzője
.PARAMETER Row
    A kiválasztott cella sora. Alapértelmezett: 1
.PARAMETER Col
    A kiválasztott cella oszlopa. Alapértelmezett: 1
.PARAMETER Workbook
    Az Excel-munkafüzet, amelybe a munkalapot be kell szúrni (opcionális)
    Ha nincs megadva, akkor a legutóbbi Create-Workbook által létrehozott munkafüzet lesz az alapértelmezett
.LINK
    Create-Workbook
    Add-Worksheet
    Set-Worksheet-Styles
    Set-Workbook-Properties
    Save-Workbook
.EXAMPLE
    Add-Comment "Eredmények" "Megjegyzés" "Szerző"
.EXAMPLE
    Add-Comment "Eredmények" "Megjegyzés 5. sor 10. oszlop" "Szerző" 5 10
.NOTES
    Aliasok:
    Add-Comments
#>
function Add-Comment
{
  param([parameter(Mandatory=$true,HelpMessage="Munkalap neve")] [string] $Worksheet,
        [parameter(Mandatory=$true,HelpMessage="Megjegyzés szövege")] [string] $Text,
        [parameter(Mandatory=$true,HelpMessage="Megjegyzés szerzője")] [string] $Author,
        [int] $Row = 1,
        [int] $Col = 1,
        [BachReport.Excel] $Workbook = $_lastWorkbook)
  if (!$Workbook) { throw "Nincs megadva munkafüzet" }
  $Workbook.AddComment($Worksheet, $Text, $Author, $Row, $Col)
}

Set-Alias -Name Add-Comments -Value Add-Comment

<#
.SYNOPSIS
    Excel-munkafüzet mentése
.DESCRIPTION
    Excel-munkafüzet mentése XLSX formátumba
.PARAMETER FileName
    Mentési fájlnév (kötelező)
.PARAMETER Dispose
    A beszúrt eredmény objektumok felszabadítása (opcionális).
    Alapértelmezés: $true
.PARAMETER Workbook
    A mentendő Excel-munkafüzet (opcionális)
    Ha nincs megadva, akkor a legutóbbi Create-Workbook által létrehozott munkafüzet lesz az alapértelmezett
.LINK
    Create-Workbook
    Add-Worksheet
    Add-Comment
    Set-Worksheet-Styles
    Set-Workbook-Properties
.EXAMPLE
    Save-Workbook "eredmény.xlsx"
.EXAMPLE
    Save-Workbook "eredmény.xlsx" -Workbook $workbook
.NOTES
    Aliasok:
    Save-Excel
#>
function Save-Workbook
{
  param([parameter(Mandatory=$true,HelpMessage="Mentési fájlnév")] [string] $FileName,
        [bool] $Dispose = $true,
        [BachReport.Excel] $Workbook = $_lastWorkbook)
  if (!$Workbook) { throw "Nincs megadva munkafüzet" }
  if ($Workbook -eq $_lastWorkbook)
  {
    Clear-Variable -Name _lastWorkbook -Scope Script
  }
  Consolidate-WorkingDir
  $Workbook.Save($FileName, $Dispose)
}

Set-Alias -Name Save-Excel -Value Save-Workbook

<#
.SYNOPSIS
    Beállítja a következőleg beszúrt munkalap generálási tulajdonságait
.DESCRIPTION
    Beállítja a következőleg beszúrt munkalap generálási tulajdonságait
.PARAMETER AutoFilter
    Autoszűrő bekapcsolása a fejléc oszlopaira (opcionális)
    Alapértelmezés: $true
.PARAMETER AutoFit
    Az oszlopok automatikus szélességének beállítása az adatok beszúrása után (opcionális)
    Alapértelmezés: $true
.PARAMETER Header
    Fejlécek beszúrása (opcionális)
    Alapértelmezés: $true
.PARAMETER HeaderBold
    Fejlécek félkövér stílusú (opcionális)
    Alapértelmezés: $true
.PARAMETER FreezePanes
    A fejléc sor rögzítése (opcionális)
    Alapértelmezés: $true
    A rögzítés az adatok beszúrási helyétől indul, alapértelmezésként az első sor rögzítésével egyenértékű
.PARAMETER HeaderTextRotation
    Fejlécek szöveg elforgatása (opcionális)
    Lehetséges értékek: 0-180
    Alapértelmezés: 0
.PARAMETER HeaderStartRow
    A fejléc beszúrása ettől a sortól (opcionális)
    Alapértelmezés: 1
.PARAMETER HeaderStartCol
    A fejléc beszúrása ettől az oszloptól (opcionális)
    Alapértelmezés: 1
.PARAMETER DataStartRow
    Az adatok beszúrása ettől a sortól (opcionális)
    Alapértelmezés: 2
.PARAMETER DataStartCol
    Az adatok beszúrása ettől az oszloptól (opcionális)
.PARAMETER DefaultColWidth
    Alapértelmezett oszlopszélesség (opcionális)
    Alapértelmezés: 0 (nincs megadva)
.PARAMETER DefaultRowHeight
    Alapértelmezett sormagasság (opcionális)
    Alapértelmezés: 0 (nincs megadva)
.PARAMETER MinColWidth
    Minimális oszlopszélesség automatikus szélesség beállítása után (opcionális)
    Alapértelmezés: 0 (nincs megadva)
.PARAMETER MaxColWidth
    Maximális oszlopszélesség automatikus szélesség beállítása után (opcionális)
    Alapértelmezés: 0 (nincs megadva)
.PARAMETER DateFormat
    Dátumformátum (opcionális)
    Alapértelmezés: yyyy.mm.dd
.PARAMETER Reset
    Tulajdonságok beállítása az alapértlemezett értékekre (a többi paraméter beállítása előtt)
.LINK
    Create-Workbook
    Add-Worksheet
    Add-Comment
    Set-Workbook-Properties
    Save-Workbook
.EXAMPLE
    Set-Worksheet-Styles -AutoFit $true -AutoFilter $true -FreezePanes $false -HeaderStartRow 3 -DataStartRow 4
    Add-Worksheet $results "Munkalap"
    # Autoszűrő be, automatikus oszlopszélesség be, felső sor rögzítése ki, fejléc 3. sorban, adatok 4. sortól
.NOTES
    Aliasok:
    Set-Worksheet-Style
#>
function Set-Worksheet-Styles
{
  param([bool] $AutoFilter,
        [bool] $AutoFit,
        [bool] $Header,
        [bool] $FreezePanes,
        [bool] $HeaderBold,
        [ValidateRange(0,180)] [int] $HeaderTextRotation,
        [ValidateRange(1,1048576)] [int] $HeaderStartRow,
        [ValidateRange(1,16384)] [int] $HeaderStartCol,
        [ValidateRange(1,1048576)] [int] $DataStartRow,
        [ValidateRange(1,16384)] [int] $DataStartCol,
        [ValidateRange(0,255)] [double] $DefaultColWidth,
        [ValidateRange(0,409)] [double] $DefaultRowHeight,
        [ValidateRange(0,255)] [double] $MinColWidth,
        [ValidateRange(0,255)] [double] $MaxColWidth,
        [string] $DateFormat,
        [switch] $Reset)
  if ($Reset) { [BachReport.Excel]::ResetStyles() }
  if ($PSBoundParameters.ContainsKey("AutoFilter")) { [BachReport.Excel]::AutoFilter = $AutoFilter }
  if ($PSBoundParameters.ContainsKey("AutoFit")) { [BachReport.Excel]::AutoFit = $AutoFit }
  if ($PSBoundParameters.ContainsKey("Header")) { [BachReport.Excel]::Header = $Header }
  if ($PSBoundParameters.ContainsKey("HeaderBold")) { [BachReport.Excel]::HeaderBold = $HeaderBold }
  if ($PSBoundParameters.ContainsKey("HeaderTextRotation")) { [BachReport.Excel]::HeaderTextRotation = $HeaderTextRotation }
  if ($PSBoundParameters.ContainsKey("FreezePanes")) { [BachReport.Excel]::FreezePanes = $FreezePanes }
  if ($PSBoundParameters.ContainsKey("HeaderStartRow")) { [BachReport.Excel]::HeaderStartRow = $HeaderStartRow }
  if ($PSBoundParameters.ContainsKey("HeaderStartCol")) { [BachReport.Excel]::HeaderStartCol = $HeaderStartCol }
  if ($PSBoundParameters.ContainsKey("DataStartRow")) { [BachReport.Excel]::DataStartRow = $DataStartRow }
  if ($PSBoundParameters.ContainsKey("DataStartCol")) { [BachReport.Excel]::DataStartCol = $DataStartCol }
  if ($PSBoundParameters.ContainsKey("DefaultColWidth")) { [BachReport.Excel]::DataStartRow = $DefaultColWidth }
  if ($PSBoundParameters.ContainsKey("DefaultRowHeight")) { [BachReport.Excel]::DataStartCol = $DefaultRowHeight }
  if ($PSBoundParameters.ContainsKey("MinColWidth")) { [BachReport.Excel]::MinColWidth = $MinColWidth }
  if ($PSBoundParameters.ContainsKey("MaxColWidth")) { [BachReport.Excel]::MaxColWidth = $MaxColWidth }
  if ($PSBoundParameters.ContainsKey("DateFormat")) { [BachReport.Excel]::DateFormat = $DateFormat }
}

Set-Alias -Name Set-Worksheet-Style -Value Set-Worksheet-Styles

<#
.SYNOPSIS
    Beállítja a generált munkafüzet tulajdonságait
.DESCRIPTION
    Beállítja a generált munkafüzet tulajdonságait.
    A tulajdonságokat mentés előtt lehet beállítani
.PARAMETER Author
    A dokumentum szerzője (opcionális)
.PARAMETER Category
    A dokumentum kategóriája (opcionális)
.PARAMETER Comments
    A dokumentum megjegyzés mezője (opcionális)
.PARAMETER Company
    A dokumentum cégnév mezője (opcionális)
.PARAMETER Keywords
    A dokumentum kulcsszavai (opcionális)
.PARAMETER LastModifiedBy
    A dokumentum utolsó módosítója (opcionális)
.PARAMETER LastPrinted
    A dokumentumot utoljára kinyomtatta (opcionális)
.PARAMETER Manager
    A dokumentum felelőse (opcionális)
.PARAMETER Status
    A dokumentum állapota (opcionális)
.PARAMETER Subject
    A dokumentum tárgya (opcionális)
.PARAMETER Title
    A dokumentum címe (opcionális)
.PARAMETER Workbook
    A módosítandó Excel-munkafüzet objektum (opcionális)
    Ha nincs megadva, akkor a legutóbbi Create-Workbook által létrehozott munkafüzet lesz az alapértelmezett
.LINK
    Create-Workbook
    Add-Worksheet
    Add-Comment
    Set-Worksheet-Styles
    Save-Workbook
.EXAMPLE
    Set-Workbook-Properties -Title "Cím" -Author "Szerző" -Felelős "Főnök"
.NOTES
    Aliasok:
    Set-Workbook-Prop
    Set-Workbook-Props
#>
function Set-Workbook-Properties
{
  param([string] $Author,
        [string] $Category,
        [string] $Comments,
        [string] $Company,
        [string] $Keywords,
        [string] $LastModifiedBy,
        [string] $LastPrinted,
        [string] $Manager,
        [string] $Status,
        [string] $Subject,
        [string] $Title,
        [BachReport.Excel] $Workbook = $_lastWorkbook)
  if (!$Workbook) { throw "Nincs megadva munkafüzet" }
  if ($PSBoundParameters.ContainsKey("Author"))         { $Workbook.Author         = $Author }
  if ($PSBoundParameters.ContainsKey("Category"))       { $Workbook.Category       = $Category }
  if ($PSBoundParameters.ContainsKey("Comments"))       { $Workbook.Comments       = $Comments }
  if ($PSBoundParameters.ContainsKey("Company"))        { $Workbook.Company        = $Company }
  if ($PSBoundParameters.ContainsKey("Keywords"))       { $Workbook.Keywords       = $Keywords }
  if ($PSBoundParameters.ContainsKey("LastModifiedBy")) { $Workbook.LastModifiedBy = $LastModifiedBy }
  if ($PSBoundParameters.ContainsKey("LastPrinted"))    { $Workbook.LastPrinted    = $LastPrinted }
  if ($PSBoundParameters.ContainsKey("Manager"))        { $Workbook.Manager        = $Manager }
  if ($PSBoundParameters.ContainsKey("Status"))         { $Workbook.Status         = $Status }
  if ($PSBoundParameters.ContainsKey("Subject"))        { $Workbook.Subject        = $Subject }
  if ($PSBoundParameters.ContainsKey("Title"))          { $Workbook.Title          = $Title }
}

Set-Alias -Name Set-Workbook-Prop -Value Set-Workbook-Properties
Set-Alias -Name Set-Workbook-Props -Value Set-Workbook-Properties

<#
.SYNOPSIS
    Email küldése
.DESCRIPTION
    Email küldése
.PARAMETER From
    Feladó (kötelező)
.PARAMETER To
    Címzett (kötelező)
.PARAMETER Subject
    Tárgy (kötelező)
.PARAMETER Body
    Üzenettörzs (kötelező)
.PARAMETER Attachments
    Csatolmányok (opcionális)
.PARAMETER Cc
    Másolatot kap (opcionális)
.PARAMETER Bcc
    Titkos másolat (opcionális)
.PARAMETER Priority
    Prioritás (opcionális)
.PARAMETER Html
    Html küldése (opcionális)
.EXAMPLE
    Send-Mail "felado@telekom.hu" "cimzett@telekom.hu" "Tárgy" "Üzenettörzs"
.EXAMPLE
    Send-Mail "felado@telekom.hu" "cimzett@telekom.hu" "Tárgy" "Üzenettörzs" "csatolmány.kit"
.EXAMPLE
    Send-Mail "felado@telekom.hu" "cimzett@telekom.hu" "Tárgy" "Üzenettörzs" "csatolmány1.kit","csatolmány2.kit","csatolmány3.kit"
.EXAMPLE
    Send-Mail "felado@telekom.hu" "cimzett@telekom.hu" "Tárgy" "Üzenettörzs" "csatolmány.kit" "masolat@telekom.hu"
.EXAMPLE
    Send-Mail "felado@telekom.hu" "cimzett@telekom.hu" "Tárgy" "Üzenettörzs" "csatolmány.kit" "masolat@telekom.hu" "titkos@telekom.hu"
.EXAMPLE
    Send-Mail "felado@telekom.hu" "cimzett@telekom.hu" "Tárgy" "Üzenettörzs" -Cc "masolat@telekom.hu" -Bcc "titkos@telekom.hu"
.LINK
    Get-Mail-Body-Encoding
    Set-Mail-Body-Encoding
    Get-Mail-Subject-Encoding
    Set-Mail-Subject-Encoding
.NOTES
    Aliasok:
    Send-Email
    SendMail
    EmailSend
    Mail-Send
    MailSend
#>
function Send-Mail
{
  param([parameter(Mandatory=$true,HelpMessage="Feladó")] [string] $From,
        [parameter(Mandatory=$true,HelpMessage="Címzett")] [string] $To,
        [parameter(Mandatory=$true,HelpMessage="Tárgy")] [string] $Subject,
        [parameter(Mandatory=$true,HelpMessage="Üzenettörzs")] [string] $Body,
        [string[]] $Attachments,
        [string] $Cc,
        [string] $Bcc,
        [System.Net.Mail.MailPriority] $Priority = [System.Net.Mail.MailPriority]::Normal,
        [string] $SmtpHost,
        [int] $SmtpPort = 25,
        [switch] $Html
        )
  if ($Attachments) {
    Consolidate-WorkingDir
    [BachReport.Mail]::Send($From, $To, $Cc, $Bcc, $Subject, $Body, $Priority, $Html, $SmtpHost, $SmtpPort, $Attachments)
  } else {
    [BachReport.Mail]::Send($From, $To, $Cc, $Bcc, $Subject, $Body, $Priority, $Html, $SmtpHost, $SmtpPort)
  }
}

Set-Alias -Name Send-Email -Value Send-Mail
Set-Alias -Name SendMail -Value Send-Mail
Set-Alias -Name EmailSend -Value Send-Mail
Set-Alias -Name Mail-Send -Value Send-Mail
Set-Alias -Name MailSend -Value Send-Mail


<#
.SYNOPSIS
    Lekérdezi az elküldött emailek üzenettörzsének kódolását.
.DESCRIPTION
    Lekérdezi az elküldött emailek üzenettörzsének kódolását.
    Lehetséges kódolások: ASCII, UTF-7, UTF-8, UTF-32, Unicode.
.LINK
    Send-Mail
    Set-Mail-Body-Encoding
    Get-Mail-Subject-Encoding
    Set-Mail-Subject-Encoding
    http://msdn.microsoft.com/en-us/library/system.text.encoding.aspx
.EXAMPLE
    Set-Mail-Body-Encoding Unicode
    Get-Mail-Body-Encoding
    Unicode
.NOTES
    Az alapértelmezett érték: UTF-8
#>
function Get-Mail-Body-Encoding
{
  [BachReport.Mail]::BodyEncoding.EncodingName
}


<#
.SYNOPSIS
    Beállítja az elküldött emailek üzenettörzsének kódolását.
.DESCRIPTION
    Beállítja az elküldött emailek üzenettörzsének kódolását. A beállítás után az elküldött üzenetek fejlécében az itt megadott kódolás lesz megjelölve a üzenettörzshöz.
    Lehetséges kódolások: ASCII, UTF-7, UTF-8, UTF-32, Unicode.
.PARAMETER Encoding
    Az emailek üzenettörzsének új kódolása (kötelező)
    Lehetséges kódolások: ASCII, UTF-7, UTF-8, UTF-32, Unicode.
.LINK
    Send-Mail
    Get-Mail-Body-Encoding
    Get-Mail-Subject-Encoding
    Set-Mail-Subject-Encoding
    http://msdn.microsoft.com/en-us/library/system.text.encoding.aspx
.EXAMPLE
    Set-Mail-Body-Encoding Unicode
.NOTES
    Az alapértelmezett érték: UTF-8
#>
function Set-Mail-Body-Encoding
{
  param ([parameter(Mandatory=$true, HelpMessage="Az emailek üzenettörzsének kódolása")] [string] $Encoding)
  [BachReport.Mail]::BodyEncoding = [System.Text.Encoding]::GetEncoding($Encoding)
}


<#
.SYNOPSIS
    Lekérdezi az elküldött emailek tárgyának kódolását.
.DESCRIPTION
    Lekérdezi az elküldött emailek tárgyának kódolását.
    Lehetséges kódolások: ASCII, UTF-7, UTF-8, UTF-32, Unicode.
.LINK
    Send-Mail
    Get-Mail-Body-Encoding
    Set-Mail-Body-Encoding
    Set-Mail-Subject-Encoding
    http://msdn.microsoft.com/en-us/library/system.text.encoding.aspx
.EXAMPLE
    Set-Mail-Subject-Encoding Unicode
    Get-Mail-Subject-Encoding
    Unicode
.NOTES
    Az alapértelmezett érték: UTF-8
#>
function Get-Mail-Subject-Encoding
{
  [BachReport.Mail]::SubjectEncoding.EncodingName
}


<#
.SYNOPSIS
    Beállítja az elküldött emailek tárgyának kódolását.
.DESCRIPTION
    Beállítja az elküldött emailek tárgyának a kódolását. A beállítás után az elküldött üzenetek fejlécében az itt megadott kódolással lesz feltűntetve az üzenet tárgya.
    Lehetséges kódolások: ASCII, UTF-7, UTF-8, UTF-32, Unicode.
.PARAMETER Encoding
    Az üzenettörzsek új kódolása (kötelező)
    Lehetséges kódolások: ASCII, UTF-7, UTF-8, UTF-32, Unicode.
.LINK
    Send-Mail
    Get-Mail-Body-Encoding
    Get-Mail-Subject-Encoding
    http://msdn.microsoft.com/en-us/library/system.text.encoding.aspx
.EXAMPLE
    Set-Mail-Subject-Encoding Unicode
.NOTES
    Az alapértelmezett érték: UTF-8
#>
function Set-Mail-Subject-Encoding
{
  param ([parameter(Mandatory=$true, HelpMessage="Az emailek tárgyának kódolása")] [string] $Encoding)
  [BachReport.Mail]::SubjectEncoding = [System.Text.Encoding]::GetEncoding($Encoding)
}


<#
.SYNOPSIS
    Fájl betöltése sztringként (sorok gyűjteménye helyett)
.DESCRIPTION
    Fájl betöltése sztringként (sorok gyűjteménye helyett).
    A cmdlet egy rövidítése az alábbi hívásnak:
    [string]::Join([Environment]::NewLine, (Get-Content $FileName))
.PARAMETER FileName
    A betöltendő fájl elérési útvonala (kötelező)
.LINK
    http://msdn.microsoft.com/en-us/library/system.text.encoding.aspx
.EXAMPLE
    Load-File "fájlnév.kit"
.EXAMPLE
    Execute-Query (Load-File "szkript.sql")
#>
function Load-File
{
  param([parameter(Mandatory=$true,HelpMessage="Fájlnév")] [string] $FileName)
  $content = Get-Content $FileName;
  if (!$content) { return ""; }
  [string]::Join([Environment]::NewLine, $content)
}


<#
.SYNOPSIS
    Munkakönyvtár szinkronizálása
.DESCRIPTION
    Szinkronizálja a munkakönyvtárat a PowerShell és a .NET keretrendszer között
    Belső használatra
#>
function Consolidate-WorkingDir
{
    # Munkakönyvtár beállítása
    [Environment]::CurrentDirectory = (Get-Location -PSProvider FileSystem).ProviderPath
}


<#
.SYNOPSIS
    Lekérdezés átalakítása pivottá
.DESCRIPTION
    Lekérdezés átalakítása pivottá
.PARAMETER Sql
    A pivottá alakítandó lekérdezés
.PARAMETER FilterCol
    Szűrő oszlop
.PARAMETER ValueCol
    Érték oszlop
.PARAMETER Function
    A csoportosító függvény, pl. SUM, MIN, MAX. Alapértelmezett: SUM
.PARAMETER ColPrefix
    Az aggregált oszlopok nevének előtagja
.PARAMETER ColPostfix
    Az aggregált oszlopok nevének utótagja
.PARAMETER FallbackValue
    Az üres sorok értéke az aggregációnál. Alapértelmezett: 0
.PARAMETER DbConn
    Az adatbázis-kapcsolat, amelyen a lekérdezést le kell futtatni.
    Ha nincs megadva, akkor a legutóbbi Db-Connection vagy Db-Connection2 által megnyitott kapcsolat lesz az alapértelmezett
.OUTPUTS
    String. A pivottá alakított lekérdezés szövegként
.EXAMPLE
    Make-Pivot (Load-File "szkript.sql") Szuro Ertek
.NOTES
    Aliasok:
    Create-Pivot
#>
function Make-Pivot
{
  param ([parameter(Mandatory=$true,HelpMessage="A pivottá alakítandó lekérdezés")] [string] $Sql,
         [parameter(Mandatory=$true,HelpMessage="Szűrő oszlop")] [string] $FilterCol,
         [parameter(Mandatory=$true,HelpMessage="Érték oszlop")] [string] $ValueCol,
         [string] $Function = "SUM",
         [string] $ColPrefix = "",
         [string] $ColPostfix = "",
         [string] $FallbackValue = 0,
         [BachReport.IDb] $DbConn = $_lastDbConn)
  # Adatbázis-kapcsolat ellenőrzése
  if (!$DbConn) { throw "Nincs adatbázis-kapcsolat" }

  # Lekérdezés oszlopainak lekérdezése
  $columns = (Execute-Query -Sql ("select * from ({0}) where 1=2" -f $Sql) -DbConn $DbConn).Columns
  if (!$columns) { throw "A lekérdezés hibára futott" }

  # A szűrő oszlop distinct értékeinek lekérdezése
  $filterVals = (Execute-Query -Sql ("select distinct {0} from ({1})" -f $filterCol,$Sql) -DbConn $DbConn)
  if ($filterVals.IsEmpty) { throw "A szűrő oszlopban nincsenek értékek" }

  # Csoportosító oszlopok kigyűjtése a szűrő és az érték oszlopok kihagyásával
  $groupColumns = @()
  foreach ($column in $columns)
  {
    if ($column -ne $FilterCol -and $column -ne $ValueCol)
    {
      $groupColumns += $column
    }
  }

  # Aggregáló oszlopok előállítása
  $aggregatedColumns = @()
  foreach ($filterVal in $filterVals.Data)
  {
      $aggregatedColumns += ("{0}(decode({1},'{2}',{3},'{4}')) `"{5}{6}{7}`"" -f $Function, $FilterCol, $filterVal, $ValueCol, $FallbackValue, $ColPrefix, $filterVal, $ColPostfix)
  }
  return "with internal_query as ({0})`nselect {1}`n,{2}`nfrom internal_query`ngroup by {1}" -f $Sql.Trim(),[string]::Join(",", $groupColumns),[string]::Join("`n,", $aggregatedColumns)
}

Set-Alias -Name Create-Pivot -Value Make-Pivot

<#
.SYNOPSIS
    XLSX fájl átalakítása CSV formátumra
.DESCRIPTION
    XLSX fájl átalakítása CSV formátumra
.PARAMETER InputFile
    Bemeneti fájl neve
.PARAMETER OutputFile
    Kimeneti fájl neve
.PARAMETER WorksheetId
    Kiválasztott munkalap sorszáma, 1-től sorszámozva. Alapértelmezett: 1
.PARAMETER WorksheetName
    Kiválasztott munkalap neve (felülbírálja a WorksheetId paramétert, opcionális)
.PARAMETER Delimiter
    Értékelválasztó. Alapértelmezett: ;
.PARAMETER Quote
    Értékeket körülvevő karakterek. Alapértelmezett: "
.PARAMETER Encoding
    A kimeneti fájl kódolása. Alapértelmezett: UTF-8.
    Lehetséges kódolások: ASCII, UTF-7, UTF-8, UTF-32, Unicode.
.PARAMETER Password
    A bemeneti fájl jelszava (opcionális)
.PARAMETER Append
    A kimeneti fájl folytatólagos írása (nem felülírja, hanem folytatja a kimeneti fájlt, ha már létezik)
.PARAMETER Transform
    Értéktranszformációk
.EXAMPLE
    Xlsx-To-Csv bemenet.xlsx kimenet.csv -Delimiter ","
.LINK
    http://msdn.microsoft.com/en-us/library/system.text.encoding.aspx
.NOTES
    Aliasok:
    Excel-To-Csv
    Convert-To-Csv
#>
function Xlsx-To-Csv
{
  param ([parameter(Mandatory=$true,HelpMessage="Bemeneti fájlnév")] [string] $InputFile,
         [parameter(Mandatory=$true,HelpMessage="Kimeneti fájlnév")] [string] $OutputFile,
         [int] $WorksheetId = 1,
         [string] $WorksheetName,
         [string] [Alias('Separator')] $Delimiter = ";",
         [char]   $Quote = '"',
         [string] $Encoding = "UTF-8",
         [string] $Password,
         [switch] $Append,
         [HashTable] $Transform)
  Consolidate-WorkingDir
  [BachReport.Excel]::ConvertToCsv($InputFile, $OutputFile, $WorksheetId, $WorksheetName, $Delimiter, $Quote, [System.Text.Encoding]::GetEncoding($Encoding), $Password, $Append, $Transform)
}

Set-Alias -Name Excel-To-Csv -Value Xlsx-To-Csv
Set-Alias -Name Convert-To-Csv -Value Xlsx-To-Csv


<#
.SYNOPSIS
    Eredményhalmaz kiírása CSV formátumba
.DESCRIPTION
    Eredményhalmaz kiírása CSV formátumba
.PARAMETER OutputFile
    Kimeneti fájl neve
.PARAMETER Delimiter
    Értékelválasztó. Alapértelmezett: ;
.PARAMETER Quote
    Értékeket körülvevő karakterek. Alapértelmezett: "
.PARAMETER Encoding
    A kimeneti fájl kódolása. Alapértelmezett: UTF-8.
    Lehetséges kódolások: ASCII, UTF-7, UTF-8, UTF-32, Unicode.
.PARAMETER Append
    A kimeneti fájl folytatólagos írása (nem felülírja, hanem folytatja a kimeneti fájlt, ha már létezik)
.EXAMPLE
    Save-To-Csv $Results kimenet.csv -Delimiter ","
.LINK
    http://msdn.microsoft.com/en-us/library/system.text.encoding.aspx
.NOTES
    Aliasok:
    Results-To-Csv
#>
function Save-To-Csv
{
  param ([parameter(Mandatory=$true,HelpMessage="Az elmentendő eredményhalmaz")] [BachReport.Results] $Results,
         [parameter(Mandatory=$true,HelpMessage="Kimeneti fájlnév")] [string] $OutputFile,
         [string] [Alias('Separator')] $Delimiter = ";",
         [char]   $Quote = '"',
         [string] $Encoding = "UTF-8",
         [switch] $Append)
  Consolidate-WorkingDir
  $Results.SaveToCsv($OutputFile, $Delimiter, $Quote, [System.Text.Encoding]::GetEncoding($Encoding), $Append)
}

Set-Alias -Name Results-To-Csv -Value Save-To-Csv


<#
.SYNOPSIS
    Eredményhalmaz exportálása HTML formátumra
.DESCRIPTION
    Eredményhalmaz exportálása HTML formátumra
.PARAMETER InputFile
    Az exportálandó eredményhalmaz
.PARAMETER TableStyle
    A <table> címkék style attribútuma
.PARAMETER TrStyle
    A <tr> címkék style attribútuma
.PARAMETER TdStyle
    A <td> címkék style attribútuma
.PARAMETER ThStyle
    A <th> címkék style attribútuma
.PARAMETER TheadStyle
    A <thead> címkék style attribútuma
.PARAMETER TbodyStyle
    A <tbody> címkék style attribútuma
.OUTPUTS
    String. Az eredménytáblázat HTML táblázatként formázva
.EXAMPLE
    Export-To-Html $Results
#>
function Export-To-Html
{
  param([parameter(Mandatory=$true,HelpMessage="Az exportálandó eredményhalmaz")] [BachReport.Results] $Results,
         [string] $TableStyle,
         [string] $TrStyle,
         [string] $TdStyle,
         [string] $ThStyle,
         [string] $TheadStyle,
         [string] $TbodyStyle)
  if ($Results.IsEmpty) { return "" }
  $sb = New-Object System.Text.StringBuilder("<table")
  if ($TableStyle) { [void]$sb.AppendFormat(' style="{0}"', $TableStyle) }
  [void]$sb.Append("><thead")
  if ($TheadStyle) { [void]$sb.AppendFormat(' style="{0}"', $TheadStyle) }
  [void]$sb.Append("><tr")
  if ($TrStyle) { [void]$sb.AppendFormat(' style="{0}"', $TrStyle) }
  [void]$sb.Append(">")
  foreach ($col in $Results.Columns)
  {
    if ($ThStyle)
    {
      [void]$sb.AppendFormat('<th style="{0}">{1}</th>', $ThStyle, $col)
    }
    else
    {
      [void]$sb.AppendFormat("<th>{0}</th>", $col)
    }
  }
  [void]$sb.Append("</thead><tbody")
  if ($TbodyStyle) { [void]$sb.AppendFormat(' style="{0}"', $TbodyStyle) }
  [void]$sb.Append(">")
  $rows = $Results.RowCount
  $cols = $Results.ColumnCount
  for ($i = 0; $i -lt $rows; $i++)
  {
    if ($TrStyle)
    {
      [void]$sb.AppendFormat('<tr style="{0}">', $TrStyle)
    }
    else
    {
      [void]$sb.Append("<tr>")
    }
    for ($j = 0; $j -lt $cols; $j++)
    {
      if ($TdStyle)
      {
        [void]$sb.AppendFormat('<td style="{0}">{1}</td>', $TdStyle, $Results.Data[($i,$j)])
      }
      else
      {
        [void]$sb.AppendFormat("<td>{0}</td>", $Results.Data[($i,$j)])
      }
    }
    [void]$sb.Append("</tr>")
  }
  [void]$sb.Append("</tbody></table>")
  return $sb.ToString()
}