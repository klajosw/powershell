$nl = [Environment]::NewLine
$wrdc_mining = @{}
$words = "'Twas the night before Santa's party, when the elves' tools ' " ## direkt szöveg megadás
 
## $words = Get-Content "C:\files\kl\source.txt"    ## beolvasás
$words = $words.ToLower() -replace "[.,!?)(]",""    ## egséges kisbetûsítés és tisztítás


# $words = $words -replace "([^a-z])'([^a-z])","`${1}`${2}" ## szócsere
$words = $words -replace "'","''"
foreach ($iwc in $words.Split(" "))
{
 if (($wrdc_mining[$iwc]) -ne $null)
 {
  $cnt = $wrdc_mining[$iwc] + 1
  $wrdc_mining.Remove($iwc)
  $wrdc_mining.Add($iwc, $cnt)
 }
 else
 {
  $wrdc_mining.Add($iwc, 1)
 }
}

foreach ($k in $wrdc_mining.Keys)
{
 $sql_insert += "INSERT INTO $table (Word, WordCount) VALUES ('$k'," + $wrdc_mining[$k] + ")" + $nl
}

####----
Function TextMining_WordCounts($file, $table, $server, $db)
{

    $nl = [Environment]::NewLine
    $wrdc_mining = @{}
 
    $agg_snt = Get-Content $file
    $agg_snt = $agg_snt.ToLower() -replace "[.,!?)(]",""
    $agg_snt = $agg_snt -replace "([^a-z])'([^a-z])","`${1}`${2}"
    $agg_snt = $agg_snt -replace "'","''"

    foreach ($iwc in $agg_snt.Split(" "))
    {
     if (($wrdc_mining[$iwc]) -ne $null)
     {
      $cnt = $wrdc_mining[$iwc] + 1
      $wrdc_mining.Remove($iwc)
      $wrdc_mining.Add($iwc, $cnt)
     }
     else
     {
      $wrdc_mining.Add($iwc, 1)
     }
    }

    $sql_insert = "IF OBJECT_ID('$table') IS NULL BEGIN CREATE TABLE $table (Word VARCHAR(250), WordCount INT NULL, WordDate DATE DEFAULT GETDATE()) END" + $nl

    foreach ($k in $wrdc_mining.Keys)
    {
        $sql_insert += "INSERT INTO $table (Word, WordCount) VALUES ('$k'," + $wrdc_mining[$k] + ")" + $nl
    }

    $scon = New-Object System.Data.SqlClient.SqlConnection
    $scon.ConnectionString = "SERVER=" + $server + ";DATABASE=" + $db + ";Integrated Security=true"

    $adddata = New-Object System.Data.SqlClient.SqlCommand
    $adddata.Connection = $scon
    $adddata.CommandText = $sql_insert

    $scon.Open()
    $adddata.ExecuteNonQuery()
    $scon.Close()
    $scon.Dispose()
}

TextMining_WordCounts -file "C:\files\OurFile.txt" -table "OurTable" -server "OURSERVER\OURINSTANCE" -db "OURDATABASE"

### ----
Function MineFile_WordPosition ($file, $srv, $db)
{
    $nl = [Environment]::NewLine
    $sn = Get-Content $file
    $sn = $sn -replace "'","''"
    $sn = $sn -replace "[,;:]",""
 
    $cnt = 1
    $add = "IF OBJECT_ID('Word_Stage') IS NOT NULL BEGIN DROP TABLE Word_Stage END CREATE TABLE Word_Stage(Word VARCHAR(250), WordPosition INT, StatementDate DATETIME DEFAULT GETDATE())"
    foreach ($s in $sn.Split(" "))
    {
        if ($s -ne "")
        {
            $add += $nl + "INSERT INTO Word_Stage (Word,WordPosition) VALUES ('$s',$cnt)"
            $cnt++
        }
    }
 
    $scon = New-Object System.Data.SqlClient.SqlConnection
    $scon.ConnectionString = "SERVER=$srv;DATABASE=$db;integrated security=true"
 
    $cmd = New-Object System.Data.SqlClient.SqlCommand
    $cmd.CommandText = $add
    $cmd.Connection = $scon
 
    $scon.Open()
    $cmd.ExecuteNonQuery()
    $scon.Close()
    $scon.Dispose()
}

MineFile_WordPosition -file "C:\files\Other\fedstatement.txt" -srv "TIMOTHY\SQLEXPRESS" -db "MSSQLTips"

### ---------


