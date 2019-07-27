function Import-Excel($nam = "d:\magan\test.xls", $sheet = 1){
 
 $xl = new-object -com Excel.Application
 # The locale could cause problems, so let's set it to en-US
 # details: http://support.microsoft.com/kb/320369
 $ci = [System.Globalization.CultureInfo]'en-us'
 $wb = $xl.workbooks.psbase.gettype().InvokeMember("Open",[Reflection.BindingFlags]::InvokeMethod, $null, $xl.workbooks, ($nam,$false,$true), $ci)
 
 # worksheet
 $sh = $wb.sheets.item($sheet)
 # column headers
 $head = $sh.range($sh.range("A1"), $sh.range("A1").end(-4161))
 set-variable $ImportExcelCount -scope global -value ($sh.range("A1").end(-4121).row - 1)
 
 # read each row as a associative array
 #$sh.range($sh.range("A2"), $sh.range("A1").end(-4121)).rows | foreach{ $row=$_; $out=@{}; $head | foreach{ $out[$_.formulalocal]=$row.range($_.addresslocal()).formulalocal}; $out }
 $sh.range($sh.range("A2"), $sh.range("A4")).rows | foreach{ $row=$_; $out=@{}; $head | foreach{ $out[$_.formulalocal]=$row.range($_.addresslocal()).formulalocal}; $out }
 $null = $wb.psbase.gettype().InvokeMember("Close",[Reflection.BindingFlags]::InvokeMethod, $null, $wb, $false, $ci)
 $null = $xl.quit()
 #unload from memory
 # see: http://www.microsoft.com/technet/scriptcenter/resources/pstips/nov07/pstip1130.mspx
 $null = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl)
} 
Import-Excel