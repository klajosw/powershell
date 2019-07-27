Rem KL PS meghívó cmd
REM --------------------------------------
rem   Set-ExecutionPolicy RemoteSigned  	REM --- PS engedélyezése
rem   C:\Scripts\Test.ps1
rem   .\Test.ps1
rem $a = $env:path; $a.Split(";")           REM --- PATH ellenörzése
rem  & "d:\prg\Reporting\Reports\kl\kl.ps1 d:\prg\Reporting\Reports\kl\kl.ps1"
rem vbs -ps
rem -- Set objShell = CreateObject("Wscript.Shell")
rem -- objShell.Run("powershell.exe -noexit c:\scripts\test.ps1")
rem powershell.exe -noexit &'d:\prg\Reporting\Reports\kl\kl.ps1 d:\prg\Reporting\Reports\kl\kl.ps1'  rem különleges karakternél
REM --------------------------------------

powershell.exe d:\prg\kl_ps\forgalom\futtato.ps1 




pause

Exit