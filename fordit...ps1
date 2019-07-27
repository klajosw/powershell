rem cscript ldap_lista.vbs "kecs*" kl.txt
rem powershell.exe .\Create-EXEFrom.ps1 .\kl1.ps1 .\kl1.exe

powershell.exe ./ps2exe.ps1 ./kl1.ps1 ./kl1.exe

pause