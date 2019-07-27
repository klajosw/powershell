$directory1 = "\\K-file1\k-users\kokaipeter"
New-Item -Path $directory1 -Name "TestFoDir" -type Directory
$directory = "\\K-file1\k-users\kokaipeter\TestFoDir"
# alkönyvtárra öröklődés beállítása: szükséges flagek
$inherit = [system.security.accesscontrol.InheritanceFlags]"ContainerInherit, ObjectInherit"
$propagation = [system.security.accesscontrol.PropagationFlags]"None"
$acl = Get-Acl $directory
#jogok megadása
$accessrule = New-Object system.security.AccessControl.FileSystemAccessRule("olahzoltan@kfki.corp", "FullControl", $inherit, $propagation, "Allow")
$acl.AddAccessRule($accessrule)
#jogok könyvtárra alkalmazása
set-acl -aclobject $acl $directory

$directory1 = "\\K-file1\k-users\kokaipeter\TestFoDir"
New-Item -Path $directory1 -Name "TestAlDir" -type Directory
$directory = "\\K-file1\k-users\kokaipeter\TestFoDir\TestAlDir"
$inherit = [system.security.accesscontrol.InheritanceFlags]"None"
$propagation = [system.security.accesscontrol.PropagationFlags]"None"
$acl = Get-Acl $directory
#jogok megadása
$accessrule = New-Object system.security.AccessControl.FileSystemAccessRule("sneff gabor@kfki.corp", "Modify", $inherit, $propagation, "Allow")
$acl.AddAccessRule($accessrule)
#jogok könyvtárra alkalmazása
set-acl -aclobject $acl $directory