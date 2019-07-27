$olTaskItem = 1
$olFolderTasks = 9



$outlook = New-Object -Com Outlook.Application
$task = $outlook.Application.CreateItem($olTaskItem) 
$hasError = $false 


$mapi = $outlook.GetNamespace("MAPI") 
$items = $mapi.GetDefaultFolder($olFolderTasks).Items 

$today = Get-Date
$startdate = $today.addmonths(-3)
$today = $today.ToLongDateString()
$startdate = $startdate.ToLongDateString()
$items | ForEach-Object `
{
 if ($today -eq ($_.CreationTime).ToLongDateString()) {
 	$_.subject
	If ($_.subject -eq "Beérkezett." -and $_.subject -eq "Távozott.") {
		#$_.delete()
		$_
	}
  }
}