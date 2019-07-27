$WebResponse= Invoke-WebRequest https://mywebsite.com/page
ForEach ($Image in $WebResponse.Images)
{
    $FileName = Split-Path $Image.src -Leaf
    Invoke-WebRequest $Image.src -OutFile $FileName
	
	1

}

$WebResponse.AllElements | Where {$_.TagName -eq "a"}

$WebResponse.ParsedHtml.IHTMLDocument2_lastModified

$a = @(1,2,3,4,5,5,6,7,8,9,0,0)
$a = $a | select -uniq