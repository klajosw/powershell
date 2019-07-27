Get-Content Excelsupport.csv | Select -First 1 | Set-Content "Export.csv"
Import-Csv -Delimiter ";" -Path .\Excelsupport.csv | % { `
  $Row = $_; `
  $_."Gyártási számok".Split(";") | % { `
    "{0};{1};{2};{3};{4};{5}" -f `
    $Row."Cikkszám", `
    $Row."Megnevezés", `
    $Row."Gyártó", `
    $Row."Mennyiség", `
    $Row."Leltári szám", `
    $_ `
  } `
} | Add-Content Export.csv