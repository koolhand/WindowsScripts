$htmlTemplateTop = @"
<!DOCTYPE html>
<html lang="en">
  <head>
    <meta charset="utf-8" />
    <title>{0}</title>
    <style>
    {1}
    </style>
  </head>
  <body>
  <h1>{0}</h1>
  {2}
"@

$htmlTemplateBottom = @"
</body>
</html>
"@

$htmlDescriptionTemplate = @"
<h2>{0}</h2>
<p>{1}</p>
"@

$htmlTitle = "The Title"
$htmlIntro = "<p>A bunch of useful reports from {0}</p>" -f (Get-Date -Format "dddd d MMM yyyy, h:mm:ss tt")
$htmlCss = (Get-Content -Path "./html/gutenberg-min.css" | Out-String)

($htmlDescriptionTemplate -f "Services","A list of Windows services.")      | Out-File './html/temp/1.html'
Get-Service | ConvertTo-Html  -Fragment -Property DisplayName, Name, Status | Out-File './html/temp/1.html' -Append

($htmlDescriptionTemplate -f "Modules","A list of PowerShell services.")    | Out-File './html/temp/2.html'
Get-Module  | ConvertTo-Html  -Fragment -Property Name, Version             | Out-File './html/temp/2.html' -Append

($htmlTemplateTop -f $htmlTitle, $htmlCss, $htmlIntro), (Get-Content './html/temp/*.htm*'), $htmlTemplateBottom | Set-Content './html/output/report.html'

Invoke-Item './html/output/report.html'