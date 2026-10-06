Add-Type -Path "itextsharp.dll"

$filePath = "example.pdf"

$pdf = New-Object iTextSharp.text.pdf.PdfReader -ArgumentList $filePath

$text = ""

foreach($page in 1..$($pdf.NumberOfPages)){

    $text = $text + [iTextSharp.text.pdf.parser.PdfTextExtractor]::GetTextFromPage($pdf, $page)

}

#Write-Host $text

$beginn = "Quelle Bewertung Rotten Tomatoes (Tomatometer) 89 %[9]"
$end = "Prädikat der FBW besonders wertvoll[11]"

Write-Host ($text -split "\$beginn\")[0]

#$result = (($text.Split($beginn))[1]).Split($end)[0]

#Write-Host $result

$pdf.close