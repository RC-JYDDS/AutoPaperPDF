# Synthetic PDF fixtures contain no workbook data or real signatures.
param([string]$KeepFixtures = '')
$ErrorActionPreference = 'Stop'
[void][Reflection.Assembly]::LoadFrom($env:AUTOPAPER_TEST_EXE)
[AutoPaper.EmbeddedDependencies]::Register()
$count = 0
function Check([bool]$ok, [string]$label) { if (-not $ok) { throw "FAIL $label" }; $script:count++; Write-Output "PASS $label" }
function Number([double]$n) { $n.ToString('0.########', [Globalization.CultureInfo]::InvariantCulture) }
function New-TestPdf([string]$path, [int]$rotation = 0, [double]$unit = 1, [double]$originX = 0, [double]$originY = 0,
    [double]$contentW = 562.39370079, [double]$contentH = 291.96850394, [double]$left = 14.17322835, [double]$top = 14.17322835,
    [int]$pageCount = 1, [bool]$differentCrop = $false) {
    $paperW = 612; $paperH = 396
    $rot = (($rotation % 360) + 360) % 360
    $w = $paperW / $unit; $h = $paperH / $unit
    if ($rot -eq 90 -or $rot -eq 270) { $w = $paperH / $unit; $h = $paperW / $unit }
    $x1 = Number $originX; $y1 = Number $originY
    $x2 = Number ($originX + $w); $y2 = Number ($originY + $h)
    $scale = Number (1 / $unit); $neg = Number (-1 / $unit)
    $matrix = switch ($rot) {
        0   { "$scale 0 0 $scale $x1 $y1" }
        90  { "0 $scale $neg 0 $x2 $y1" }
        180 { "$neg 0 0 $neg $x2 $y2" }
        270 { "0 $neg $scale 0 $x1 $y2" }
        default { '1 0 0 1 0 0' }
    }
    $bottom = $paperH - $top - $contentH
    $content = "q $matrix cm 0.2 0.3 0.4 RG 0.5 w $(Number $left) $(Number $bottom) $(Number $contentW) $(Number $contentH) re S`n"
    $content += "BT /F1 12 Tf 1 0 0 1 $(Number ($left+8)) $(Number ($paperH-$top-20)) Tm (Amount 7996745.67 - ORIGINAL 100%) Tj ET`n"
    # A vector mark plus a tiny RGB image exercise both content types.
    $content += "1 0 0 RG $(Number ($left+20)) $(Number ($bottom+20)) m $(Number ($left+35)) $(Number ($bottom+27)) l $(Number ($left+50)) $(Number ($bottom+20)) l S`n"
    $content += "q 12 0 0 12 $(Number ($left+65)) $(Number ($bottom+18)) cm /Im1 Do Q Q`n"
    $crop = if ($differentCrop) { '/CropBox [5 5 600 380]' } else { "/CropBox [$x1 $y1 $x2 $y2]" }
    $kids = if ($pageCount -eq 2) { '3 0 R 8 0 R' } else { '3 0 R' }
    $objects = [Collections.Generic.List[string]]::new()
    $objects.Add('<< /Type /Catalog /Pages 2 0 R >>')
    $objects.Add("<< /Type /Pages /Count $pageCount /Kids [$kids] >>")
    $page = "<< /Type /Page /Parent 2 0 R /MediaBox [$x1 $y1 $x2 $y2] $crop /Rotate $rotation /UserUnit $(Number $unit) /TrimBox [$x1 $y1 $x2 $y2] /BleedBox [$x1 $y1 $x2 $y2] /ArtBox [$x1 $y1 $x2 $y2] /Resources << /Font << /F1 5 0 R >> /XObject << /Im1 7 0 R >> >> /Contents 4 0 R /Annots [6 0 R] >>"
    $objects.Add($page)
    $objects.Add("<< /Length $($content.Length) >>`nstream`n${content}endstream")
    $objects.Add('<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>')
    $objects.Add('<< /Type /Annot /Subtype /Link /Rect [30 350 200 370] /Border [0 0 0] /A << /S /URI /URI (https://example.org/) >> >>')
    $objects.Add("<< /Type /XObject /Subtype /Image /Width 3 /Height 1 /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /ASCIIHexDecode /Length 19 >>`nstream`nff000000ff000000ff>`nendstream")
    if ($pageCount -eq 2) { $objects.Add($page) }
    $buffer = [Text.StringBuilder]::new("%PDF-1.7`n")
    $offsets = [Collections.Generic.List[int]]::new()
    for ($i = 0; $i -lt $objects.Count; $i++) {
        $offsets.Add($buffer.Length)
        [void]$buffer.Append("$($i+1) 0 obj`n$($objects[$i])`nendobj`n")
    }
    $xref = $buffer.Length
    [void]$buffer.Append("xref`n0 $($objects.Count+1)`n0000000000 65535 f `n")
    foreach ($offset in $offsets) { [void]$buffer.Append($offset.ToString('0000000000') + " 00000 n `n") }
    [void]$buffer.Append("trailer`n<< /Size $($objects.Count+1) /Root 1 0 R >>`nstartxref`n$xref`n%%EOF`n")
    [IO.File]::WriteAllText($path, $buffer.ToString(), [Text.Encoding]::ASCII)
}
function Read-Pdf([string]$path) { [PdfSharp.Pdf.IO.PdfReader]::Open($path, [PdfSharp.Pdf.IO.PdfDocumentOpenMode]::Modify) }
function Content-Bytes($page) { [Convert]::ToBase64String($page.Contents.Elements[0].Value.Stream.UnfilteredValue) }
$dir = if ($KeepFixtures) { $KeepFixtures } else { Join-Path ([IO.Path]::GetTempPath()) ('page-fit-test-'+[Guid]::NewGuid().ToString('N')) }
[void][IO.Directory]::CreateDirectory($dir)
try {
    $expectedW = [AutoPaper.PageMargins]::ToPoints(208.4)
    $expectedH = [AutoPaper.PageMargins]::ToPoints(113)
    foreach ($rotation in @(0,90,180,270,-90,450)) {
        foreach ($unit in @(1,2)) {
            $input = Join-Path $dir "before-r$rotation-u$unit.pdf"
            $output = Join-Path $dir "after-r$rotation-u$unit.pdf"
            New-TestPdf $input $rotation $unit 23 37
            [AutoPaper.PdfPageFitter]::Fit($input,$output,$expectedW,$expectedH)
            $before = Read-Pdf $input; $after = Read-Pdf $output
            try {
                $p = $after.Pages[0]; $box = $p.MediaBox; $rot = (($rotation % 360)+360)%360
                $w = $box.Width * $unit; $h = $box.Height * $unit
                if ($rot -eq 90 -or $rot -eq 270) { $w = $box.Height * $unit; $h = $box.Width * $unit }
                Check (($after.PageCount -eq 1) -and ([Math]::Abs($w-$expectedW) -lt 0.01) -and ([Math]::Abs($h-$expectedH) -lt 0.01)) "physical size r$rotation u$unit"
                Check (($p.CropBox -eq $box) -and ($p.TrimBox -eq $box) -and ($p.BleedBox -eq $box) -and ($p.ArtBox -eq $box)) "all page boxes agree r$rotation u$unit"
                Check ((Content-Bytes $before.Pages[0]) -eq (Content-Bytes $p)) "text vector image commands unchanged r$rotation u$unit"
                Check ($p.Annotations.Count -eq 1) "link annotation retained r$rotation u$unit"
            } finally { $before.Dispose(); $after.Dispose() }
        }
    }
    foreach ($case in @(@(0,0,0,0),@(2.3,7.8,12.5,0),@(5,5,5,5))) {
        $t,$b,$l,$r = $case
        $input=Join-Path $dir "margins-$t-$b-$l-$r-before.pdf"; $output=Join-Path $dir "margins-$t-$b-$l-$r-after.pdf"
        New-TestPdf $input 0 1 0 0 400 200 ([AutoPaper.PageMargins]::ToPoints($l)) ([AutoPaper.PageMargins]::ToPoints($t))
        $w=400+[AutoPaper.PageMargins]::ToPoints($l+$r); $h=200+[AutoPaper.PageMargins]::ToPoints($t+$b)
        [AutoPaper.PdfPageFitter]::Fit($input,$output,$w,$h)
        $doc=Read-Pdf $output
        try { Check (([Math]::Abs($doc.Pages[0].MediaBox.Width-$w) -lt 0.01) -and ([Math]::Abs($doc.Pages[0].MediaBox.Height-$h) -lt 0.01)) "zero/asymmetric margins $t $b $l $r" } finally { $doc.Dispose() }
    }
    $input=Join-Path $dir 'failure-input.pdf'; $output=Join-Path $dir 'failure-output.pdf'
    foreach ($case in @('multiple-pages','too-small','invalid-size','pre-cropped','invalid-rotation','same-path')) {
        New-TestPdf $input
        switch ($case) {
            'multiple-pages' { New-TestPdf $input -pageCount 2 }
            'pre-cropped' { New-TestPdf $input -differentCrop $true }
            'invalid-rotation' { New-TestPdf $input -rotation 45 }
        }
        $before=[IO.File]::ReadAllBytes($input)
        [IO.File]::WriteAllText($output,'existing destination')
        $rejected=$false
        try {
            switch ($case) {
                'too-small' { [AutoPaper.PdfPageFitter]::Fit($input,$output,700,500) }
                'invalid-size' { [AutoPaper.PdfPageFitter]::Fit($input,$output,[double]::NaN,300) }
                'same-path' { [AutoPaper.PdfPageFitter]::Fit($input,$input,500,300) }
                default { [AutoPaper.PdfPageFitter]::Fit($input,$output,500,300) }
            }
        } catch { $rejected=$true }
        Check ($rejected -and ([IO.File]::ReadAllText($output) -eq 'existing destination') -and ([Convert]::ToBase64String($before) -eq [Convert]::ToBase64String([IO.File]::ReadAllBytes($input)))) "reject $case without replacing files"
    }
} finally { if (-not $KeepFixtures) { Remove-Item -LiteralPath $dir -Recurse -Force } }
Add-Type -TypeDefinition @'
public class TestPrintLayout {
 public string LeftHeader {get;set;} public string CenterHeader {get;set;} public string RightHeader {get;set;}
 public string LeftFooter {get;set;} public string CenterFooter {get;set;} public string RightFooter {get;set;}
 public string PrintTitleRows {get;set;} public string PrintTitleColumns {get;set;}
 public bool PrintHeadings {get;set;} public bool DifferentFirstPageHeaderFooter {get;set;} public bool OddAndEvenPagesHeaderFooter {get;set;}
}
'@
[AutoPaper.PrintAreaLayout]::Validate([TestPrintLayout]::new(),1,13,3,42); Check $true 'plain print range accepted'
foreach ($field in @('LeftHeader','CenterFooter','PrintTitleRows','PrintTitleColumns','PrintHeadings','DifferentFirstPageHeaderFooter','OddAndEvenPagesHeaderFooter')) {
    $p=[TestPrintLayout]::new()
    if ($field -in @('PrintHeadings','DifferentFirstPageHeaderFooter','OddAndEvenPagesHeaderFooter')) { $p.$field=$true } else { $p.$field='additional content' }
    $rejected=$false; try {[AutoPaper.PrintAreaLayout]::Validate($p,1,13,3,42)} catch {$rejected=$true}
    Check $rejected "do not silently crop $field"
}

# Regression: the reimbursement template prints C1:AP13 and repeats rows 1:4.
# Those rows already belong to the one-page content; they must not block export.
foreach ($title in @('$1:$4','1:4',"'费用报销单'!`$1:`$4",'$1:$13','  $1:$4  ')) {
    $p=[TestPrintLayout]::new(); $p.PrintTitleRows=$title
    [AutoPaper.PrintAreaLayout]::Validate($p,1,13,3,42)
    Check ($p.PrintTitleRows -eq $title) "contained title rows accepted without changing settings: $title"
}
foreach ($title in @('$C:$E','$C:$AP','c:ap',"'Sheet 1'!`$C:`$AP")) {
    $p=[TestPrintLayout]::new(); $p.PrintTitleColumns=$title; $p.PrintTitleRows='$1:$4'
    [AutoPaper.PrintAreaLayout]::Validate($p,1,13,3,42)
    Check ($p.PrintTitleColumns -eq $title -and $p.PrintTitleRows -eq '$1:$4') "contained title columns and rows accepted: $title"
}
foreach ($case in @(@('PrintTitleRows','$1:$14'),@('PrintTitleRows','$14:$15'),@('PrintTitleRows','$4:$1'),@('PrintTitleRows','$0:$4'),@('PrintTitleRows','$1:$99999999999999999999'),@('PrintTitleRows','$1:$2,$4:$4'),@('PrintTitleColumns','$A:$B'),@('PrintTitleColumns','$B:$E'),@('PrintTitleColumns','$C:$AQ'),@('PrintTitleColumns','$AP:$C'))) {
    $p=[TestPrintLayout]::new(); $p.($case[0])=$case[1]; $message=''
    try {[AutoPaper.PrintAreaLayout]::Validate($p,1,13,3,42)} catch {$message=$_.Exception.ToString()}
    Check ($message.Contains($case[1])) "external or invalid title reports its exact range: $($case[1])"
}
$p=[TestPrintLayout]::new();$p.PrintTitleRows='$1:$4';$message=''
try {[AutoPaper.PrintAreaLayout]::Validate($p,3,20,3,42)}catch{$message=$_.Exception.ToString()}
Check ($message.Contains('$1:$4')) 'partly external title rows are not treated as contained'
$p=[TestPrintLayout]::new();$p.PrintTitleRows='$1048576:$1048576';$p.PrintTitleColumns='$XFD:$XFD'
[AutoPaper.PrintAreaLayout]::Validate($p,1048576,1048576,16384,16384)
Check $true 'Excel last row and last column boundaries supported'

Write-Output "ALL PASS: $count PDF page and layout checks; no live Office COM."
