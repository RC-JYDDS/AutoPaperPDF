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
    [int]$pageCount = 1, [bool]$differentCrop = $false, [bool]$probe=$false, [string]$mode="stroke", [int]$missing=-1, [double]$paperW=792, [double]$paperH=612) {
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
    $content += "q $matrix cm BT /F1 10 Tf 1 0 0 1 $(Number (($left+$contentW)-122)) $(Number ($bottom+8)) Tm (RIGHT-END 7996745.67) Tj ET Q`n"
    $marks=''
    if($probe) {
        $edges=@(@($left,$bottom,$left,($bottom+$contentH)),@($left,($bottom+$contentH),($left+$contentW),($bottom+$contentH)),@(($left+$contentW),$bottom,($left+$contentW),($bottom+$contentH)),@($left,$bottom,($left+$contentW),$bottom))
        for($i=0;$i -lt 4;$i++) {
            if($i -eq $missing){continue}
            $c=$colors[$i];$rgb="$(Number (($c -band 255)/255.0)) $(Number ((($c -shr 8) -band 255)/255.0)) $(Number ((($c -shr 16) -band 255)/255.0))"
            $x0,$y0,$x3,$y3=$edges[$i]
            if($mode -eq 'fill') {
                $rx=$x0;$ry=$y0;$rw=$x3-$x0;$rh=$y3-$y0
                if($rw -eq 0){$rx-=0.5;$rw=1}else{$ry-=0.5;$rh=1}
                $marks+="$rgb rg $(Number $rx) $(Number $ry) $(Number $rw) $(Number $rh) re f`n"
            }else{
                $marks+="$rgb RG 0.5 w $(Number $x0) $(Number $y0) m $(Number $x3) $(Number $y3) l S`n"
            }
        }
        if($mode -eq 'form'){$content+="q $matrix cm /Probe Do Q`n"}else{$content+="q $matrix cm`n$marks Q`n"}
    }
    $crop = if ($differentCrop) { '/CropBox [5 5 600 380]' } else { "/CropBox [$x1 $y1 $x2 $y2]" }
    $kids = if ($pageCount -eq 2) { '3 0 R 8 0 R' } else { '3 0 R' }
    $objects = [Collections.Generic.List[string]]::new()
    $objects.Add('<< /Type /Catalog /Pages 2 0 R >>')
    $objects.Add("<< /Type /Pages /Count $pageCount /Kids [$kids] >>")
    $page = "<< /Type /Page /Parent 2 0 R /MediaBox [$x1 $y1 $x2 $y2] $crop /Rotate $rotation /UserUnit $(Number $unit) /TrimBox [$x1 $y1 $x2 $y2] /BleedBox [$x1 $y1 $x2 $y2] /ArtBox [$x1 $y1 $x2 $y2] /Resources << /Font << /F1 5 0 R >> /XObject << /Im1 7 0 R /Probe 9 0 R >> >> /Contents 4 0 R /Annots [6 0 R] >>"
    $objects.Add($page)
    $objects.Add("<< /Length $($content.Length) >>`nstream`n${content}endstream")
    $objects.Add('<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>')
    $objects.Add('<< /Type /Annot /Subtype /Link /Rect [30 350 200 370] /Border [0 0 0] /A << /S /URI /URI (https://example.org/) >> >>')
    $objects.Add("<< /Type /XObject /Subtype /Image /Width 3 /Height 1 /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /ASCIIHexDecode /Length 19 >>`nstream`nff000000ff000000ff>`nendstream")
    if ($pageCount -eq 2) { $objects.Add($page) } else { $objects.Add('<< >>') }
    $objects.Add("<< /Type /XObject /Subtype /Form /BBox [0 0 $paperW $paperH] /Matrix [1 0 0 1 0 0] /Resources << >> /Length $($marks.Length) >>`nstream`n${marks}endstream")
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
$colors=@(0x5311EF,0x61BF17,0xE34F2F,0xD325C1)
function Read-Pdf([string]$path) { [PdfSharp.Pdf.IO.PdfReader]::Open($path, [PdfSharp.Pdf.IO.PdfDocumentOpenMode]::Modify) }
function Bytes($page) { [Convert]::ToBase64String($page.Contents.Elements[0].Value.Stream.UnfilteredValue) }
function Margins($top=5,$bottom=5,$left=5,$right=5) {
    $m=[AutoPaper.PageMargins]::new();$m.UseCustom=$true;$m.TopMm=$top;$m.BottomMm=$bottom;$m.LeftMm=$left;$m.RightMm=$right;return $m
}
$dir=if($KeepFixtures){$KeepFixtures}else{Join-Path ([IO.Path]::GetTempPath()) ('page-fit-test-'+[Guid]::NewGuid().ToString('N'))}
[void][IO.Directory]::CreateDirectory($dir)
$renderCases=@()
try {
    foreach($mode in @('stroke','fill','form')) {
        foreach($rotation in @(0,90,180,270,-90,450)) {
            foreach($unit in @(1,2)) {
                $name="$mode-r$rotation-u$unit";$input=Join-Path $dir "$name-before.pdf";$probe=Join-Path $dir "$name-probe.pdf";$output=Join-Path $dir "$name-after.pdf"
                New-TestPdf $input $rotation $unit 23 37
                New-TestPdf $probe $rotation $unit 23 37 -probe $true -mode $mode
                $size=[AutoPaper.PdfPageFitter]::FitMeasured($input,$probe,$output,(Margins),$colors)
                $before=Read-Pdf $input;$after=Read-Pdf $output
                try {
                    $p=$after.Pages[0];$box=$p.MediaBox;$rot=(($rotation%360)+360)%360
                    $w=$box.Width*$unit;$h=$box.Height*$unit
                    if($rot -eq 90 -or $rot -eq 270){$w=$box.Height*$unit;$h=$box.Width*$unit}
                    Check ([Math]::Abs($w-[AutoPaper.PageMargins]::ToPoints(208.4)) -lt 0.02 -and [Math]::Abs($h-[AutoPaper.PageMargins]::ToPoints(113)) -lt 0.02) "measured physical size $name"
                    Check (($p.CropBox -eq $box) -and ($p.TrimBox -eq $box) -and ($p.BleedBox -eq $box) -and ($p.ArtBox -eq $box)) "page boxes $name"
                    Check ((Bytes $before.Pages[0]) -eq (Bytes $p)) "original text image vector bytes $name"
                    Check ($p.Annotations.Count -eq 1) "link retained $name"
                }finally{$before.Dispose();$after.Dispose()}
            }
        }
    }
    # Width/height and position deliberately disagree with the old COM estimate.
    # The original failure would hide the right label and image on these pages.
    foreach($case in @(@('wider',620,292,14.17322835,14.17322835,5,5,5,5),@('taller',562,390,14.17322835,14.17322835,5,5,5,5),@('shifted',620,330,40,28,5,5,5,5),@('asymmetric',620,330,40,28,2.3,7.8,12.5,1),@('zero',620,330,20,20,0,0,0,0),@('expanded-paper',620,330,10,10,20,20,20,20))) {
        $name,$cw,$ch,$l,$t,$mt,$mb,$ml,$mr=$case
        $input=Join-Path $dir "$name-before.pdf";$probe=Join-Path $dir "$name-probe.pdf";$output=Join-Path $dir "$name-after.pdf"
        New-TestPdf $input -contentW $cw -contentH $ch -left $l -top $t
        New-TestPdf $probe -contentW $cw -contentH $ch -left $l -top $t -probe $true
        $size=[AutoPaper.PdfPageFitter]::FitMeasured($input,$probe,$output,(Margins $mt $mb $ml $mr),$colors)
        # The fixture strokes its outer rectangle at 0.5 pt, centered on each boundary.
        $expectedW=$cw+[Math]::Max(0.25,[AutoPaper.PageMargins]::ToPoints($ml))+[Math]::Max(0.25,[AutoPaper.PageMargins]::ToPoints($mr))
        $expectedH=$ch+[Math]::Max(0.25,[AutoPaper.PageMargins]::ToPoints($mt))+[Math]::Max(0.25,[AutoPaper.PageMargins]::ToPoints($mb))
        Check ([Math]::Abs($size[0]-$expectedW) -lt 0.02 -and [Math]::Abs($size[1]-$expectedH) -lt 0.02) "actual range, including blank cells and full outer strokes: $name"
        $renderCases+=@{name=$name;input=$input;output=$output;probe=$probe;contentW=$cw;contentH=$ch;left=$l;top=$t;margins=@($mt,$mb,$ml,$mr)}
    }
    $input=Join-Path $dir 'failure-input.pdf';$probe=Join-Path $dir 'failure-probe.pdf';$output=Join-Path $dir 'failure-output.pdf'
    foreach($case in @('multiple-pages','probe-multiple-pages','missing-edge','outside-paper','pre-cropped','probe-pre-cropped','different-geometry','shifted-content','invalid-rotation','invalid-margin','same-path','marker-collision')) {
        New-TestPdf $input;New-TestPdf $probe -probe $true
        switch($case){
            'multiple-pages'{New-TestPdf $input -pageCount 2}
            'probe-multiple-pages'{New-TestPdf $probe -probe $true -pageCount 2}
            'missing-edge'{New-TestPdf $probe -probe $true -missing 2}
            'outside-paper'{New-TestPdf $probe -probe $true -contentW 900}
            'pre-cropped'{New-TestPdf $input -differentCrop $true}
            'probe-pre-cropped'{New-TestPdf $probe -probe $true -differentCrop $true}
            'different-geometry'{New-TestPdf $probe -probe $true -rotation 90}
            'shifted-content'{New-TestPdf $probe -probe $true -left 40}
            'invalid-rotation'{New-TestPdf $input -rotation 45}
            'marker-collision'{New-TestPdf $input -probe $true}
        }
        $before=[Convert]::ToBase64String([IO.File]::ReadAllBytes($input));[IO.File]::WriteAllText($output,'existing destination');$rejected=$false
        $m=Margins;if($case -eq 'invalid-margin'){$m.LeftMm=[double]::NaN}
        try{if($case -eq 'same-path'){[void][AutoPaper.PdfPageFitter]::FitMeasured($input,$probe,$input,$m,$colors)}else{[void][AutoPaper.PdfPageFitter]::FitMeasured($input,$probe,$output,$m,$colors)}}catch{$rejected=$true}
        Check ($rejected -and ([IO.File]::ReadAllText($output) -eq 'existing destination') -and $before -eq [Convert]::ToBase64String([IO.File]::ReadAllBytes($input))) "reject $case without overwriting either file"
    }
    if($KeepFixtures){ConvertTo-Json -InputObject $renderCases -Depth 5|Set-Content -LiteralPath (Join-Path $dir 'render-cases.json') -Encoding UTF8}
}finally{if(-not $KeepFixtures){Remove-Item -LiteralPath $dir -Recurse -Force}}
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
