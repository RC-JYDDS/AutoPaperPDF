$ErrorActionPreference='Stop'
[void][Reflection.Assembly]::LoadFrom($env:AUTOPAPER_TEST_EXE)
$count=0
function Check([bool]$ok,[string]$label){if(-not $ok){throw ('FAIL '+$label)};$script:count++;Write-Output ('PASS '+$label)}
$day=[datetime]::new(2026,11,20)
Check ([AutoPaper.Core]::PdfName('{日期}','差旅报销','',$day) -eq '20261120差旅报销.pdf') 'date prefix'
Check ([AutoPaper.Core]::PdfName('','差旅报销','{日期}',$day) -eq '差旅报销20261120.pdf') 'date suffix'
Check ([AutoPaper.Core]::PdfName('财务','差旅报销','定稿',$day) -eq '财务差旅报销定稿.pdf') 'literal prefix and suffix'
Check ([AutoPaper.Core]::PdfName('{日期}_','差旅报销','_已核对',$day) -eq '20261120_差旅报销_已核对.pdf') 'explicit separators only'
Check ([AutoPaper.Core]::PdfName('','费用.PDF','',$day) -eq '费用.pdf') 'single extension'
Check ([AutoPaper.Core]::PdfName('{日期}','费用','',$day.AddDays(1)) -eq '20261121费用.pdf') 'next-day update'
foreach($bad in @('','../file','CON','a:b',('x'*252))){$rejected=$false;try{[AutoPaper.Core]::PdfName('',$bad,'',$day)}catch{$rejected=$true};Check $rejected ('reject invalid name '+$bad.Substring(0,[Math]::Min($bad.Length,12)))}
Check ([AutoPaper.Core]::PdfName('','CON费用','',$day) -eq 'CON费用.pdf') 'valid reserved-word prefix'
$culture=[Threading.Thread]::CurrentThread.CurrentCulture
try{[Threading.Thread]::CurrentThread.CurrentCulture=[Globalization.CultureInfo]::GetCultureInfo('th-TH');Check ([AutoPaper.Core]::PdfName('{日期}','费用','',$day) -eq '20261120费用.pdf') 'Gregorian date independent of display culture'}finally{[Threading.Thread]::CurrentThread.CurrentCulture=$culture}
$folder=Join-Path ([IO.Path]::GetTempPath()) ('native-app-test-'+[Guid]::NewGuid().ToString('N'))
[void][IO.Directory]::CreateDirectory($folder)
try{
 $path=Join-Path $folder 'settings.json'
 $s=[AutoPaper.Core]::ReadSettings($path)
 Check ($s.Remember -and $s.Prefix -eq '{日期}') 'first-run settings'
 $s.BaseName='差旅报销';$s.Suffix='定稿';[AutoPaper.Core]::SaveSettings($path,$s)
 $s=[AutoPaper.Core]::ReadSettings($path)
 Check ($s.BaseName -eq '差旅报销' -and $s.Suffix -eq '定稿') 'remember filename and affixes'
 $s.BaseName='其他费用';[AutoPaper.Core]::SaveSettings($path,$s)
 Check ([AutoPaper.Core]::ReadSettings($path).BaseName -eq '其他费用') 'replace settings atomically'
 $s.Remember=$false;[AutoPaper.Core]::SaveSettings($path,$s);$s=[AutoPaper.Core]::ReadSettings($path)
 Check (-not $s.Remember -and $s.BaseName -eq '' -and $s.Prefix -eq '' -and $s.Suffix -eq '') 'forget all naming settings'
 [IO.File]::WriteAllText($path,'{"Version":1,"Remember":true,"Prefix":"{日期}","Suffix":"旧设置"}',[Text.Encoding]::UTF8)
 $s=[AutoPaper.Core]::ReadSettings($path)
 Check ($s.Warning -eq '' -and $s.Suffix -eq '旧设置' -and $s.BaseName -eq '') 'migrate original script preferences'
 [IO.File]::WriteAllText($path,'{"Version":2,"Remember":true,"Prefix":"{日期}","BaseName":"报销单","Suffix":""}',[Text.Encoding]::UTF8)
 $s=[AutoPaper.Core]::ReadSettings($path)
 Check ($s.Warning -eq '' -and $s.BaseName -eq '报销单') 'read version 1.3 preferences'
 [IO.File]::WriteAllText($path,'broken');$s=[AutoPaper.Core]::ReadSettings($path)
 Check ($s.Warning.Length -gt 0 -and $s.Prefix -eq '{日期}') 'corrupt preferences handled'
 $temp=Join-Path $folder 'new.pdf';$dest=Join-Path $folder 'out.pdf'
 [IO.File]::WriteAllText($temp,'%PDF-1.7 MOCK HEADER '+('x'*200));[AutoPaper.Core]::CommitPdf($temp,$dest)
 Check ((Test-Path $dest) -and -not (Test-Path $temp)) 'commit completed file'
 [IO.File]::WriteAllText($temp,'INVALID '+('x'*200));$before=[IO.File]::ReadAllText($dest)
 $rejected=$false;try{[AutoPaper.Core]::CommitPdf($temp,$dest)}catch{$rejected=$true}
 Check ($rejected -and [IO.File]::ReadAllText($dest) -eq $before) 'bad output preserves previous PDF'
 [IO.File]::WriteAllText($temp,'%PDF-1.7 SECOND MOCK '+('y'*200));[AutoPaper.Core]::CommitPdf($temp,$dest)
 Check ([IO.File]::ReadAllText($dest).Contains('SECOND')) 'replace previous PDF only after validation'
 Check (@(Get-ChildItem $folder -Filter *.tmp).Count -eq 0) 'no leftover temporary preferences'
}finally{Remove-Item -LiteralPath $folder -Recurse -Force}
function Paper([string]$name,[int]$kind,[double]$w,[double]$h){$p=[AutoPaper.Paper]::new();$p.Name=$name;$p.Kind=$kind;$p.WidthPt=$w*72/25.4;$p.HeightPt=$h*72/25.4;return $p}
[AutoPaper.Paper[]]$papers=@((Paper A4 9 210 297),(Paper A5 11 148 210),(Paper DL 27 110 220))
$c=[AutoPaper.Core]::Candidates($papers,195*72/25.4,105*72/25.4)
Check ($c[0].Kind -eq 27 -and $c[0].Orientation -eq 2) 'small form selects DL landscape'
$c=[AutoPaper.Core]::Candidates($papers,195*72/25.4,130*72/25.4)
Check ($c[0].Kind -eq 11 -and $c[0].Orientation -eq 2) 'added rows select A5 landscape'
$c=[AutoPaper.Core]::Candidates($papers,195*72/25.4,180*72/25.4)
Check ($c[0].Kind -eq 9 -and $c[0].Orientation -eq 1) 'larger form selects A4 portrait'
$c=[AutoPaper.Core]::Candidates($papers,270*72/25.4,190*72/25.4)
Check ($c[0].Kind -eq 9 -and $c[0].Orientation -eq 2) 'wide form selects A4 landscape'
Check ([AutoPaper.Core]::Candidates($papers,2000,2000).Count -eq 0) 'oversize rejected without scaling'
Check ([AutoPaper.Core]::InsideBreaks([int[]]@(1,5,21),5,20) -eq 0) 'page-boundary breaks excluded'
Check ([AutoPaper.Core]::InsideBreaks([int[]]@(5,6,20,21),5,20) -eq 2) 'interior breaks included'
Check ([AutoPaper.Core]::RangeAddress(' $c$1:$ao$13 ') -eq '$C$1:$AO$13') 'arbitrary range'
$rejected=$false;try{[AutoPaper.Core]::RangeAddress('A1:C3,F2:G7')}catch{$rejected=$true};Check $rejected 'multiple print regions rejected'
Check ([AutoPaper.Core]::PrinterName('Printer Pro on Ne00:',[string[]]@('Printer','Printer Pro')) -eq 'Printer Pro') 'matching printer uses longest prefix'
$rejected=$false;try{[AutoPaper.Core]::PrinterName('Missing',[string[]]@('Printer'))}catch{$rejected=$true};Check $rejected 'unknown printer rejected'
Write-Output ('ALL PASS: '+$count+' compiled-core assertions; Windows GUI and Office COM not exercised.')
