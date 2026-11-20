$ErrorActionPreference='Stop'
[void][Reflection.Assembly]::LoadFrom($env:AUTOPAPER_TEST_EXE)
Add-Type -TypeDefinition @'
public class TestPage {
 public double TopMargin { get; set; }
 public double BottomMargin { get; set; }
 public double LeftMargin { get; set; }
 public double RightMargin { get; set; }
 public bool CenterHorizontally { get; set; }
 public bool CenterVertically { get; set; }
 public int Zoom { get; set; }
 public string LeftHeader { get; set; }
}
public class ClampingPage {
 double top;
 public double TopMargin { get { return top; } set { top=System.Math.Max(36,value); } }
 public double BottomMargin { get; set; }
 public double LeftMargin { get; set; }
 public double RightMargin { get; set; }
}
'@
$count=0
function Check([bool]$ok,[string]$label){if(-not $ok){throw ('FAIL '+$label)};$script:count++;Write-Output ('PASS '+$label)}
function Near([double]$a,[double]$b){return [Math]::Abs($a-$b) -lt 0.000001}
function Margins([double]$t,[double]$b,[double]$l,[double]$r){$m=[AutoPaper.PageMargins]::new();$m.UseCustom=$true;$m.TopMm=$t;$m.BottomMm=$b;$m.LeftMm=$l;$m.RightMm=$r;return $m}
Check (Near ([AutoPaper.PageMargins]::ToPoints(25.4)) 72) 'millimetres to points'
Check (Near ([AutoPaper.PageMargins]::ToMm(72)) 25.4) 'points to millimetres'
$page=[TestPage]::new();$old=Margins 3 3 4.5 4.5;$old.Apply($page);$page.CenterHorizontally=$true;$page.CenterVertically=$true;$page.Zoom=100;$page.LeftHeader='header'
$state=[AutoPaper.PageMarginState]::new($page)
$custom=Margins 2.3 7.8 12.5 0
$state.ApplyForExport($page,$custom)
$actual=[AutoPaper.PageMargins]::FromPage($page)
Check ((Near $actual.TopMm 2.3) -and (Near $actual.BottomMm 7.8) -and (Near $actual.LeftMm 12.5) -and (Near $actual.RightMm 0)) 'four asymmetric margins applied without swapping'
Check (-not $page.CenterHorizontally -and -not $page.CenterVertically) 'custom placement disables centering'
Check ($page.Zoom -eq 100 -and $page.LeftHeader -eq 'header') 'margin setting leaves scale and header untouched'
$state.ApplyForExport($page,[AutoPaper.PageMargins]::new())
$actual=[AutoPaper.PageMargins]::FromPage($page)
Check ((Near $actual.TopMm 3) -and (Near $actual.BottomMm 3) -and (Near $actual.LeftMm 4.5) -and (Near $actual.RightMm 4.5)) 'switch back restores original sheet margins'
Check ($page.CenterHorizontally -and -not $page.CenterVertically) 'restore horizontal centering and retain export top placement'
$custom.TopMm=9;$state.ApplyForExport($page,$custom);$state.ApplyForExport($page,$null)
Check (Near ([AutoPaper.PageMargins]::FromPage($page).TopMm) 3) 'repeat export does not overwrite original margin snapshot'
$copy=$state.Original;$copy.TopMm=66
Check (Near $state.Original.TopMm 3) 'original margins returned as independent copy'
$second=[TestPage]::new();(Margins 10 11 12 13).Apply($second);$secondState=[AutoPaper.PageMarginState]::new($second)
$secondState.ApplyForExport($second,$custom);$secondState.ApplyForExport($second,$null)
Check ((Near ([AutoPaper.PageMargins]::FromPage($second).LeftMm) 12) -and (Near $state.Original.LeftMm 4.5)) 'worksheets retain independent defaults'
(Margins 0 0 0 0).Validate();(Margins 100 100 100 100).Validate()
Check $true 'zero and maximum margins supported'
foreach($bad in @(-0.1,100.1,[double]::NaN,[double]::PositiveInfinity)){
 $m=Margins 5 5 5 $bad;$rejected=$false;try{$m.Validate()}catch{$rejected=$true};Check $rejected ('invalid margin rejected '+$bad)
}
$clamped=[ClampingPage]::new();$m=Margins 0 5 5 5;$m.Apply($clamped);$rejected=$false;try{$m.Verify($clamped)}catch{$rejected=$true}
Check $rejected 'driver silently clamping margins is detected'
$dir=Join-Path ([IO.Path]::GetTempPath()) ('margin-test-'+[Guid]::NewGuid().ToString('N'));[void][IO.Directory]::CreateDirectory($dir)
try{
 $path=Join-Path $dir 'page-margins.json';$s=[AutoPaper.PageMargins]::Load($path)
 Check (-not $s.UseCustom -and $s.TopMm -eq 5 -and $s.Warning -eq '') 'first run preserves spreadsheet margins'
 $s=Margins 2.3 7.8 12.5 0;[AutoPaper.PageMargins]::Save($path,$s);$s=[AutoPaper.PageMargins]::Load($path)
 Check ($s.UseCustom -and (Near $s.TopMm 2.3) -and (Near $s.BottomMm 7.8) -and (Near $s.LeftMm 12.5) -and $s.RightMm -eq 0) 'custom values and mode persist'
 $s.UseCustom=$false;[AutoPaper.PageMargins]::Save($path,$s);$s=[AutoPaper.PageMargins]::Load($path)
 Check (-not $s.UseCustom -and (Near $s.TopMm 2.3)) 'original mode persists without losing saved custom values'
 [IO.File]::WriteAllText($path,'{}');$s=[AutoPaper.PageMargins]::Load($path)
 Check (-not $s.UseCustom -and $s.TopMm -eq 5 -and $s.BottomMm -eq 5 -and $s.LeftMm -eq 5 -and $s.RightMm -eq 5) 'missing optional fields use safe defaults'
 [IO.File]::WriteAllText($path,'broken');$s=[AutoPaper.PageMargins]::Load($path)
 Check (-not $s.UseCustom -and $s.Warning.Length -gt 0) 'corrupt margin settings fall back with warning'
 [IO.File]::WriteAllText($path,'{"UseCustom":true,"TopMm":-1}');$s=[AutoPaper.PageMargins]::Load($path)
 Check (-not $s.UseCustom -and $s.Warning.Length -gt 0) 'invalid saved margins fall back with warning'
 $namePath=Join-Path $dir 'settings.json';$n=[AutoPaper.NamingSettings]::new();$n.BaseName='报销单';[AutoPaper.Core]::SaveSettings($namePath,$n)
 [AutoPaper.PageMargins]::Save($path,(Margins 8 9 10 11))
 Check ([AutoPaper.Core]::ReadSettings($namePath).BaseName -eq '报销单') 'old filename memory remains independent'
 Check (@(Get-ChildItem $dir -Filter '*.tmp').Count -eq 0) 'atomic settings save leaves no temporary files'
}finally{Remove-Item -LiteralPath $dir -Recurse -Force}
function Paper([string]$name,[int]$kind,[double]$w,[double]$h){$p=[AutoPaper.Paper]::new();$p.Name=$name;$p.Kind=$kind;$p.WidthPt=[AutoPaper.PageMargins]::ToPoints($w);$p.HeightPt=[AutoPaper.PageMargins]::ToPoints($h);return $p}
[AutoPaper.Paper[]]$papers=@((Paper Small 10 200 110),(Paper Large 11 220 140))
$m=Margins 5 5 5 5;$w=[AutoPaper.PageMargins]::ToPoints(180+$m.LeftMm+$m.RightMm)+1;$h=[AutoPaper.PageMargins]::ToPoints(80+$m.TopMm+$m.BottomMm)+1
Check ([AutoPaper.Core]::Candidates($papers,$w,$h)[0].Name -eq 'Small') 'paper selection includes small requested margins'
$m=Margins 5 5 15 15;$w=[AutoPaper.PageMargins]::ToPoints(180+$m.LeftMm+$m.RightMm)+1
Check ([AutoPaper.Core]::Candidates($papers,$w,$h)[0].Name -eq 'Large') 'larger margins select larger paper instead of scaling content'
$m=Margins 100 100 100 100;$w=[AutoPaper.PageMargins]::ToPoints(180+$m.LeftMm+$m.RightMm)+1;$h=[AutoPaper.PageMargins]::ToPoints(80+$m.TopMm+$m.BottomMm)+1
Check ([AutoPaper.Core]::Candidates($papers,$w,$h).Count -eq 0) 'excessive margins stop export rather than shrink'
Write-Output ('ALL PASS: '+$count+' margins checks; mocked page settings, not live Office COM.')
