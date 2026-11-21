$ErrorActionPreference='Stop'
[void][Reflection.Assembly]::LoadFrom($env:AUTOPAPER_TEST_EXE)
$count=0
function Check([bool]$ok,[string]$label){if(-not $ok){throw ('FAIL '+$label)};$script:count++;Write-Output ('PASS '+$label)}
Add-Type -TypeDefinition @'
using System;
using System.IO;
using System.Collections.Generic;
using System.Runtime.InteropServices;
public class InteropSheet {
 public bool BadFile,NoFile,FailExport,FailDelete;
 public int Index=1;
 public InteropSheet Previous;
 public bool ProtectContents,PasswordRequired;
 public int UnprotectCalls;
 public void Unprotect(string password) {
  UnprotectCalls++;
  if(password!="" || PasswordRequired)throw new COMException("protected",unchecked((int)0x8FE30C28));
  ProtectContents=false;
 }
 public int Calls,CopyCalls;public object Before;public string Name="Original";
 public InteropApp App;public InteropBook Book;
 public InteropPage PageSetup=new InteropPage();
 public InteropRanges Range=new InteropRanges();
 public void ExportAsFixedFormat(int type,string path,int quality,bool properties,bool ignore) {
  Calls++;
  if(type!=0 || quality!=0 || ignore)throw new Exception("Wrong print area/type/quality");
  if(FailExport)throw new COMException("simulated PDF error",unchecked((int)0x8FE30C28));
  if(NoFile)return;
  File.WriteAllText(path,(BadFile?"BAD!!":"%PDF-")+new String('x',200));
 }
 public void Copy(object before) {
  if(!Object.ReferenceEquals(before,this))throw new Exception("Copy target changed");
  CopyCalls++;Before=before;
  if(Book!=null) {
   var s=new InteropSheet {Name="Measure",App=App,Book=Book,FailExport=Book.ProbeFail,FailDelete=Book.DeleteFail};
   s.ProtectContents=ProtectContents;s.PasswordRequired=PasswordRequired;
   Previous=s;Index++;App.ActiveSheet=s;Book.Worksheets.Count++;Book.Copy=s;
  }
 }
 public void Select(bool replace) {if(App!=null)App.ActiveSheet=this;}
 public void Delete() {if(FailDelete)throw new COMException("simulated delete error",unchecked((int)0x80004005));if(Book!=null)Book.Worksheets.Count--;}
}
public class InteropApp {public InteropSheet ActiveSheet;}
public class InteropBook {public InteropSheets Worksheets=new InteropSheets();public InteropSheet Copy;public bool ProbeFail,DeleteFail;}
public class InteropSheets {public int Count=1;}
public class InteropRanges {public InteropRange Value=new InteropRange();public InteropRange this[string address]{get {if(address!="C1:AP13")throw new Exception("Wrong range");return Value;}}}
public class InteropRange {public double Width=562,Height=292;public InteropBorders Borders=new InteropBorders();}
public class InteropBorders {public InteropBorders Item {get{return this;}}public InteropEdge[] Edges={new InteropEdge(),new InteropEdge(),new InteropEdge(),new InteropEdge()};[System.Runtime.CompilerServices.IndexerName("Edge")] public InteropEdge this[int i]{get{return Edges[i-7];}}}
public class InteropEdge {public int LineStyle,Weight,Color;}
public static class MeasurementHarness {
 public static void Run(object session,InteropSheet original,string path,Action<string> logger) {
  var flags=System.Reflection.BindingFlags.Instance|System.Reflection.BindingFlags.NonPublic;
  var type=session.GetType();type.GetField("app",flags).SetValue(session,original.App);type.GetField("book",flags).SetValue(session,original.Book);
  type.GetMethod("ExportMeasurement",flags).Invoke(session,new object[]{original,original.Range.Value,"C1:AP13",original.PageSetup,path,new int[]{1,2,3,4},logger});
 }
}
public class InteropPage {
 public string PrintArea; public int Zoom,PaperSize=9,Orientation=2;
 public bool FitToPagesWide,FitToPagesTall,CenterHorizontally,CenterVertically,BlackAndWhite,Draft;
 public double TopMargin=14.17322835,BottomMargin=14.17322835,LeftMargin=14.17322835,RightMargin=14.17322835;
}
'@
$dir=Join-Path ([IO.Path]::GetTempPath()) ('office-interop-'+[Guid]::NewGuid().ToString('N'));[void][IO.Directory]::CreateDirectory($dir)
try {
 $path=Join-Path $dir 'export.pdf'
 $sheet=[InteropSheet]::new();[AutoPaper.OfficeInterop]::ExportPdf($sheet,$path)
 Check ($sheet.Calls -eq 1 -and (Test-Path $path)) 'PDF export works with exactly five explicit parameters'
 [AutoPaper.OfficeInterop]::CopyBefore($sheet)
 Check ($sheet.CopyCalls -eq 1 -and [Object]::ReferenceEquals($sheet,$sheet.Before)) 'worksheet copy uses one Before parameter without Missing.Value'
 foreach($case in @('NoFile','BadFile','FailExport')) {
  Remove-Item $path -ErrorAction SilentlyContinue;$sheet=[InteropSheet]::new();$sheet.$case=$true;$failed=$false
  try{[AutoPaper.OfficeInterop]::ExportPdf($sheet,$path)}catch{$failed=$true}
  Check $failed "reject export $case"
 }
 $logs=[Collections.Generic.List[string]]::new();$logger=[Action[string]]{$logs.Add($args[0])};$message=''
 try{[AutoPaper.OfficeInterop]::Run('导出原始 PDF',[Action]{throw [Runtime.InteropServices.COMException]::new('simulated',-1880945624)},$logger)}catch{$message=$_.Exception.ToString()}
 Check ($message.Contains('导出原始 PDF') -and $message.Contains('0x8FE30C28') -and $logs.Count -eq 1) 'original stage and HRESULT retained; no false completion log'
 function Paper($kind,$orientation,$width,$height){$p=[AutoPaper.Paper]::new();$p.Kind=$kind;$p.Orientation=$orientation;$p.WidthPt=$width;$p.HeightPt=$height;return $p}
 [AutoPaper.Paper[]]$papers=@((Paper 256 2 600 340),(Paper 11 2 595 420),(Paper 9 1 595 842),(Paper 9 2 842 595),(Paper 8 2 1191 842))
 $order=[AutoPaper.OfficeInterop]::ExportPapers($papers,560,290)
 Check ($order[0].Kind -eq 9 -and $order[0].Orientation -eq 2) 'landscape range prefers standard A4 landscape over smaller driver form'
 $order=[AutoPaper.OfficeInterop]::ExportPapers($papers,290,560)
 Check ($order[0].Kind -eq 9 -and $order[0].Orientation -eq 1) 'portrait range prefers standard A4 portrait'
 $order=[AutoPaper.OfficeInterop]::ExportPapers([AutoPaper.Paper[]]@($papers[0]),560,290)
 Check ($order.Count -eq 1 -and $order[0].Kind -eq 256) 'available printer forms retained when no common paper exists'
 # Exercise the real measurement orchestration with in-memory COM substitutes.
 # The actual Windows COM dispatch/driver remains outside this test.
 $type=[AutoPaper.OfficeSession];$flags=[Reflection.BindingFlags]'Instance,NonPublic'
 foreach($case in @('success','protected','password','probe-failure','probe-and-cleanup-failure')) {
  $session=[Runtime.Serialization.FormatterServices]::GetUninitializedObject($type)
  $app=[InteropApp]::new();$book=[InteropBook]::new();$original=[InteropSheet]::new();$original.App=$app;$original.Book=$book;$app.ActiveSheet=$original
  $original.ProtectContents=$case -in @('protected','password');$original.PasswordRequired=$case -eq 'password'
  $book.ProbeFail=$case -like 'probe-*';$book.DeleteFail=$case -eq 'probe-and-cleanup-failure'
  $type.GetField('app',$flags).SetValue($session,$app);$type.GetField('book',$flags).SetValue($session,$book)
  $logs.Clear();$message=''
  try{[MeasurementHarness]::Run($session,$original,$path,$logger)}catch{$message=$_.Exception.ToString()}
  Check ([Object]::ReferenceEquals($app.ActiveSheet,$original)) "restore original sheet: $case"
  Check ($original.Range.Value.Borders.Edges[0].Color -eq 0) "original borders untouched: $case"
  Check ($original.UnprotectCalls -eq 0) "original protection untouched: $case"
  if($case -in @('success','protected')) {
   if($message){Write-Output $message}
   Check ($message -eq '' -and $book.Worksheets.Count -eq 1 -and $book.Copy.Calls -eq 1) 'measurement export and cleanup complete'
   Check (-not $book.Copy.ProtectContents -and $book.Copy.UnprotectCalls -eq [int]$original.ProtectContents) "copy protection: $case"
   Check (($book.Copy.Range.Value.Borders.Edges.Color -join ',') -eq '1,2,4,3') 'COM bottom=9 and right=10 use the correct PDF marker colors'
  }elseif($case -eq 'password'){
   Check ($message.Contains('空密码') -and $book.Copy.Calls -eq 0 -and $book.Worksheets.Count -eq 1) 'password-protected copy rejected and deleted before PDF export'
  }else{
   Check ($message.Contains('导出测量 PDF') -and $message.Contains('0x8FE30C28')) "identify measurement export failure: $case"
   if($case -eq 'probe-failure'){Check ($book.Worksheets.Count -eq 1) 'remove temporary sheet after failed export'}
   else{Check (($logs -join "`n").Contains('删除测量工作表副本') -and -not $message.Contains('simulated delete error')) 'cleanup failure logged without masking original HRESULT'}
  }
 }
}finally{Remove-Item -LiteralPath $dir -Recurse -Force}
Write-Output "ALL PASS: $count Office invocation and error-path checks; COM substitutes, not live Office."
