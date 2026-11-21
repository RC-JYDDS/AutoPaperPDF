using System;
using System.Collections.Generic;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;

namespace AutoPaper
{
    public sealed class OfficeSession : IDisposable
    {
        dynamic app, book;
        bool owned;
        readonly Dictionary<string, PageMarginState> sourceMargins = new Dictionary<string, PageMarginState>();
        public string FilePath { get; private set; }
        public string EngineName { get; private set; }
        public string EngineId { get; private set; }
        public string EngineVersion { get; private set; }

        static void Release(object value)
        { if (value != null && Marshal.IsComObject(value)) Marshal.FinalReleaseComObject(value); }

        public OfficeSession(string path, string engine)
        {
            if (!File.Exists(path)) throw new FileNotFoundException("所选文件不存在。", path);
            FilePath = path;
            string[] ids = engine == "WPS" ? new[] { "ket.Application" } :
                engine == "Excel" ? new[] { "Excel.Application" } : new[] { "ket.Application", "Excel.Application" };
            var failures = new List<string>();
            foreach (string id in ids)
            {
                try
                {
                    Type type = Type.GetTypeFromProgID(id);
                    if (type == null) throw new InvalidOperationException("未注册自动化接口");
                    app = Activator.CreateInstance(type); EngineId=id; break;
                }
                catch (Exception ex) { failures.Add(id + ": " + ex.Message); }
            }
            if (app == null) throw new InvalidOperationException("无法连接本机 Windows 桌面版 Excel/WPS。\r\n" + String.Join("\r\n", failures));
            try
            {
                if ((int)app.Workbooks.Count != 0)
                    throw new InvalidOperationException("该软件把请求交给了已有工作簿窗口。请先保存并关闭该软件的工作簿后重试，或选择另一软件。");
                owned = true;
                app.Visible = false; app.DisplayAlerts = false;
                try { app.EnableEvents = false; } catch { }
                try { app.AskToUpdateLinks = false; } catch { }
                bool securitySet = false;
                try { app.AutomationSecurity = 3; securitySet = (int)app.AutomationSecurity == 3; } catch { }
                if (!securitySet && !String.Equals(Path.GetExtension(path), ".xlsx", StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException("该软件未提供可验证的禁用文件宏接口；本次仅允许打开 .xlsx。");
                book = app.Workbooks.Open(path, 0, true);
                if (!(bool)book.ReadOnly) throw new InvalidOperationException("软件未以只读方式打开文件，已停止。");
                EngineName = (string)app.Name;
                try { EngineVersion=Convert.ToString(app.Version); } catch { EngineVersion="版本接口未提供"; }
            }
            catch { Dispose(); throw; }
        }

        public List<string> SheetNames()
        {
            var result = new List<string>(); dynamic sheets = book.Worksheets;
            try
            {
                for (int i = 1; i <= (int)sheets.Count; i++)
                {
                    dynamic sheet = sheets.Item[i];
                    try { if ((int)sheet.Visible == -1) result.Add((string)sheet.Name); }
                    finally { Release((object)sheet); }
                }
            }
            finally { Release((object)sheets); }
            return result;
        }

        public string ActiveSheetName()
        {
            dynamic sheet = book.ActiveSheet;
            try { return (string)sheet.Name; } finally { Release((object)sheet); }
        }

        public string PrintArea(string name)
        {
            dynamic sheet = book.Worksheets.Item[name]; dynamic page = null;
            try { page = sheet.PageSetup; return (string)page.PrintArea; }
            finally { Release((object)page); Release((object)sheet); }
        }

        void CaptureMargins(string name, object value)
        {
            if (sourceMargins.ContainsKey(name)) return;
            sourceMargins.Add(name, new PageMarginState(value));
        }

        public PageMargins OriginalMargins(string name)
        {
            dynamic sheet = book.Worksheets.Item[name]; dynamic page = null;
            try { page = sheet.PageSetup; CaptureMargins(name, (object)page); return sourceMargins[name].Original; }
            finally { Release((object)page); Release((object)sheet); }
        }

        List<Paper> Papers(Action<string> log)
        {
            string active = (string)app.ActivePrinter;
            string printer = Core.PrinterName(active, PrinterSettings.InstalledPrinters.Cast<string>());
            var settings = new PrinterSettings { PrinterName = printer };
            if (!settings.IsValid) throw new InvalidOperationException("纸型配置不可用：" + printer);
            log("纸型来源：" + printer + "（仅读取纸型，不发送打印任务）");
            return settings.PaperSizes.Cast<PaperSize>().Select(p => new Paper {
                Name = p.PaperName, Kind = p.RawKind, WidthPt = p.Width * 0.72, HeightPt = p.Height * 0.72 }).ToList();
        }

        static bool SinglePage(dynamic sheet, dynamic range)
        {
            sheet.DisplayPageBreaks = false; sheet.DisplayPageBreaks = true;
            var horizontal = new List<int>(); var vertical = new List<int>();
            dynamic hBreaks = null, vBreaks = null;
            try
            {
                hBreaks = sheet.HPageBreaks; vBreaks = sheet.VPageBreaks;
                for (int i = 1; i <= (int)hBreaks.Count; i++)
                {
                    dynamic item = hBreaks.Item[i]; dynamic loc = null;
                    try { loc = item.Location; horizontal.Add((int)loc.Row); }
                    finally { Release((object)loc); Release((object)item); }
                }
                for (int i = 1; i <= (int)vBreaks.Count; i++)
                {
                    dynamic item = vBreaks.Item[i]; dynamic loc = null;
                    try { loc = item.Location; vertical.Add((int)loc.Column); }
                    finally { Release((object)loc); Release((object)item); }
                }
                int firstRow = (int)range.Row, firstCol = (int)range.Column;
                int lastRow = firstRow + (int)range.Rows.Count - 1, lastCol = firstCol + (int)range.Columns.Count - 1;
                return Core.InsideBreaks(horizontal, firstRow, lastRow) == 0 && Core.InsideBreaks(vertical, firstCol, lastCol) == 0;
            }
            finally { Release((object)hBreaks); Release((object)vBreaks); }
        }

        static string ColumnLetters(int number)
        {
            string value="";
            while(number>0){number--;value=(char)('A'+number%26)+value;number/=26;}
            return value;
        }

        void ExportMeasurement(dynamic sheet,dynamic originalRange,string address,dynamic originalPage,string output,int[] colors,Action<string> log)
        {
            dynamic copy=null,range=null,page=null,borders=null;
            string originalName=(string)sheet.Name;
            Exception failure=null;
            try
            {
                int originalIndex = (int)sheet.Index;
                int sheetCount=(int)book.Worksheets.Count;
                OfficeInterop.Run("创建测量工作表副本",()=>OfficeInterop.CopyBefore((object)sheet),log);
                if ((int)book.Worksheets.Count!=sheetCount+1) throw new InvalidOperationException("表格软件未创建独立测量副本，已停止导出。");
                if ((int)sheet.Index != originalIndex + 1)
                    throw new InvalidOperationException("测量副本未插入原表前方，已停止导出。");
                copy = sheet.Previous;
                bool protectedCopy = (bool)copy.ProtectContents;
                log("原表 " + originalName + " ProtectContents=" + sheet.ProtectContents +
                    "；测量副本 " + copy.Name + " ProtectContents=" + protectedCopy);
                if (protectedCopy)
                    OfficeInterop.Run("用空密码解除测量副本保护（需要密码的工作表无法测量）", () => {
                        copy.Unprotect("");
                        if ((bool)copy.ProtectContents)
                            throw new InvalidOperationException("空密码调用后测量副本仍受保护。");
                        log("测量副本 ProtectContents=False。");
                    }, log);
                OfficeInterop.Run("读取测量副本的打印区域",()=>{
                    range=copy.Range[address];page=copy.PageSetup;
                    if (Math.Abs((double)range.Width-(double)originalRange.Width)>0.01 ||
                        Math.Abs((double)range.Height-(double)originalRange.Height)>0.01)
                        throw new InvalidOperationException("测量副本的单元格尺寸发生变化，未裁切 PDF。");
                },log);
                OfficeInterop.Run("设置测量副本纸型、缩放与页边距",()=>{
                    page.PrintArea=address;page.Zoom=100;page.FitToPagesWide=false;page.FitToPagesTall=false;
                    page.Orientation=originalPage.Orientation;page.PaperSize=originalPage.PaperSize;
                    PageMargins.FromPage((object)originalPage).Apply((object)page);
                    page.CenterHorizontally=false;page.CenterVertically=false;
                    page.BlackAndWhite=false;page.Draft=false;
                },log);
                OfficeInterop.Run("设置四边测量标记",()=>{
                    borders=range.Borders;
                    // PdfRangeMeasurement reads left, top, right, bottom.
                    int[] edgeIndices = { 7, 8, 10, 9 };
                    for (int i=0;i<edgeIndices.Length;i++)
                    {
                        dynamic edge=borders.Item[edgeIndices[i]];
                        try { edge.LineStyle=1;edge.Weight=2;edge.Color=colors[i]; }
                        finally { Release((object)edge); }
                    }
                },log);
                OfficeInterop.Run("激活测量工作表",()=>copy.Select(true),log);
                OfficeInterop.Run("导出测量 PDF",()=>OfficeInterop.ExportPdf((object)copy,output),log);
            }
            catch(Exception ex) { failure=ex; }
            finally
            {
                Release((object)borders);Release((object)range);Release((object)page);
                try { if(copy!=null)OfficeInterop.Run("删除测量工作表副本",()=>copy.Delete(),log); }
                catch(Exception cleanup)
                {
                    log(cleanup.Message);
                    if(failure==null)failure=cleanup;
                    else log("保留此前失败步骤，清理异常未覆盖原始错误。");
                }
                finally { Release((object)copy); }
                try { OfficeInterop.Run("恢复原工作表",()=>sheet.Select(true),log); }
                catch(Exception cleanup) { log(cleanup.Message);if(failure==null)failure=cleanup; }
            }
            if(failure!=null)System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(failure).Throw();
        }

        double[] ExportMeasuredPage(dynamic sheet, dynamic range, string address, dynamic page,
            string original, string measurement, string fitted, PageMargins margins, int[] colors, Action<string> log)
        {
            File.Delete(original);
            File.Delete(measurement);
            OfficeInterop.Run("导出原始 PDF", () => OfficeInterop.ExportPdf((object)sheet, original), log);
            ExportMeasurement(sheet, range, address, page, measurement, colors, log);
            log("开始：核对四边边界并调整 PDF 页面");
            double[] size = PdfPageFitter.FitMeasured(original, measurement, fitted, margins, colors);
            log("完成：核对四边边界并调整 PDF 页面");
            return size;
        }

        public void Export(string sheetName, string addressOverride, string destination, Action<string> log, bool optimizeThinLines = true, PageMargins marginSettings = null)
        {
            bool wps = EngineId == "ket.Application";
            try { ExportOnce(sheetName,addressOverride,destination,log,optimizeThinLines,marginSettings,wps); }
            catch (PaperLayoutMismatchException ex)
            {
                if (!wps) throw;
                log(ex.Message + " 用当前打印机配置重新只读打开工作簿后复核。");
                book.Close(false);
                Release((object)book); book=null;
                book=app.Workbooks.Open(FilePath,0,true);
                if (!(bool)book.ReadOnly) throw new InvalidOperationException("重新打开的工作簿不是只读。");
                sourceMargins.Clear();
                ExportOnce(sheetName,addressOverride,destination,log,optimizeThinLines,marginSettings,false);
            }
        }

        void ExportOnce(string sheetName, string addressOverride, string destination, Action<string> log, bool optimizeThinLines, PageMargins marginSettings, bool reopenOnMismatch)
        {
            dynamic sheet = null, page = null, range = null;
            string temporary = null, optimized = null, fitted = null, measurement = null;
            try
            {
                log("导出引擎："+EngineName+"；接口 "+EngineId+"；版本 "+EngineVersion+"。");
                sheet = book.Worksheets.Item[sheetName]; sheet.Select(true); page = sheet.PageSetup;
                CaptureMargins(sheetName, (object)page);
                bool useCustomMargins = marginSettings != null && marginSettings.UseCustom;
                PageMargins effectiveMargins = sourceMargins[sheetName].Resolve(marginSettings);
                string address = (string)page.PrintArea;
                if (!String.IsNullOrWhiteSpace(addressOverride)) address = Core.RangeAddress(addressOverride);
                if (String.IsNullOrWhiteSpace(address)) throw new InvalidOperationException("此工作表未设置打印区域，请填写一个连续区域，例如 A1:N20。");
                range = sheet.Range[address];
                if ((int)range.Areas.Count != 1) throw new InvalidOperationException("打印区域含多块不连续范围，请指定一块连续区域。");
                int firstRow = (int)range.Row, firstColumn = (int)range.Column;
                int lastRow=firstRow+(int)range.Rows.Count-1,lastColumn=firstColumn+(int)range.Columns.Count-1;
                address=ColumnLetters(firstColumn)+firstRow+":"+ColumnLetters(lastColumn)+lastRow;
                PrintAreaLayout.Validate((object)page, firstRow, firstRow + (int)range.Rows.Count - 1,
                    firstColumn, firstColumn + (int)range.Columns.Count - 1);
                double width = (double)range.Width, height = (double)range.Height;
                if (width <= 0 || height <= 0) throw new InvalidOperationException("打印区域没有可见尺寸。");
                double finalW = width + PageMargins.ToPoints(effectiveMargins.LeftMm + effectiveMargins.RightMm);
                double finalH = height + PageMargins.ToPoints(effectiveMargins.TopMm + effectiveMargins.BottomMm);
                // The 1-point fitting guard belongs to intermediate paper only.
                var candidates = OfficeInterop.ExportPapers(Core.Candidates(Papers(log), finalW + 1.0, finalH + 1.0),width,height);
                log(String.Format("接口预估尺寸：{0:N1} × {1:N1} 毫米；缩放为 100%。", width * 25.4 / 72, height * 25.4 / 72));
                log((useCustomMargins ? "自定义页边距：" : "沿用表格页边距：") + effectiveMargins.Description());
                if (candidates.Count == 0) throw new InvalidOperationException("现有纸型没有能以 100% 大小容纳该区域的单页纸张，未生成 PDF。");
                page.PrintArea = address; page.Zoom = 100; page.FitToPagesWide = false; page.FitToPagesTall = false;
                page.CenterVertically = false; sheet.ResetAllPageBreaks();
                Paper chosen = null; string lastReason = "";
                string dir=Path.GetDirectoryName(destination), token=Guid.NewGuid().ToString("N");
                temporary=Path.Combine(dir,"AutoPaperPDF-temp-"+token+".pdf");
                measurement=Path.Combine(dir,"AutoPaperPDF-temp-"+token+"-measure.pdf");
                fitted=Path.Combine(dir,"AutoPaperPDF-temp-"+token+"-page.pdf");
                int[] colors=PdfRangeMeasurement.CreateColors();
                foreach (var paper in candidates)
                {
                    log(String.Format("尝试中间纸型：{0}；编号 {1}；{2}；{3:N1} × {4:N1} 毫米。",paper.Name,paper.Kind,
                        paper.Orientation==1?"纵向":"横向",PageMargins.ToMm(paper.WidthPt),PageMargins.ToMm(paper.HeightPt)));
                    try
                    {
                        page.Orientation = paper.Orientation; page.PaperSize = paper.Kind;
                        sourceMargins[sheetName].ApplyForExport((object)page, marginSettings);
                        if ((int)page.PaperSize != paper.Kind || (int)page.Orientation != paper.Orientation || (int)page.Zoom != 100)
                            throw new InvalidOperationException("软件未接受纸型或缩放设置。");
                        if (!SinglePage(sheet, range)) { lastReason="原尺寸仍被分成多页。";continue; }
                    }
                    catch (Exception ex) { lastReason=OfficeInterop.ErrorText(ex);log("该纸型未通过设置检查："+lastReason);continue; }
                    // Width is only a paper-selection estimate. The exported
                    // cell-edge measurement, not that estimate, controls cropping.
                    try
                    {
                        double[] size = ExportMeasuredPage(sheet,range,address,page,temporary,measurement,
                            fitted,effectiveMargins,colors,log);
                        finalW=size[0];finalH=size[1];chosen=paper;break;
                    }
                    catch (PaperTooSmallException ex) { lastReason=ex.Message;log(lastReason); }
                    catch (PaperLayoutMismatchException ex)
                    {
                        if (reopenOnMismatch) throw;
                        lastReason=ex.Message;log(lastReason+" 尝试下一纸型。");
                    }
                }
                if (chosen == null) throw new InvalidOperationException("没有能容纳实际导出内容的纸型，原有 PDF 未覆盖。"+lastReason);
                log(String.Format("实测打印区域：{0:N2} × {1:N2} 毫米；四边边界已核对。",
                    PageMargins.ToMm(finalW)-effectiveMargins.LeftMm-effectiveMargins.RightMm,
                    PageMargins.ToMm(finalH)-effectiveMargins.TopMm-effectiveMargins.BottomMm));
                log(String.Format("中间排版纸型：{0}，{1}，{2:N1} × {3:N1} 毫米。", chosen.Name,
                    chosen.Orientation == 1 ? "纵向" : "横向", chosen.WidthPt * 25.4 / 72, chosen.HeightPt * 25.4 / 72));
                File.Delete(temporary);temporary=fitted;fitted=null;
                if (optimizeThinLines)
                {
                    optimized = Path.Combine(Path.GetDirectoryName(destination), ".autopaper-" + Guid.NewGuid().ToString("N") + "-lines.pdf");
                    int changed;
                    try { changed = PdfLineNormalizer.NormalizeWpsPdf(temporary, optimized); }
                    catch (Exception ex) { throw new InvalidOperationException("细边框优化未完成。原有 PDF 未覆盖；可取消“细边框优化”后重新导出。" + ex.Message, ex); }
                    if (changed > 0)
                    {
                        File.Delete(temporary); temporary = optimized; optimized = null;
                        log("已将 " + changed + " 条 WPS 表格细边框调整为 0.5 磅；文字、图片和页面尺寸保持不变。");
                    }
                    else log("未发现需要优化的 WPS 表格细边框，保持原线条。");
                }
                Core.CommitPdf(temporary, destination); temporary = null;
                log(String.Format("最终 PDF 页面：{0:N2} × {1:N2} 毫米；打印区域加四边留白，保持 100% 原尺寸。",
                    PageMargins.ToMm(finalW), PageMargins.ToMm(finalH)));
                log("已输出：" + destination);
                log("原工作簿未改写。留白从打印区域边界计算，区域内部的空行、空列仍保留。");
            }
            finally
            {
                foreach (string path in new[] { temporary, optimized, fitted, measurement })
                {
                    if (path == null) continue;
                    try { File.Delete(path); }
                    catch (IOException ex) { log("临时文件清理失败：" + path + "。" + ex.Message); }
                    catch (UnauthorizedAccessException ex) { log("临时文件清理失败：" + path + "。" + ex.Message); }
                }
                Release((object)range); Release((object)page); Release((object)sheet);
            }
        }

        public void Dispose()
        {
            try { if (book != null) book.Close(false); } catch { }
            try { if (owned && app != null && (int)app.Workbooks.Count == 0) app.Quit(); } catch { }
            Release((object)book); Release((object)app); book = null; app = null; owned = false;
            GC.Collect(); GC.WaitForPendingFinalizers();
        }
    }
}
