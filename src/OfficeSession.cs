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
                    app = Activator.CreateInstance(type); break;
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

        public void Export(string sheetName, string addressOverride, string destination, Action<string> log, bool optimizeThinLines = true, PageMargins marginSettings = null)
        {
            dynamic sheet = null, page = null, range = null;
            string temporary = null, optimized = null, fitted = null;
            try
            {
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
                PrintAreaLayout.Validate((object)page, firstRow, firstRow + (int)range.Rows.Count - 1,
                    firstColumn, firstColumn + (int)range.Columns.Count - 1);
                double width = (double)range.Width, height = (double)range.Height;
                if (width <= 0 || height <= 0) throw new InvalidOperationException("打印区域没有可见尺寸。");
                double finalW = width + PageMargins.ToPoints(effectiveMargins.LeftMm + effectiveMargins.RightMm);
                double finalH = height + PageMargins.ToPoints(effectiveMargins.TopMm + effectiveMargins.BottomMm);
                // The 1-point fitting guard belongs to intermediate paper only.
                var candidates = Core.Candidates(Papers(log), finalW + 1.0, finalH + 1.0);
                log(String.Format("内容尺寸：{0:N1} × {1:N1} 毫米；缩放为 100%。", width * 25.4 / 72, height * 25.4 / 72));
                log((useCustomMargins ? "自定义页边距：" : "沿用表格页边距：") + effectiveMargins.Description());
                if (candidates.Count == 0) throw new InvalidOperationException("现有纸型没有能以 100% 大小容纳该区域的单页纸张，未生成 PDF。");
                page.PrintArea = address; page.Zoom = 100; page.FitToPagesWide = false; page.FitToPagesTall = false;
                page.CenterVertically = false; sheet.ResetAllPageBreaks();
                Paper chosen = null; string lastReason = "";
                foreach (var paper in candidates)
                {
                    try
                    {
                        page.Orientation = paper.Orientation; page.PaperSize = paper.Kind;
                        sourceMargins[sheetName].ApplyForExport((object)page, marginSettings);
                        if ((int)page.PaperSize != paper.Kind || (int)page.Orientation != paper.Orientation || (int)page.Zoom != 100)
                            throw new InvalidOperationException("软件未接受纸型或缩放设置。");
                        if (SinglePage(sheet, range)) { chosen = paper; break; }
                        lastReason = "原尺寸仍被分成多页。";
                    }
                    catch (Exception ex) { lastReason = ex.Message; }
                }
                if (chosen == null) throw new InvalidOperationException("没有通过原尺寸分页检查的纸型，未生成 PDF。" + lastReason);
                log(String.Format("中间排版纸型：{0}，{1}，{2:N1} × {3:N1} 毫米。", chosen.Name,
                    chosen.Orientation == 1 ? "纵向" : "横向", chosen.WidthPt * 25.4 / 72, chosen.HeightPt * 25.4 / 72));
                temporary = Path.Combine(Path.GetDirectoryName(destination), ".autopaper-" + Guid.NewGuid().ToString("N") + ".pdf");
                sheet.ExportAsFixedFormat(0, temporary, 0, true, false, Type.Missing, Type.Missing, false);
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
                fitted = Path.Combine(Path.GetDirectoryName(destination), ".autopaper-" + Guid.NewGuid().ToString("N") + "-page.pdf");
                PdfPageFitter.Fit(temporary, fitted, finalW, finalH);
                File.Delete(temporary); temporary = fitted; fitted = null;
                Core.CommitPdf(temporary, destination); temporary = null;
                log(String.Format("最终 PDF 页面：{0:N2} × {1:N2} 毫米；打印区域加四边留白，保持 100% 原尺寸。",
                    PageMargins.ToMm(finalW), PageMargins.ToMm(finalH)));
                log("已输出：" + destination);
                log("原工作簿未改写。留白从打印区域边界计算，区域内部的空行、空列仍保留。");
            }
            finally
            {
                if (temporary != null && File.Exists(temporary)) File.Delete(temporary);
                if (optimized != null && File.Exists(optimized)) File.Delete(optimized);
                if (fitted != null && File.Exists(fitted)) File.Delete(fitted);
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
