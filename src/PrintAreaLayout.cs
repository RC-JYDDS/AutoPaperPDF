using System;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Microsoft.CSharp.RuntimeBinder;

namespace AutoPaper
{
    public static class PrintAreaLayout
    {
        public static void Validate(object value, int firstRow, int lastRow, int firstColumn, int lastColumn)
        {
            if (firstRow < 1 || lastRow < firstRow || lastRow > 1048576 ||
                firstColumn < 1 || lastColumn < firstColumn || lastColumn > 16384)
                throw new ArgumentException("打印区域边界无效。");
            dynamic page = value;
            string[] names = { "左页眉", "中页眉", "右页眉", "左页脚", "中页脚", "右页脚" };
            string[] texts = { page.LeftHeader, page.CenterHeader, page.RightHeader,
                page.LeftFooter, page.CenterFooter, page.RightFooter };
            for (int i = 0; i < texts.Length; i++)
                if (!String.IsNullOrWhiteSpace(texts[i])) throw Unsupported(names[i] + "含有内容");
            // Title rows/columns already inside the selected area are ordinary
            // table content on this one-page export, not additional headers.
            ValidateTitle((string)page.PrintTitleRows, true, firstRow, lastRow);
            ValidateTitle((string)page.PrintTitleColumns, false, firstColumn, lastColumn);
            if ((bool)page.PrintHeadings) throw Unsupported("启用了打印行列标号");
            if (SpecialHeaders(value, true)) throw Unsupported("启用了首页不同的页眉页脚");
            if (SpecialHeaders(value, false)) throw Unsupported("启用了奇偶页不同的页眉页脚");
        }

        static void ValidateTitle(string title, bool rows, int first, int last)
        {
            if (String.IsNullOrWhiteSpace(title)) return;
            string range = title.Trim();
            int bang = range.LastIndexOf('!');
            if (bang >= 0) range = range.Substring(bang + 1);
            string pattern = rows ? @"^\$?([1-9][0-9]*):\$?([1-9][0-9]*)$" : @"^\$?([A-Za-z]{1,3}):\$?([A-Za-z]{1,3})$";
            Match match = Regex.Match(range, pattern);
            int begin = 0, end = 0;
            bool valid = match.Success;
            if (valid && rows)
                valid = Int32.TryParse(match.Groups[1].Value, out begin) && Int32.TryParse(match.Groups[2].Value, out end);
            else if (valid)
            { begin = ColumnNumber(match.Groups[1].Value); end = ColumnNumber(match.Groups[2].Value); }
            string name = rows ? "重复打印标题行" : "重复打印标题列";
            if (!valid || begin > end) throw Unsupported(name + "的范围无法识别：" + title);
            if (begin < first || end > last) throw Unsupported(name + " " + title + " 不完全位于当前打印区域内");
        }

        static int ColumnNumber(string letters)
        {
            int result = 0;
            foreach (char letter in letters.ToUpperInvariant()) result = result * 26 + letter - 'A' + 1;
            return result;
        }

        static Exception Unsupported(string reason)
        {
            return new InvalidOperationException("未导出：" + reason + "。精确页面按选定打印区域计算；请调整相关设置或打印区域后重试，避免遗漏额外内容。");
        }

        static bool SpecialHeaders(object value, bool firstPage)
        {
            dynamic page = value;
            try { return firstPage ? (bool)page.DifferentFirstPageHeaderFooter : (bool)page.OddAndEvenPagesHeaderFooter; }
            // Older WPS interfaces may not implement these optional features.
            catch (RuntimeBinderException) { return false; }
            catch (COMException ex)
            {
                if (ex.ErrorCode == unchecked((int)0x80020003) || ex.ErrorCode == unchecked((int)0x80020006)) return false;
                throw;
            }
        }
    }
}
