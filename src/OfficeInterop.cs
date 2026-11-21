using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;

namespace AutoPaper
{
    public static class OfficeInterop
    {
        public static string ErrorText(Exception error)
        {
            Exception root=error;
            while(root.InnerException!=null)root=root.InnerException;
            return String.Format("{0}（0x{1:X8}）",root.Message,unchecked((uint)root.HResult));
        }

        public static void Run(string stage,Action action,Action<string> log)
        {
            log("开始："+stage);
            try { action();log("完成："+stage); }
            catch(Exception ex)
            { throw new InvalidOperationException("失败步骤："+stage+"。"+ErrorText(ex),ex); }
        }

        public static void ExportPdf(object value,string output)
        {
            dynamic sheet=value;
            // Omit unused optional arguments. In particular, do not marshal
            // Missing.Value for From/To into older spreadsheet COM servers.
            // Keep IgnorePrintAreas=false explicit; never export only page 1.
            sheet.ExportAsFixedFormat(0,output,0,true,false);
            if(!File.Exists(output) || new FileInfo(output).Length<100)
                throw new InvalidDataException("表格软件未生成有效 PDF 文件。");
            using(var stream=File.OpenRead(output))
            {
                var header=new byte[5];
                if(stream.Read(header,0,5)!=5 || System.Text.Encoding.ASCII.GetString(header)!="%PDF-")
                    throw new InvalidDataException("表格软件返回的文件不是 PDF。");
            }
        }

        public static void CopyBefore(object value)
        {
            dynamic sheet=value;
            // The temporary sheet is deleted later, so insertion order does
            // not matter. A single Before argument avoids a missing argument.
            sheet.Copy(sheet);
        }

        public static List<Paper> ExportPapers(IEnumerable<Paper> papers,double width,double height)
        {
            // Final PDF dimensions are measured independently. Use a common
            // intermediate paper before printer-specific/custom small forms.
            int orientation=width>=height?2:1;
            return papers.OrderBy(p=>p.Kind==9?0:p.Kind==8?1:p.Kind==1?2:p.Kind==5?3:4)
                .ThenBy(p=>p.Orientation==orientation?0:1).ThenBy(p=>p.Area).ToList();
        }
    }
}
