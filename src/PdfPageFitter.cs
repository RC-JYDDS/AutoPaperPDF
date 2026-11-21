using System;
using System.IO;
using System.Globalization;
using System.Text;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

namespace AutoPaper
{
    // Crop to a measured cell range, not the COM Width/Height estimate.
    // Never import, redraw, scale or rasterize the page contents.
    public static class PdfPageFitter
    {
        const double Tolerance = 0.01; // points, for PDF numeric serialization

        static bool Positive(double value)
        { return !Double.IsNaN(value) && !Double.IsInfinity(value) && value > 0; }

        internal static int PageRotation(PdfPage page)
        {
            int value = ((page.Rotate % 360) + 360) % 360;
            if (value % 90 != 0) throw new InvalidDataException("PDF 页面旋转角度不受支持，未调整页面。");
            return value;
        }

        internal static double PageUnit(PdfPage page)
        {
            double value = page.Elements.ContainsKey("/UserUnit") ? page.Elements.GetReal("/UserUnit") : 1;
            if (!Positive(value)) throw new InvalidDataException("PDF 页面单位无效。");
            return value;
        }

        static bool SameBox(PdfRectangle a, PdfRectangle b, double unit)
        {
            return Math.Abs(a.X1 - b.X1) * unit <= Tolerance && Math.Abs(a.Y1 - b.Y1) * unit <= Tolerance &&
                Math.Abs(a.X2 - b.X2) * unit <= Tolerance && Math.Abs(a.Y2 - b.Y2) * unit <= Tolerance;
        }

        public static double[] FitMeasured(string input, string measurement, string output, PageMargins margins, int[] colors)
        {
            if (margins == null) throw new ArgumentNullException("margins");
            margins.Validate();
            if (String.Equals(Path.GetFullPath(input), Path.GetFullPath(output), StringComparison.OrdinalIgnoreCase))
                throw new ArgumentException("页面调整须使用临时输出文件。");
            if (String.Equals(Path.GetFullPath(measurement), Path.GetFullPath(output), StringComparison.OrdinalIgnoreCase))
                throw new ArgumentException("不能覆盖边界测量文件。");
            PdfRectangle target;
            int rotation;
            double unit;
            double widthPt, heightPt;
            using (var document = PdfReader.Open(input, PdfDocumentOpenMode.Modify))
            using (var probe = PdfReader.Open(measurement, PdfDocumentOpenMode.Modify))
            {
                if (document.PageCount != 1 || probe.PageCount != 1) throw new PaperTooSmallException("表格软件实际导出了多页，需要更大的中间纸型。");
                if (document.Internals.Catalog.Elements.ContainsKey("/Perms") || document.Internals.Catalog.Elements.ContainsKey("/AcroForm"))
                    throw new InvalidDataException("PDF 含有表单或签名信息，未调整页面。");
                PdfPage page = document.Pages[0];
                PdfRectangle box = page.MediaBox;
                unit = PageUnit(page); rotation = PageRotation(page);
                if (!Positive(box.Width) || !Positive(box.Height)) throw new InvalidDataException("PDF 原页面尺寸无效。");
                if (page.HasCropBox && !SameBox(page.CropBox, box, unit))
                    throw new InvalidDataException("表格软件输出的 PDF 已有不同的裁切边界，无法确认原始内容位置，未调整页面。");
                bool sideways = rotation == 90 || rotation == 270;
                double availableW = (sideways ? box.Height : box.Width) * unit;
                double availableH = (sideways ? box.Width : box.Height) * unit;
                var pp = probe.Pages[0];
                if (PageRotation(pp) != rotation || Math.Abs(PageUnit(pp)-unit)>0.000001 || !SameBox(pp.MediaBox,box,unit) ||
                    (pp.HasCropBox && !SameBox(pp.CropBox,box,unit)))
                    throw new PaperLayoutMismatchException("边界测量与原 PDF 的页面不一致，未裁切 PDF。");
                var anchors = new System.Collections.Generic.List<XPoint>();
                XRect frame = PdfRangeMeasurement.Read(pp,colors,anchors);
                if (frame.Left < -0.6 || frame.Top < -0.6 || frame.Right > availableW+0.6 || frame.Bottom > availableH+0.6)
                    throw new PaperTooSmallException("实际打印区域超出中间纸张，需要更大的纸型。");
                XRect bounds = PdfRangeMeasurement.PageWithMargins(page,colors,frame,margins,anchors);
                widthPt = bounds.Width;
                heightPt = bounds.Height;
                if (!Positive(widthPt) || !Positive(heightPt)) throw new InvalidDataException("实际打印区域尺寸无效。");
                double left = bounds.Left, top = bounds.Top;
                // Do not rescale content. A shifted MediaBox places the measured
                // range at exactly the requested margins, even if Office rounded
                // its margin values or used a different printed column width.
                double w=widthPt/unit,h=heightPt/unit,l=left/unit,t=top/unit;
                double x=box.X1+l,y=box.Y2-t-h,tw=w,th=h;
                switch (rotation)
                {
                    case 90: x=box.X1+t;y=box.Y1+l;tw=h;th=w;break;
                    case 180: x=box.X2-l-w;y=box.Y1+t;break;
                    case 270: x=box.X2-t-h;y=box.Y2-l-w;tw=h;th=w;break;
                }
                target = new PdfRectangle(new XPoint(x, y), new XSize(tw, th));
                // If exact margins extend beyond the old paper, preserve its
                // original clipping so previously invisible ink is not revealed.
                if(target.X1<box.X1 || target.Y1<box.Y1 || target.X2>box.X2 || target.Y2>box.Y2)
                {
                    string clip=String.Format(CultureInfo.InvariantCulture,"q {0:R} {1:R} {2:R} {3:R} re W n\n",box.X1,box.Y1,box.Width,box.Height);
                    page.Contents.PrependContent().CreateStream(Encoding.ASCII.GetBytes(clip));
                    page.Contents.AppendContent().CreateStream(Encoding.ASCII.GetBytes("\nQ\n"));
                }
                // MediaBox controls the physical PDF paper size, not just the
                // visible crop. All page boxes agree; page streams stay intact.
                page.MediaBox = target;
                page.CropBox = target;
                if (page.HasTrimBox) page.TrimBox = target;
                if (page.HasBleedBox) page.BleedBox = target;
                if (page.HasArtBox) page.ArtBox = target;
                document.Save(output);
            }
            using (var check = PdfReader.Open(output, PdfDocumentOpenMode.Import))
            {
                if (check.PageCount != 1 || PageRotation(check.Pages[0]) != rotation ||
                    Math.Abs(PageUnit(check.Pages[0]) - unit) > 0.000001 ||
                    !SameBox(check.Pages[0].MediaBox, target, unit) || !SameBox(check.Pages[0].CropBox, target, unit))
                    throw new InvalidDataException("最终 PDF 页面尺寸校验失败，原有 PDF 未覆盖。");
            }
            return new[]{widthPt,heightPt};
        }
    }
    public sealed class PaperTooSmallException : IOException
    { public PaperTooSmallException(string message):base(message){} }
    internal sealed class PaperLayoutMismatchException : IOException
    { public PaperLayoutMismatchException(string message):base(message){} }
}
