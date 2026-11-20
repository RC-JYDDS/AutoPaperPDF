using System;
using System.IO;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

namespace AutoPaper
{
    // Office exports at 100% onto a supported paper first. Keep its displayed
    // top-left origin, and remove only the surplus paper on the right/bottom.
    // Never import, redraw, scale or rasterize the page contents.
    public static class PdfPageFitter
    {
        const double Tolerance = 0.01; // points, for PDF numeric serialization

        static bool Positive(double value)
        { return !Double.IsNaN(value) && !Double.IsInfinity(value) && value > 0; }

        static int Rotation(PdfPage page)
        {
            int value = ((page.Rotate % 360) + 360) % 360;
            if (value % 90 != 0) throw new InvalidDataException("PDF 页面旋转角度不受支持，未调整页面。");
            return value;
        }

        static double UserUnit(PdfPage page)
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

        public static void Fit(string input, string output, double widthPt, double heightPt)
        {
            if (!Positive(widthPt) || !Positive(heightPt)) throw new ArgumentException("最终 PDF 的宽度和高度必须为正数。");
            if (String.Equals(Path.GetFullPath(input), Path.GetFullPath(output), StringComparison.OrdinalIgnoreCase))
                throw new ArgumentException("页面调整须使用临时输出文件。");
            PdfRectangle target;
            int rotation;
            double unit;
            using (var document = PdfReader.Open(input, PdfDocumentOpenMode.Modify))
            {
                if (document.PageCount != 1) throw new InvalidDataException("表格软件实际导出了多页，已停止页面调整，原有 PDF 未覆盖。");
                if (document.Internals.Catalog.Elements.ContainsKey("/Perms") || document.Internals.Catalog.Elements.ContainsKey("/AcroForm"))
                    throw new InvalidDataException("PDF 含有表单或签名信息，未调整页面。");
                PdfPage page = document.Pages[0];
                PdfRectangle box = page.MediaBox;
                unit = UserUnit(page); rotation = Rotation(page);
                if (!Positive(box.Width) || !Positive(box.Height)) throw new InvalidDataException("PDF 原页面尺寸无效。");
                if (page.HasCropBox && !SameBox(page.CropBox, box, unit))
                    throw new InvalidDataException("表格软件输出的 PDF 已有不同的裁切边界，无法确认原始内容位置，未调整页面。");
                bool sideways = rotation == 90 || rotation == 270;
                double availableW = (sideways ? box.Height : box.Width) * unit;
                double availableH = (sideways ? box.Width : box.Height) * unit;
                if (widthPt > availableW + Tolerance || heightPt > availableH + Tolerance)
                    throw new InvalidDataException("表格软件输出的纸张小于所需页面，未缩放或裁切表格。");

                double w = widthPt / unit, h = heightPt / unit;
                double x = box.X1, y = box.Y1, tw = w, th = h;
                switch (rotation)
                {
                    case 0: y = box.Y2 - h; break;
                    case 90: tw = h; th = w; break;
                    case 180: x = box.X2 - w; break;
                    case 270: x = box.X2 - h; y = box.Y2 - w; tw = h; th = w; break;
                }
                target = new PdfRectangle(new XPoint(x, y), new XSize(tw, th));
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
                if (check.PageCount != 1 || Rotation(check.Pages[0]) != rotation ||
                    Math.Abs(UserUnit(check.Pages[0]) - unit) > 0.000001 ||
                    !SameBox(check.Pages[0].MediaBox, target, unit) || !SameBox(check.Pages[0].CropBox, target, unit))
                    throw new InvalidDataException("最终 PDF 页面尺寸校验失败，原有 PDF 未覆盖。");
            }
        }
    }
}
