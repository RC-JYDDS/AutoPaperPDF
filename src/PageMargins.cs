using System;
using System.IO;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;

namespace AutoPaper
{
    public sealed class PageMarginState
    {
        readonly PageMargins original;
        readonly bool centered;

        public PageMarginState(object value)
        {
            dynamic page = value;
            original = PageMargins.FromPage(value); centered = (bool)page.CenterHorizontally;
        }

        public PageMargins Original { get { return original.Copy(); } }

        public PageMargins Resolve(PageMargins settings)
        {
            if (settings == null || !settings.UseCustom) return Original;
            settings.Validate(); return settings.Copy();
        }

        public void ApplyForExport(object value, PageMargins settings)
        {
            dynamic page = value;
            PageMargins effective = Resolve(settings); effective.Apply(value);
            page.CenterHorizontally = settings != null && settings.UseCustom ? false : centered;
            page.CenterVertically = false;
            effective.Verify(value);
        }
    }

    [DataContract]
    public sealed class PageMargins
    {
        [DataMember] public bool UseCustom = false;
        [DataMember] public double TopMm = 5;
        [DataMember] public double BottomMm = 5;
        [DataMember] public double LeftMm = 5;
        [DataMember] public double RightMm = 5;
        public string Warning = "";

        [OnDeserializing]
        void Initialize(StreamingContext context)
        { TopMm = BottomMm = LeftMm = RightMm = 5; Warning = ""; }

        public PageMargins Copy()
        { return new PageMargins { UseCustom = UseCustom, TopMm = TopMm, BottomMm = BottomMm, LeftMm = LeftMm, RightMm = RightMm }; }

        public static double ToPoints(double mm) { return mm * 72.0 / 25.4; }
        public static double ToMm(double points) { return points * 25.4 / 72.0; }

        public void Validate()
        {
            foreach (double value in new[] { TopMm, BottomMm, LeftMm, RightMm })
                if (Double.IsNaN(value) || Double.IsInfinity(value) || value < 0 || value > 100)
                    throw new ArgumentException("页边距必须在 0 至 100 毫米之间。");
        }

        public string Description()
        { return String.Format("上 {0:0.#}、下 {1:0.#}、左 {2:0.#}、右 {3:0.#} 毫米", TopMm, BottomMm, LeftMm, RightMm); }

        public static PageMargins FromPage(object value)
        {
            dynamic page = value;
            return new PageMargins { TopMm = ToMm((double)page.TopMargin), BottomMm = ToMm((double)page.BottomMargin),
                LeftMm = ToMm((double)page.LeftMargin), RightMm = ToMm((double)page.RightMargin) };
        }

        public void Apply(object value)
        {
            dynamic page = value;
            page.TopMargin = ToPoints(TopMm); page.BottomMargin = ToPoints(BottomMm);
            page.LeftMargin = ToPoints(LeftMm); page.RightMargin = ToPoints(RightMm);
        }

        public void Verify(object value)
        {
            PageMargins actual = FromPage(value);
            if (Math.Abs(actual.TopMm - TopMm) > 0.05 || Math.Abs(actual.BottomMm - BottomMm) > 0.05 ||
                Math.Abs(actual.LeftMm - LeftMm) > 0.05 || Math.Abs(actual.RightMm - RightMm) > 0.05)
                throw new InvalidOperationException("表格软件未接受指定的页边距。请调整边距或检查纸型设置。");
        }

        public static PageMargins Load(string path)
        {
            if (!File.Exists(path)) return new PageMargins();
            try
            {
                using (var stream = File.OpenRead(path))
                {
                    var result = (PageMargins)new DataContractJsonSerializer(typeof(PageMargins)).ReadObject(stream);
                    if (result == null) throw new InvalidDataException();
                    result.Validate(); return result;
                }
            }
            catch { return new PageMargins { Warning = "原页边距设置无法读取，本次沿用表格设置。" }; }
        }

        public static void Save(string path, PageMargins settings)
        {
            settings.Validate();
            string folder = Path.GetDirectoryName(path); Directory.CreateDirectory(folder);
            string temporary = Path.Combine(folder, Guid.NewGuid().ToString("N") + ".tmp");
            try
            {
                using (var stream = File.Create(temporary))
                    new DataContractJsonSerializer(typeof(PageMargins)).WriteObject(stream, settings);
                if (File.Exists(path)) File.Replace(temporary, path, null); else File.Move(temporary, path);
            }
            finally { if (File.Exists(temporary)) File.Delete(temporary); }
        }
    }
}
