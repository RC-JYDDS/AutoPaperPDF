using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Text.RegularExpressions;

namespace AutoPaper
{
    [DataContract]
    public sealed class NamingSettings
    {
        [DataMember] public int Version = 2;
        [DataMember] public bool Remember = true;
        [DataMember] public string Prefix = "{日期}";
        [DataMember] public string BaseName = "";
        [DataMember] public string Suffix = "";
        public string Warning = "";
    }

    public sealed class Paper
    {
        public string Name;
        public int Kind, Orientation;
        public double WidthPt, HeightPt;
        public double Area { get { return WidthPt * HeightPt; } }
    }

    public static class Core
    {
        public static string PdfName(string prefix, string name, string suffix, DateTime now)
        {
            name = (name ?? "").Trim();
            if (name.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase)) name = name.Substring(0, name.Length - 4).Trim();
            if (String.IsNullOrWhiteSpace(name)) throw new ArgumentException("请填写文件名称。");
            string date = now.ToString("yyyyMMdd", CultureInfo.InvariantCulture);
            string result = (prefix ?? "").Replace("{日期}", date) + name + (suffix ?? "").Replace("{日期}", date) + ".pdf";
            if (Regex.IsMatch(result, "[\\x00-\\x1f\\\\/:*?\"<>|]"))
                throw new ArgumentException("文件名不能包含 \\ / : * ? \" < > | 或换行。");
            if (Regex.IsMatch(result, @"^(CON|PRN|AUX|NUL|COM[1-9¹²³]|LPT[1-9¹²³])(?:\.|$)", RegexOptions.IgnoreCase))
                throw new ArgumentException("该名称属于 Windows 保留名称，请换一个文件名。");
            if (result.Length > 255) throw new ArgumentException("完整文件名超过 255 个字符，请缩短名称或前后缀。");
            return result;
        }

        public static NamingSettings ReadSettings(string path)
        {
            var defaults = new NamingSettings();
            if (!File.Exists(path)) return defaults;
            try
            {
                var serializer = new DataContractJsonSerializer(typeof(NamingSettings));
                using (var stream = new MemoryStream(Encoding.UTF8.GetBytes(File.ReadAllText(path, Encoding.UTF8))))
                {
                    var data = (NamingSettings)serializer.ReadObject(stream);
                    if (data == null || (data.Version != 1 && data.Version != 2) || data.Prefix == null || data.Suffix == null ||
                        (data.Version == 2 && data.BaseName == null)) throw new InvalidDataException();
                    data.BaseName = data.BaseName ?? "";
                    data.Warning = "";
                    if (!data.Remember) { data.Prefix = ""; data.BaseName = ""; data.Suffix = ""; }
                    return data;
                }
            }
            catch { defaults.Warning = "原命名设置无法读取，本次已使用默认设置。"; return defaults; }
        }

        public static void SaveSettings(string path, NamingSettings settings)
        {
            string folder = Path.GetDirectoryName(path);
            Directory.CreateDirectory(folder);
            string temporary = Path.Combine(folder, Guid.NewGuid().ToString("N") + ".tmp");
            var data = new NamingSettings { Version = 2, Remember = settings.Remember,
                Prefix = settings.Remember ? (settings.Prefix ?? "") : "",
                BaseName = settings.Remember ? (settings.BaseName ?? "") : "",
                Suffix = settings.Remember ? (settings.Suffix ?? "") : "" };
            try
            {
                using (var stream = File.Create(temporary)) new DataContractJsonSerializer(typeof(NamingSettings)).WriteObject(stream, data);
                if (File.Exists(path)) File.Replace(temporary, path, null); else File.Move(temporary, path);
            }
            finally { if (File.Exists(temporary)) File.Delete(temporary); }
        }

        public static List<Paper> Candidates(IEnumerable<Paper> papers, double width, double height)
        {
            if (width <= 0 || height <= 0 || Double.IsNaN(width) || Double.IsNaN(height)) throw new ArgumentException("打印区域尺寸无效。");
            var result = new List<Paper>(); var seen = new HashSet<string>();
            foreach (var paper in papers)
            {
                if (paper.Kind <= 0 || paper.WidthPt <= 0 || paper.HeightPt <= 0) continue;
                for (int orientation = 1; orientation <= 2; orientation++)
                {
                    string key = paper.Kind + "/" + orientation;
                    if (!seen.Add(key)) continue;
                    double w = orientation == 1 ? paper.WidthPt : paper.HeightPt;
                    double h = orientation == 1 ? paper.HeightPt : paper.WidthPt;
                    if (w + 0.01 < width || h + 0.01 < height) continue;
                    result.Add(new Paper { Name = paper.Name, Kind = paper.Kind, Orientation = orientation, WidthPt = w, HeightPt = h });
                }
            }
            return result.OrderBy(p => p.Area).ThenBy(p => p.WidthPt).ThenBy(p => p.HeightPt).ThenBy(p => p.Kind).ToList();
        }

        public static string RangeAddress(string text)
        {
            text = (text ?? "").Trim();
            if (!Regex.IsMatch(text, @"^\$?[A-Za-z]{1,3}\$?[1-9][0-9]*(?::\$?[A-Za-z]{1,3}\$?[1-9][0-9]*)?$"))
                throw new ArgumentException("请填写一个连续区域，例如 A1:N20；不支持多块区域或整行、整列。");
            return text.ToUpperInvariant();
        }

        public static int InsideBreaks(IEnumerable<int> positions, int first, int last)
        { return positions.Count(p => p > first && p <= last); }

        public static string PrinterName(string active, IEnumerable<string> installed)
        {
            var names = installed.ToList();
            var exact = names.FirstOrDefault(n => String.Equals(n, active, StringComparison.OrdinalIgnoreCase));
            if (exact != null) return exact;
            var prefix = names.Where(n => active.StartsWith(n + " ", StringComparison.OrdinalIgnoreCase)).OrderByDescending(n => n.Length).FirstOrDefault();
            if (prefix != null) return prefix;
            throw new InvalidOperationException("无法匹配表格软件的纸型配置：" + active);
        }

        public static void CommitPdf(string temporary, string destination)
        {
            if (!File.Exists(temporary) || new FileInfo(temporary).Length < 100) throw new InvalidDataException("软件没有生成有效的 PDF。");
            using (var stream = File.OpenRead(temporary))
            {
                byte[] header = new byte[5];
                if (stream.Read(header, 0, 5) != 5 || Encoding.ASCII.GetString(header) != "%PDF-") throw new InvalidDataException("输出文件不是 PDF。");
            }
            if (File.Exists(destination)) File.Replace(temporary, destination, null); else File.Move(temporary, destination);
        }
    }
}
