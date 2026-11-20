using System;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Windows.Forms;

[assembly: AssemblyTitle(AutoPaper.AppInfo.ProductName)]
[assembly: AssemblyDescription("按打印区域和四边留白生成原尺寸 PDF，支持本机 Excel/WPS")]
[assembly: AssemblyProduct(AutoPaper.AppInfo.ProductName)]
[assembly: AssemblyVersion(AutoPaper.AppInfo.AssemblyVersion)]
[assembly: AssemblyFileVersion(AutoPaper.AppInfo.AssemblyVersion)]
[assembly: AssemblyInformationalVersion(AutoPaper.AppInfo.InformationalVersion)]

namespace AutoPaper
{
    static class Program
    {
        [STAThread]
        static void Main()
        {
            EmbeddedDependencies.Register();
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            try { Application.Run(new MainForm()); }
            catch (Exception ex) { MessageBox.Show(ex.Message, "工具启动失败", MessageBoxButtons.OK, MessageBoxIcon.Error); }
        }
    }

    public sealed class MainForm : Form
    {
        ComboBox engine, sheets;
        TextBox range, name, prefix, suffix, preview, log;
        Button open, export, prefixDate, suffixDate, margins;
        Label path;
        CheckBox remember, thinLines;
        readonly Timer dateTimer = new Timer();
        readonly string settingsPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "AutoPaperPdf", "settings.json");
        readonly string marginsPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "AutoPaperPdf", "page-margins.json");
        PageMargins marginSettings;
        OfficeSession session;
        bool busy;

        public MainForm()
        {
            Text = AppInfo.ProductName + " " + AppInfo.DisplayVersion;
            Icon = System.Drawing.Icon.ExtractAssociatedIcon(Application.ExecutablePath);
            ShowIcon = true;
            ClientSize = new Size(730, 670); MinimumSize = new Size(746, 709);
            StartPosition = FormStartPosition.CenterScreen;
            Font = new Font("Microsoft YaHei UI", 10); AutoScaleMode = AutoScaleMode.Dpi;
            LabelAt("PDF 页面 = 打印区域 + 四边留白，保持 100% 原尺寸", 20,16,690,28);
            LabelAt("先保存表格，再选择文件。每次导出重新计算，适用于不同表格。",20,46,690,26);
            LabelAt("表格软件",20,88,85,27);
            engine = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Location = new Point(112,84), Size = new Size(170,28) };
            engine.Items.AddRange(new object[] { "自动（优先 WPS）", "WPS", "Excel" }); engine.SelectedIndex = 0; Controls.Add(engine);
            open = ButtonAt("选择表格文件…",308,82,165,32); open.Click += OpenFile;
            path = LabelAt("尚未选择文件",20,125,690,44); path.AutoEllipsis = true;
            LabelAt("工作表",20,180,85,27);
            sheets = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Location = new Point(112,176), Size = new Size(360,28), Enabled = false };
            Controls.Add(sheets); sheets.SelectedIndexChanged += SheetChanged;
            LabelAt("打印区域",20,220,85,27); range = TextAt(112,216,360,28); range.Enabled = false;
            margins = ButtonAt("页边距设置…",490,174,212,72); margins.Click += ShowMargins;
            margins.TextAlign = ContentAlignment.MiddleCenter;
            LabelAt("留空使用表格已有打印区域；也可填 A1:N20。",112,250,590,26);
            LabelAt("文件名称",20,285,85,27); name = TextAt(112,281,590,28);
            LabelAt("前缀",20,325,85,27); prefix = TextAt(112,321,460,28);
            prefixDate = ButtonAt("插入日期",588,319,114,32);
            LabelAt("后缀",20,365,85,27); suffix = TextAt(112,361,460,28);
            suffixDate = ButtonAt("插入日期",588,359,114,32);
            LabelAt("文件名预览",20,405,90,27); preview = TextAt(112,401,590,28); preview.ReadOnly = true;
            remember = new CheckBox { Text = "记住命名设置", Location = new Point(20,438), Size = new Size(195,25) }; Controls.Add(remember);
            LabelAt("日期为 yyyyMMdd；不自动添加横线或下划线。",220,439,490,26);
            export = ButtonAt("按原尺寸导出 PDF",20,475,245,38); export.Enabled = false; export.Click += Export;
            thinLines = new CheckBox { Text = "细边框优化（0.5 磅）", Checked = true,
                Location = new Point(283,481), Size = new Size(290,29) }; Controls.Add(thinLines);
            var about = ButtonAt("关于",602,475,100,38);
            about.Click += delegate {
                MessageBox.Show(this, AppInfo.ProductName + "\r\n版本：V" + AppInfo.Version +
                    "\r\n发布日期：" + AppInfo.ReleaseDateText +
                    "\r\n\r\n程序通过本机 Excel 或 WPS 按 100% 原尺寸排版，\r\n最终 PDF 页面为打印区域加四边留白。" +
                    "\r\n页边距支持沿用表格或自定义上、下、左、右毫米数，\r\n设置会自动记住，不改写原工作簿。" +
                    "\r\n细边框优化：识别 WPS 导出的约 1 磅实心表格线，\r\n自动调整为 0.5 磅。保留文字、图片和页面尺寸。" +
                    "\r\n\r\n内含 PDFsharp 6.2.4（MIT）等开源组件。\r\n版权及许可见程序内嵌 ThirdPartyNotices.txt。",
                    "关于", MessageBoxButtons.OK, MessageBoxIcon.Information);
            };
            log = TextAt(20,528,690,122); log.Multiline = true; log.ReadOnly = true; log.ScrollBars = ScrollBars.Vertical;
            log.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            var settings = Core.ReadSettings(settingsPath);
            marginSettings = PageMargins.Load(marginsPath); UpdateMarginButton();
            prefix.Text = settings.Prefix; name.Text = settings.BaseName; suffix.Text = settings.Suffix; remember.Checked = settings.Remember;
            name.TextChanged += delegate { UpdatePreview(); }; prefix.TextChanged += delegate { UpdatePreview(); }; suffix.TextChanged += delegate { UpdatePreview(); };
            prefixDate.Click += delegate { prefix.SelectedText = "{日期}"; prefix.Focus(); };
            suffixDate.Click += delegate { suffix.SelectedText = "{日期}"; suffix.Focus(); };
            Activated += delegate { UpdatePreview(); };
            dateTimer.Interval = 60000; dateTimer.Tick += delegate { UpdatePreview(); }; dateTimer.Start();
            FormClosing += ClosingForm;
            UpdatePreview();
            if (!String.IsNullOrEmpty(settings.Warning)) Log(settings.Warning);
            if (!String.IsNullOrEmpty(marginSettings.Warning)) Log(marginSettings.Warning);
        }

        Label LabelAt(string text, int x, int y, int w, int h)
        { var c = new Label { Text=text, Location=new Point(x,y), Size=new Size(w,h) }; Controls.Add(c); return c; }
        TextBox TextAt(int x, int y, int w, int h)
        { var c = new TextBox { Location=new Point(x,y), Size=new Size(w,h) }; Controls.Add(c); return c; }
        Button ButtonAt(string text,int x,int y,int w,int h)
        { var c = new Button { Text=text, Location=new Point(x,y), Size=new Size(w,h) }; Controls.Add(c); return c; }

        void Log(string text)
        { log.AppendText(text + Environment.NewLine); log.SelectionStart = log.TextLength; log.ScrollToCaret(); Application.DoEvents(); }

        void UpdatePreview()
        {
            if (String.IsNullOrWhiteSpace(name.Text)) { preview.ForeColor=SystemColors.GrayText; preview.Text="填写文件名称后显示完整文件名"; return; }
            try { preview.Text=Core.PdfName(prefix.Text,name.Text,suffix.Text,DateTime.Now); preview.ForeColor=SystemColors.WindowText; }
            catch(Exception ex) { preview.ForeColor=Color.DarkRed; preview.Text=ex.Message; }
        }

        void SaveNames()
        {
            try { Core.SaveSettings(settingsPath,new NamingSettings { Prefix=prefix.Text, BaseName=name.Text, Suffix=suffix.Text, Remember=remember.Checked }); }
            catch(Exception ex) { Log("命名设置未能保存：" + ex.Message); }
        }

        void UpdateMarginButton()
        {
            if (marginSettings == null) return;
            string title = marginSettings.UseCustom ? "页边距：自定义…" : "页边距：沿用表格…";
            PageMargins values = marginSettings;
            if (!marginSettings.UseCustom)
            {
                if (session == null || sheets.SelectedIndex < 0)
                { margins.Text = title + "\r\n选定工作表后显示数值"; return; }
                try { values = session.OriginalMargins((string)sheets.SelectedItem); }
                catch { margins.Text = title + "\r\n原页边距读取失败"; return; }
            }
            margins.Text = title + String.Format("\r\n上 {0:0.#}　下 {1:0.#} 毫米\r\n左 {2:0.#}　右 {3:0.#} 毫米",
                values.TopMm, values.BottomMm, values.LeftMm, values.RightMm);
        }

        void ShowMargins(object sender, EventArgs args)
        {
            PageMargins source = null;
            try { if (session != null && sheets.SelectedIndex >= 0) source = session.OriginalMargins((string)sheets.SelectedItem); }
            catch (Exception ex) { Log("原页边距读取失败：" + ex.Message); }
            using (var dialog = new MarginsDialog(marginSettings, source))
            {
                if (dialog.ShowDialog(this) != DialogResult.OK) return;
                marginSettings = dialog.Result; UpdateMarginButton();
                try { PageMargins.Save(marginsPath, marginSettings); }
                catch (Exception ex) { Log("本次页边距已设置，但未能记住：" + ex.Message); }
                Log(marginSettings.UseCustom ? "自定义页边距：" + marginSettings.Description() : "页边距：沿用所选工作表原有设置。");
            }
        }

        void SetBusy(bool value)
        {
            busy=value; UseWaitCursor=value; open.Enabled=!value; engine.Enabled=!value;
            sheets.Enabled=!value && session!=null; range.Enabled=sheets.Enabled; export.Enabled=sheets.Enabled;
            foreach(Control c in new Control[]{name,prefix,suffix,prefixDate,suffixDate,remember,thinLines,margins}) c.Enabled=!value;
        }

        void SheetChanged(object sender,EventArgs args)
        {
            if(session==null || sheets.SelectedIndex<0) { UpdateMarginButton(); return; }
            try { range.Clear(); string current=session.PrintArea((string)sheets.SelectedItem);
                Log(String.IsNullOrWhiteSpace(current) ? "所选工作表未设置打印区域，请填写连续区域。" : "当前工作表的打印区域："+current); }
            catch(Exception ex) { Log(ex.Message); }
            finally { UpdateMarginButton(); }
        }

        void OpenFile(object sender,EventArgs args)
        {
            string chosen;
            using(var dialog=new OpenFileDialog { Filter="表格文件|*.xlsx;*.xlsm;*.xlsb;*.xls;*.et|所有文件|*.*" })
            { if(dialog.ShowDialog(this)!=DialogResult.OK) return; chosen=dialog.FileName; }
            SetBusy(true);
            try
            {
                if(session!=null) session.Dispose(); session=null;
                sheets.Items.Clear(); log.Clear(); path.Text="正在读取…"; Log("正在以只读方式打开文件…");
                session=new OfficeSession(chosen,(string)engine.SelectedItem); path.Text=chosen;
                if(String.IsNullOrWhiteSpace(name.Text)) name.Text=Path.GetFileNameWithoutExtension(chosen);
                var names=session.SheetNames(); if(names.Count==0) throw new InvalidOperationException("文件没有可见工作表。");
                sheets.Items.AddRange(names.ToArray()); string active=session.ActiveSheetName();
                if(sheets.Items.Contains(active)) sheets.SelectedItem=active; else sheets.SelectedIndex=0;
                Log("已连接："+session.EngineName+"。原文件只读，未请求更新外部链接。");
            }
            catch(Exception ex) { Log("读取失败："+ex.Message); if(session!=null) session.Dispose(); session=null; path.Text="未加载文件"; }
            finally { SetBusy(false); UpdateMarginButton(); }
        }

        void Export(object sender,EventArgs args)
        {
            if(session==null || sheets.SelectedIndex<0) return;
            string filename, destination;
            try { filename=Core.PdfName(prefix.Text,name.Text,suffix.Text,DateTime.Now); UpdatePreview(); }
            catch(Exception ex) { MessageBox.Show(this,ex.Message,"请修改文件名",MessageBoxButtons.OK,MessageBoxIcon.Warning); return; }
            SaveNames();
            using(var dialog=new SaveFileDialog { Filter="PDF 文件|*.pdf",DefaultExt="pdf",AddExtension=true,OverwritePrompt=true,
                FileName=filename,InitialDirectory=Path.GetDirectoryName(session.FilePath) })
            { if(dialog.ShowDialog(this)!=DialogResult.OK) return; destination=dialog.FileName; }
            SetBusy(true);
            try { session.Export((string)sheets.SelectedItem,range.Text,destination,Log,thinLines.Checked,marginSettings);
                MessageBox.Show(this,"已导出："+destination+"\r\n原表格未改写。","导出完成",MessageBoxButtons.OK,MessageBoxIcon.Information); }
            catch(Exception ex) { Log("导出停止："+ex.Message); MessageBox.Show(this,ex.Message,"未完成导出",MessageBoxButtons.OK,MessageBoxIcon.Warning); }
            finally { SetBusy(false); }
        }

        void ClosingForm(object sender,FormClosingEventArgs args)
        {
            if(busy) { args.Cancel=true; return; }
            dateTimer.Stop(); SaveNames(); if(session!=null) session.Dispose(); session=null;
        }
        protected override void Dispose(bool disposing)
        {
            if(disposing) { dateTimer.Dispose(); if(session!=null) session.Dispose(); session=null; }
            base.Dispose(disposing);
        }
    }
}
