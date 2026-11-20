using System;
using System.Drawing;
using System.Windows.Forms;

namespace AutoPaper
{
    public sealed class MarginsDialog : Form
    {
        RadioButton original, custom;
        NumericUpDown top, bottom, left, right;
        Panel values;
        public PageMargins Result { get; private set; }

        public MarginsDialog(PageMargins settings, PageMargins source)
        {
            Text = "页边距设置"; ClientSize = new Size(520, 360);
            FormBorderStyle = FormBorderStyle.FixedDialog; MaximizeBox = false; MinimizeBox = false;
            StartPosition = FormStartPosition.CenterParent; ShowInTaskbar = false;
            Font = new Font("Microsoft YaHei UI", 10); AutoScaleMode = AutoScaleMode.Dpi;
            original = new RadioButton { Text = "沿用表格中的页边距", Location = new Point(20, 18), Size = new Size(460, 28) };
            custom = new RadioButton { Text = "自定义页边距（毫米）", Location = new Point(20, 82), Size = new Size(460, 28) };
            Controls.Add(original); Controls.Add(custom);
            Controls.Add(new Label { Text = source == null ? "选择表格后可查看原有边距。" : "当前表格：" + source.Description(),
                Location = new Point(42, 49), Size = new Size(462, 30), ForeColor = SystemColors.GrayText });
            values = new Panel { Location = new Point(38, 118), Size = new Size(464, 130) }; Controls.Add(values);
            top = Number("上", 0, 0, settings.TopMm); bottom = Number("下", 230, 0, settings.BottomMm);
            left = Number("左", 0, 42, settings.LeftMm); right = Number("右", 230, 42, settings.RightMm);
            var five = new Button { Text = "全部 5 毫米", Location = new Point(0, 86), Size = new Size(132, 32) };
            var ten = new Button { Text = "全部 10 毫米", Location = new Point(144, 86), Size = new Size(132, 32) };
            var read = new Button { Text = "读取表格边距", Location = new Point(288, 86), Size = new Size(156, 32), Enabled = source != null };
            values.Controls.Add(five); values.Controls.Add(ten); values.Controls.Add(read);
            five.Click += delegate { top.Value = bottom.Value = left.Value = right.Value = 5; };
            ten.Click += delegate { top.Value = bottom.Value = left.Value = right.Value = 10; };
            read.Click += delegate {
                if (source == null) return;
                try { source.Validate(); SetValues(source); }
                catch (Exception ex) { MessageBox.Show(this, ex.Message, "无法读取", MessageBoxButtons.OK, MessageBoxIcon.Information); }
            };
            Controls.Add(new Label { Text = "设置会自动记住，表格保持 100% 原尺寸。\r\n自定义时从左上边距放置；纸型较大时仍可能有额外留白。",
                Location = new Point(20, 253), Size = new Size(486, 45) });
            var ok = new Button { Text = "确定", Location = new Point(290, 310), Size = new Size(96, 34) };
            var cancel = new Button { Text = "取消", Location = new Point(398, 310), Size = new Size(96, 34), DialogResult = DialogResult.Cancel };
            Controls.Add(ok); Controls.Add(cancel); AcceptButton = ok; CancelButton = cancel;
            custom.CheckedChanged += delegate { values.Enabled = custom.Checked; };
            original.Checked = !settings.UseCustom; custom.Checked = settings.UseCustom; values.Enabled = custom.Checked;
            ok.Click += delegate {
                Result = new PageMargins { UseCustom = custom.Checked, TopMm = (double)top.Value, BottomMm = (double)bottom.Value,
                    LeftMm = (double)left.Value, RightMm = (double)right.Value };
                Result.Validate(); DialogResult = DialogResult.OK;
            };
        }

        NumericUpDown Number(string label, int x, int y, double value)
        {
            values.Controls.Add(new Label { Text = label, Location = new Point(x, y + 4), Size = new Size(30, 28) });
            var input = new NumericUpDown { Location = new Point(x + 36, y), Size = new Size(174, 28),
                Minimum = 0, Maximum = 100, DecimalPlaces = 1, Increment = 0.5M, Value = (decimal)value };
            values.Controls.Add(input); return input;
        }

        void SetValues(PageMargins data)
        { top.Value = (decimal)data.TopMm; bottom.Value = (decimal)data.BottomMm; left.Value = (decimal)data.LeftMm; right.Value = (decimal)data.RightMm; }
    }
}
