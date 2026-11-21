using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace AutoPaper
{
    internal sealed class ExportCompleteDialog : Form
    {
        [DllImport("shell32.dll", CharSet = CharSet.Unicode, ExactSpelling = true)]
        static extern int SHParseDisplayName(string path, IntPtr bindContext, out IntPtr item,
            uint attributes, IntPtr attributesOut);

        [DllImport("shell32.dll", ExactSpelling = true)]
        static extern int SHOpenFolderAndSelectItems(IntPtr item, uint count, IntPtr children, uint flags);

        public ExportCompleteDialog(string destination, Font font)
        {
            Text = "导出完成";
            Font = font;
            AutoScaleMode = AutoScaleMode.Dpi;
            ClientSize = new Size(500, 158);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.CenterParent;

            Controls.Add(new Label { Text = "PDF 已导出，原表格未改写。", AutoSize = true, Location = new Point(16, 14) });
            Controls.Add(new TextBox { Text = destination, ReadOnly = true, Multiline = true,
                ScrollBars = ScrollBars.Vertical, Location = new Point(16, 44), Size = new Size(468, 52) });
            var close = new Button { Text = "关闭", DialogResult = DialogResult.Cancel,
                Location = new Point(396, 110), Size = new Size(88, 34), TabIndex = 0 };
            var reveal = new Button { Text = "打开文件所在位置", DialogResult = DialogResult.OK,
                Location = new Point(222, 110), Size = new Size(162, 34), TabIndex = 1 };
            Controls.Add(reveal);
            Controls.Add(close);
            AcceptButton = close;
            CancelButton = close;
            ActiveControl = close;
        }

        public static void ShowResult(Form owner, string destination)
        {
            using (var dialog = new ExportCompleteDialog(destination, owner.Font))
                if (dialog.ShowDialog(owner) != DialogResult.OK) return;

            IntPtr item = IntPtr.Zero;
            try
            {
                Marshal.ThrowExceptionForHR(SHParseDisplayName(destination, IntPtr.Zero, out item, 0, IntPtr.Zero));
                // A single absolute PIDL with zero children selects the file in its parent folder.
                Marshal.ThrowExceptionForHR(SHOpenFolderAndSelectItems(item, 0, IntPtr.Zero, 0));
            }
            catch (Exception ex)
            {
                MessageBox.Show(owner, destination + Environment.NewLine + ex.Message, "无法打开文件所在位置",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            finally { Marshal.FreeCoTaskMem(item); }
        }
    }
}
