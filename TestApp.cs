using System;
using System.Windows.Forms;

namespace FitsPreviewHandler
{
    static class Program
    {
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            
            var form = new Form
            {
                Text = "FITS Preview Test",
                Width = 800,
                Height = 600
            };

            var control = new FitsPreviewControl();
            control.Dock = DockStyle.Fill;
            form.Controls.Add(control);

            form.AllowDrop = true;
            form.DragEnter += (s, e) => {
                if (e.Data.GetDataPresent(DataFormats.FileDrop)) e.Effect = DragDropEffects.Copy;
            };
            form.DragDrop += (s, e) => {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                if (files.Length > 0) control.LoadFits(files[0]);
            };

            Application.Run(form);
        }
    }
}
