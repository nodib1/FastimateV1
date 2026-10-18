using System.Drawing;
using System.Windows.Forms;

namespace FASTIMATE
{
    /// <summary>
    /// Mali prozor "SolidWorks - Proces u toku" koji stoji iznad SolidWorksa dok traje obrada
    /// (isti izgled za Ponudu, Tehnologiju i Izvoz dokumentacije).
    /// </summary>
    public static class ProcesUTokuProzor
    {
        public static Form Prikazi()
        {
            try
            {
                Form box = new Form
                {
                    Text = "SolidWorks - Proces u toku",
                    Width = 420,
                    Height = 130,
                    FormBorderStyle = FormBorderStyle.FixedDialog,
                    MaximizeBox = false,
                    MinimizeBox = false,
                    ControlBox = false,
                    StartPosition = FormStartPosition.CenterScreen,
                    TopMost = true,
                    BackColor = Color.White
                };

                box.Controls.Add(new Label
                {
                    Text = "Proces u toku !",
                    ForeColor = Color.FromArgb(2, 136, 209),
                    Font = new Font("Segoe UI", 16, FontStyle.Bold),
                    Dock = DockStyle.Fill,
                    TextAlign = ContentAlignment.MiddleCenter
                });

                box.Show();
                Application.DoEvents();
                return box;
            }
            catch
            {
                return null;
            }
        }

        public static void Zatvori(Form box)
        {
            try
            {
                if (box != null && !box.IsDisposed)
                {
                    box.Close();
                    box.Dispose();
                }
            }
            catch { }
        }
    }
}
