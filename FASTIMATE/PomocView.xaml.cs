using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Windows;
using System.Windows.Controls;

namespace FASTIMATE
{
    public partial class PomocView : UserControl
    {
        public PomocView()
        {
            InitializeComponent();
        }

        public static void OpenHelpPdfSafely()
        {
            try
            {
                string tempDir = Path.Combine(Path.GetTempPath(), "FASTIMATE");
                if (!Directory.Exists(tempDir)) Directory.CreateDirectory(tempDir);

                string tempPdf = Path.Combine(tempDir, "Pomoc_FASTIMATE_Uputstvo.pdf");
                bool extracted = false;

                // 1. UVIJEK PRVO PROBAJ IZVUĆI IZ SAMOG FASTIMATE.EXE (Embedded Resource)
                try
                {
                    Assembly asm = Assembly.GetExecutingAssembly();
                    using (Stream stream = asm.GetManifestResourceStream("FASTIMATE.Pomoc_FASTIMATE.pdf"))
                    {
                        if (stream != null)
                        {
                            using (FileStream fs = new FileStream(tempPdf, FileMode.Create, FileAccess.Write))
                            {
                                stream.CopyTo(fs);
                            }
                            extracted = true;
                        }
                    }
                }
                catch { }

                // 2. FALLBACK AKO IZ BILO KOG RAZLOGA RESURS NIJE PROČITAN -> PROVJERI VANJSKE LOKACIJE
                if (!extracted || !File.Exists(tempPdf))
                {
                    string baseDir = AppDomain.CurrentDomain.BaseDirectory;
                    string p1 = Path.Combine(baseDir, "Pomoc_FASTIMATE.pdf");
                    string p2 = Path.Combine(baseDir, "fastimate.pdf");
                    string p3 = Path.Combine(AppPaths.GetDbDirectory(), "Pomoc_FASTIMATE.pdf");

                    string sourcePdf = File.Exists(p1) ? p1 : (File.Exists(p2) ? p2 : (File.Exists(p3) ? p3 : null));

                    if (sourcePdf != null && File.Exists(sourcePdf))
                    {
                        try
                        {
                            File.Copy(sourcePdf, tempPdf, true);
                            extracted = true;
                        }
                        catch
                        {
                            tempPdf = Path.Combine(tempDir, $"Pomoc_FASTIMATE_{DateTime.Now:HHmmss}.pdf");
                            File.Copy(sourcePdf, tempPdf, true);
                            extracted = true;
                        }
                    }
                }

                if (File.Exists(tempPdf))
                {
                    ProcessStartInfo psi = new ProcessStartInfo
                    {
                        FileName = tempPdf,
                        UseShellExecute = true
                    };
                    Process.Start(psi);
                }
                else
                {
                    MessageBox.Show("PDF fajl pomoći nije pronađen.", "Upozorenje", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Greška pri otvaranju PDF fajla: {ex.Message}", "Greška", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void btnOtveriPdfVani_Click(object sender, RoutedEventArgs e)
        {
            OpenHelpPdfSafely();
        }
    }
}
