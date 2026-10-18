using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Input;
using SolidWorks.Interop.sldworks;

namespace FASTIMATE
{
    public partial class IzvozDokumentacijeWindow : Window
    {
        public IzvozDokumentacijeWindow(string initialPath = "")
        {
            InitializeComponent();
            txtPutanja.Text = "";
            UpdatePlaceholderVisibility();
        }

        private void txtPutanja_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            UpdatePlaceholderVisibility();
        }

        private void UpdatePlaceholderVisibility()
        {
            if (lblPutanjaPlaceholder != null)
            {
                lblPutanjaPlaceholder.Visibility = string.IsNullOrWhiteSpace(txtPutanja.Text) ? Visibility.Visible : Visibility.Collapsed;
            }
        }

        private void btnOdabrati_Click(object sender, RoutedEventArgs e)
        {
            using (FolderBrowserDialog dialog = new FolderBrowserDialog())
            {
                dialog.Description = "Odaberite pripremni (root) direktorij projekta";
                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    txtPutanja.Text = dialog.SelectedPath;
                }
            }
        }

        private void chkSve_Checked(object sender, RoutedEventArgs e)
        {
            if (chkPdf != null) chkPdf.IsChecked = true;
            if (chkDxf != null) chkDxf.IsChecked = true;
            if (chkStep != null) chkStep.IsChecked = true;
        }

        private void chkSve_Unchecked(object sender, RoutedEventArgs e)
        {
            if (chkPdf != null) chkPdf.IsChecked = false;
            if (chkDxf != null) chkDxf.IsChecked = false;
            if (chkStep != null) chkStep.IsChecked = false;
        }

        private async void btnStartExport_Click(object sender, RoutedEventArgs e)
        {
            string rootPath = txtPutanja.Text.Trim();
            if (string.IsNullOrWhiteSpace(rootPath) || !Directory.Exists(rootPath))
            {
                System.Windows.MessageBox.Show("Molimo odaberite ispravan pripremni direktorij projekta.", "Upozorenje", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            bool doPdf = chkPdf.IsChecked == true;
            bool doDxf = chkDxf.IsChecked == true;
            bool doStep = chkStep.IsChecked == true;

            if (!doPdf && !doDxf && !doStep)
            {
                System.Windows.MessageBox.Show("Molimo odaberite bar jedan tip dokumenta za izvoz (PDF, DXF ili STEP).", "Upozorenje", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            SldWorks swApp = null;
            try
            {
                swApp = (SldWorks)Marshal.GetActiveObject("SldWorks.Application");
            }
            catch
            {
                System.Windows.MessageBox.Show("SolidWorks aplikacija nije pokrenuta.", "Greška", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (swApp == null || swApp.ActiveDoc == null)
            {
                System.Windows.MessageBox.Show("Nema otvorenog dokumenta u SolidWorks-u.", "Greška", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            loadingOverlay.Visibility = Visibility.Visible;
            btnStartExport.IsEnabled = false;
            System.Windows.Input.Mouse.OverrideCursor = System.Windows.Input.Cursors.Wait;

            await Task.Delay(100);

            try
            {
                DokumentacijaExporter.RunExport(swApp, swApp.ActiveDoc as ModelDoc2, rootPath, doPdf, doDxf, doStep);
            }
            finally
            {
                loadingOverlay.Visibility = Visibility.Collapsed;
                btnStartExport.IsEnabled = true;
                System.Windows.Input.Mouse.OverrideCursor = null;
            }
        }
    }
}
