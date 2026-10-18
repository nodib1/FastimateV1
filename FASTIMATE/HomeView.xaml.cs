using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms;  // Potrebno za FolderBrowserDialog
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Runtime.InteropServices;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Windows.Shell;

namespace FASTIMATE
{
    /// <summary>
    /// Interaction logic for HomeView.xaml
    /// </summary>
    public partial class HomeView : System.Windows.Controls.UserControl
    {

        public enum Modus
        {
            Nedefinisano,
            Ponuda,
            Tehnologija
        }

        public Modus nacinRada;

        public string sLogovaniKorisnik;

        public HomeView()
        {
            InitializeComponent();
            OcistiStareMakroeTemp();
        }

        private void btnOdabrati_Click(object sender, RoutedEventArgs e)
        {
            using (var dialog = new FolderBrowserDialog())
            {
                dialog.Description = "Izaberi folder za snimanje fajla";
                dialog.ShowNewFolderButton = true;

                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    txtPutanja.Text = dialog.SelectedPath;
                }
            }
        }

        private void txtIdent_PreviewTextInput(object sender, System.Windows.Input.TextCompositionEventArgs e)
        {
            // Provera da li je uneti tekst broj
            e.Handled = !IsNumeric(e.Text);
        }

        // Metoda koja proverava da li je tekst broj
        private bool IsNumeric(string text)
        {
            int result;
            return int.TryParse(text, out  result);
        }

        private void OcistiStareMakroeTemp()
        {
            try
            {
                string tempPath = System.IO.Path.GetTempPath();
                var files = Directory.GetFiles(tempPath, "mc_*.swp");

                foreach (var file in files)
                {
                    try
                    {
                        File.Delete(file);
                    }
                    catch
                    {
                        // Ako fajl još drži neki proces, preskoči
                    }
                }
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"Greška pri čišćenju temp fajlova: {ex.Message}");
            }
        }


        private async void btnStart_Click(object sender, RoutedEventArgs e)
        {
            if (!(txtIdent.Text.Length == 6 && IsNumeric(txtIdent.Text)))
            {
                System.Windows.MessageBox.Show("Ident mora sadržavati tačno 6 cifara.", "Upozorenje", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (nacinRada == Modus.Nedefinisano)
            {
                System.Windows.MessageBox.Show("Potrebno je odabrati način rada: Ponuda ili Tehnologija!", "Upozorenje", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // Prikaz indikatora obrade u WPF aplikaciji
            loadingOverlay.Visibility = Visibility.Visible;
            btnStart.IsEnabled = false;
            System.Windows.Input.Mouse.OverrideCursor = System.Windows.Input.Cursors.Wait;

            await Task.Delay(100);

            try
            {
                if (nacinRada == Modus.Ponuda)
                {
                    Pokrenuti(1);
                }
                else if (nacinRada == Modus.Tehnologija)
                {
                    Pokrenuti(2);
                }
            }
            finally
            {
                loadingOverlay.Visibility = Visibility.Collapsed;
                btnStart.IsEnabled = true;
                System.Windows.Input.Mouse.OverrideCursor = null;
            }
        }


        public void Pokrenuti(int tip)
        {
            try
            {
                // Pristup aktivnoj SolidWorks aplikaciji
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

                if (swApp == null)
                {
                    System.Windows.MessageBox.Show("SolidWorks aplikacija nije pokrenuta.", "Greška", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                ModelDoc2 swModel = swApp.ActiveDoc;
                if (swModel == null)
                {
                    System.Windows.MessageBox.Show("Nema aktivnog dokumenta u SolidWorks-u.", "Greška", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                // 1. Višeslojno rješavanje IDENT i PUTANJA
                string ident = txtIdent.Text;
                string putanja = txtPutanja.Text;

                if (string.IsNullOrWhiteSpace(ident))
                    ident = TehnologijaProcessor.GetModelProperty(swModel, "IDENT");

                if (string.IsNullOrWhiteSpace(putanja))
                    putanja = TehnologijaProcessor.GetModelProperty(swModel, "PUTANJA");

                if (string.IsNullOrWhiteSpace(putanja) && swModel != null)
                {
                    string mPath = swModel.GetPathName();
                    if (!string.IsNullOrWhiteSpace(mPath))
                        putanja = System.IO.Path.GetDirectoryName(mPath);
                }

                if (string.IsNullOrWhiteSpace(ident) && swModel != null)
                {
                    string mPath = swModel.GetPathName();
                    if (!string.IsNullOrWhiteSpace(mPath))
                        ident = System.IO.Path.GetFileNameWithoutExtension(mPath);
                    else
                        ident = swModel.GetTitle();
                }

                ident = ident?.Trim(' ', '"', '\'');
                putanja = putanja?.Trim(' ', '"', '\'');

                if (string.IsNullOrWhiteSpace(ident) || string.IsNullOrWhiteSpace(putanja))
                {
                    System.Windows.MessageBox.Show("Nisu pronađeni IDENT ili PUTANJA. Molimo unesite IDENT ili sačuvajte model u SolidWorks-u.", "Greška", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                // Ažuriranje UI tekstualnih polja
                txtIdent.Text = ident;
                txtPutanja.Text = putanja;

                // Upis custom property-ja u SolidWorks model (Global i Konfiguracija)
                CustomPropertyManager custPropMgr = swModel.Extension.get_CustomPropertyManager("");
                custPropMgr.Add3("IDENT", 30, ident, 2);
                custPropMgr.Add3("PUTANJA", 30, putanja, 2);
                custPropMgr.Add3("Logovani_korisnik", 30, sLogovaniKorisnik, 2);
                custPropMgr.Add3("DATUMUpisa", 30, DateTime.Now.ToString("dd.MM.yyyy"), 2);

                try
                {
                    string activeConf = swModel.ConfigurationManager?.ActiveConfiguration?.Name;
                    if (!string.IsNullOrEmpty(activeConf))
                    {
                        CustomPropertyManager confProp = swModel.Extension.get_CustomPropertyManager(activeConf);
                        confProp?.Add3("IDENT", 30, ident, 2);
                        confProp?.Add3("PUTANJA", 30, putanja, 2);
                    }
                }
                catch { }

                // 2. Obrada po tipu (Ponuda vs Tehnologija)
                if (tip == 1)
                {
                    // PONUDA: Kreiranje Excel fajla za Ponudu i pokretanje direktnog C# obrade
                    string ponudaExcelPath = System.IO.Path.Combine(putanja, $"PRORACUN_VREMENA_PONUDE_{ident}.xlsx");

                    if (!File.Exists(ponudaExcelPath))
                    {
                        byte[] excelBytes = TemplateManager.GetActiveTemplateBytes(1);
                        if (excelBytes != null && excelBytes.Length > 0)
                        {
                            try
                            {
                                Directory.CreateDirectory(putanja);
                                File.WriteAllBytes(ponudaExcelPath, excelBytes);
                            }
                            catch (Exception ex)
                            {
                                System.Windows.MessageBox.Show($"Ne mogu kreirati Excel fajl na lokaciji:\n{ponudaExcelPath}\nGreška: {ex.Message}", "Greška", MessageBoxButton.OK, MessageBoxImage.Error);
                                return;
                            }
                        }
                    }

                    PonudaProcessor.RunProcess(swApp, swModel, sLogovaniKorisnik, ident, putanja);
                }
                else if (tip == 2)
                {
                    // TEHNOLOGIJA: Kreiranje Excel fajla i pokretanje C# obrade (bez makroa)
                    string targetExcelPath = System.IO.Path.Combine(putanja, $"PRORACUN_VREMENA_TEHNOLOGIJE_{ident}.xlsx");

                    if (!File.Exists(targetExcelPath))
                    {
                        byte[] excelBytes = TemplateManager.GetActiveTemplateBytes(2);
                        if (excelBytes != null && excelBytes.Length > 0)
                        {
                            try
                            {
                                Directory.CreateDirectory(putanja);
                                File.WriteAllBytes(targetExcelPath, excelBytes);
                            }
                            catch (Exception ex)
                            {
                                System.Windows.MessageBox.Show($"Ne mogu kreirati Excel fajl na lokaciji:\n{targetExcelPath}\nGreška: {ex.Message}", "Greška", MessageBoxButton.OK, MessageBoxImage.Error);
                                return;
                            }
                        }
                    }

                    TehnologijaProcessor.RunProcess(swApp, swModel, targetExcelPath, ident, sLogovaniKorisnik, DateTime.Now.ToString("dd.MM.yyyy"));
                }
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"Greška pri pokretanju: {ex.Message}", "Greška", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void btnPonuda_Click(object sender, RoutedEventArgs e)
        {
            nacinRada = Modus.Ponuda;

            // Aktivno dugme sivo, neaktivno vrati na originalnu boju
            btnPonuda.Background = Brushes.Gray;
            btnTehnologija.ClearValue(System.Windows.Controls.Button.BackgroundProperty);
        }

        private void btnTehnologija_Click(object sender, RoutedEventArgs e)
        {
            nacinRada = Modus.Tehnologija;

            btnTehnologija.Background = Brushes.Gray;
            btnPonuda.ClearValue(System.Windows.Controls.Button.BackgroundProperty);
        }

        private void btnOsvjezi_Click(object sender, RoutedEventArgs e)
        {
            // Ispraznite oba TextBox-a
            txtIdent.Clear();
            txtPutanja.Clear();
            btnPonuda.ClearValue(System.Windows.Controls.Button.BackgroundProperty);
            btnTehnologija.ClearValue(System.Windows.Controls.Button.BackgroundProperty);
            nacinRada = Modus.Nedefinisano;
        }
    }
}
