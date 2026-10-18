using System;
using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms;
using System.Windows.Media;

namespace FASTIMATE
{
    public partial class UvozSablonView : System.Windows.Controls.UserControl
    {
        public UvozSablonView()
        {
            InitializeComponent();
            RefreshStatus();
        }

        private void RefreshStatus()
        {
            try
            {
                if (TemplateManager.HasCustomTemplate())
                {
                    lblStatusText.Text = "Prilagođeni šablon uvezen (Korisnički)";
                    lblStatusText.Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FFB74D"));
                    lblInfo.Text = "U upotrebi je vaš uvezeni Excel šablon. Svi novi obračuni će automatski koristiti ovaj fajl.";
                }
                else
                {
                    lblStatusText.Text = "Fabrički ugrađeni šablon";
                    lblStatusText.Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#81C784"));
                    lblInfo.Text = "U upotrebi je podrazumijevani fabrički Excel šablon ugrađen u aplikaciju.";
                }

                int count = FastenerDbManager.GetFastenerCount();
                if (count > 0)
                {
                    lblErpCount.Text = $"{count} identa u bazi";
                    lblErpCount.Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#4FC3F7"));
                }
                else
                {
                    lblErpCount.Text = "Tabela nije uvezena";
                    lblErpCount.Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#AAAAAA"));
                }

                int smCount = SheetMetalDbManager.GetItemCount();
                if (smCount > 0)
                {
                    lblSmCount.Text = $"{smCount} identa u bazi";
                    lblSmCount.Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#A5D6A7"));
                }
                else
                {
                    lblSmCount.Text = "Tabela nije uvezena";
                    lblSmCount.Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#AAAAAA"));
                }

                lvErpFasteners.ItemsSource = FastenerDbManager.GetAllItems();
                lvSheetMetal.ItemsSource = SheetMetalDbManager.GetAllItems();
                lvKFactors.ItemsSource = KFactorManager.GetAllRules();
            }
            catch { }
        }

        private void btnUvezi_Click(object sender, RoutedEventArgs e)
        {
            if (TemplateManager.ImportNewTemplate())
            {
                RefreshStatus();
            }
        }

        private void btnVratiFabricki_Click(object sender, RoutedEventArgs e)
        {
            TemplateManager.ResetToDefaultTemplate();
            RefreshStatus();
        }

        private void btnUveziErp_Click(object sender, RoutedEventArgs e)
        {
            using (var dialog = new OpenFileDialog())
            {
                dialog.Filter = "Tabela vijčane robe (*.xlsx;*.xls;*.csv;*.txt)|*.xlsx;*.xls;*.csv;*.txt|Excel fajlovi (*.xlsx;*.xls)|*.xlsx;*.xls|Tekstualni fajlovi (*.csv;*.txt)|*.csv;*.txt|Svi fajlovi (*.*)|*.*";
                dialog.Title = "Odaberite tabelu vijčane robe iz ERP sistema";

                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    if (FastenerDbManager.ImportFile(dialog.FileName))
                    {
                        System.Windows.MessageBox.Show($"Tabela vijčane robe iz ERP sistema je uspješno uvezena!\nUčitano je {FastenerDbManager.GetFastenerCount()} identa.", "Uspjeh", MessageBoxButton.OK, MessageBoxImage.Information);
                        RefreshStatus();
                    }
                    else
                    {
                        System.Windows.MessageBox.Show("Greška pri uvozu ERP tabele vijčane robe. Provjerite da li fajl sadrži kolone sa nazivima i identima.", "Greška", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                }
            }
        }

        private void btnSacuvajErp_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                FastenerDbManager.SaveData();
                RefreshStatus();
                System.Windows.MessageBox.Show("Izmjene u bazi vijčane robe su uspješno sačuvane!", "Uspjeh", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"Greška pri čuvanju izmjena: {ex.Message}", "Greška", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void btnObrisiErp_Click(object sender, RoutedEventArgs e)
        {
            if (FastenerDbManager.GetFastenerCount() == 0)
            {
                System.Windows.MessageBox.Show("Baza vijčane robe je već prazna.", "Informacija", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            if (System.Windows.MessageBox.Show("Da li ste sigurni da želite obrisati uvezanu bazu vijčane robe?", "Potvrda brisanja", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
            {
                FastenerDbManager.ClearData();
                RefreshStatus();
                System.Windows.MessageBox.Show("Baza vijčane robe je uspješno obrisana.", "Uspjeh", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        private void btnUveziSheetMetal_Click(object sender, RoutedEventArgs e)
        {
            using (var dialog = new OpenFileDialog())
            {
                dialog.Filter = "Tabela sirovina limova (*.xlsx;*.xls;*.csv;*.txt)|*.xlsx;*.xls;*.csv;*.txt|Excel fajlovi (*.xlsx;*.xls)|*.xlsx;*.xls|Tekstualni fajlovi (*.csv;*.txt)|*.csv;*.txt|Svi fajlovi (*.*)|*.*";
                dialog.Title = "Odaberite tabelu sirovina limova iz ERP sistema";

                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    if (SheetMetalDbManager.ImportFile(dialog.FileName))
                    {
                        System.Windows.MessageBox.Show($"Tabela sirovina limova iz ERP sistema je uspješno uvezena!\nUčitano je {SheetMetalDbManager.GetItemCount()} identa.", "Uspjeh", MessageBoxButton.OK, MessageBoxImage.Information);
                        RefreshStatus();
                    }
                    else
                    {
                        System.Windows.MessageBox.Show("Greška pri uvozu ERP tabele sirovina limova. Provjerite da li fajl sadrži kolone sa identima i materijalom.", "Greška", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                }
            }
        }

        private void btnSacuvajSheetMetal_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                SheetMetalDbManager.SaveData();
                RefreshStatus();
                System.Windows.MessageBox.Show("Izmjene u bazi sirovina limova su uspješno sačuvane!", "Uspjeh", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"Greška pri čuvanju sirovina limova: {ex.Message}", "Greška", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void btnObrisiSheetMetal_Click(object sender, RoutedEventArgs e)
        {
            if (SheetMetalDbManager.GetItemCount() == 0)
            {
                System.Windows.MessageBox.Show("Baza sirovina limova je već prazna.", "Informacija", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            if (System.Windows.MessageBox.Show("Da li ste sigurni da želite obrisati bazu sirovina limova?", "Potvrda brisanja", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
            {
                SheetMetalDbManager.ClearData();
                RefreshStatus();
                System.Windows.MessageBox.Show("Baza sirovina limova je uspješno obrisana.", "Uspjeh", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        private void txtNewSmIdent_PreviewTextInput(object sender, System.Windows.Input.TextCompositionEventArgs e)
        {
            // Dozvoljava ISKLJUČIVO brojčane cifre (0-9) - blokira unos bilo kojeg teksta ili simbola
            e.Handled = !System.Text.RegularExpressions.Regex.IsMatch(e.Text, @"^[0-9]+$");
        }

        private void txtNewSmThickness_PreviewTextInput(object sender, System.Windows.Input.TextCompositionEventArgs e)
        {
            // Dozvoljava samo brojčane cifre, tačku i zarez za debljinu (npr. 2.5 ili 2,5)
            e.Handled = !System.Text.RegularExpressions.Regex.IsMatch(e.Text, @"^[0-9\,\.]+$");
        }

        private void btnDodajSheetMetal_Click(object sender, RoutedEventArgs e)
        {
            string id = txtNewSmIdent.Text.Trim();
            string mat = cmbNewSmMaterial.Text.Trim();
            if (string.IsNullOrWhiteSpace(mat)) mat = (cmbNewSmMaterial.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? "Crni lim";
            string strTh = txtNewSmThickness.Text.Trim().Replace(',', '.');
            string desc = txtNewSmDesc.Text.Trim();

            if (string.IsNullOrWhiteSpace(id) || id.Length != 6 || !System.Text.RegularExpressions.Regex.IsMatch(id, @"^\d{6}$"))
            {
                System.Windows.MessageBox.Show("ERP Ident sirovine mora sadržavati ISKLJUČIVO BROJEVE i imati TAČNO 6 CIFARA (npr. 900123).", "Neispravan ERP Ident", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (!double.TryParse(strTh, NumberStyles.Any, CultureInfo.InvariantCulture, out double th) || th <= 0)
            {
                System.Windows.MessageBox.Show("Molimo unesite ispravnu pozitivnu debljinu u mm (npr. 2.0).", "Neispravan Unos", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (!SheetMetalDbManager.AddNewItem(id, mat, th, desc, out string error))
            {
                System.Windows.MessageBox.Show(error, "Dodavanje Nije Dozvoljeno", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            txtNewSmIdent.Clear();
            txtNewSmThickness.Clear();
            txtNewSmDesc.Clear();
            RefreshStatus();
            System.Windows.MessageBox.Show($"Nova sirovina lima (Ident: {id}, {mat}, {th} mm) je uspješno dodana!", "Uspjeh", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void btnDodajKFactor_Click(object sender, RoutedEventArgs e)
        {
            string material = (cmbNewMaterial.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? "Čelik";
            string strThickness = txtNewThickness.Text.Trim().Replace(',', '.');
            string strKFactor = txtNewKFactor.Text.Trim().Replace(',', '.');

            if (!double.TryParse(strThickness, NumberStyles.Any, CultureInfo.InvariantCulture, out double thickness) || thickness <= 0)
            {
                System.Windows.MessageBox.Show("Molimo unesite ispravnu pozitivnu debljinu (npr. 2.5).", "Neispravan Unos", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (!double.TryParse(strKFactor, NumberStyles.Any, CultureInfo.InvariantCulture, out double kFactor) || kFactor < 0)
            {
                System.Windows.MessageBox.Show("Molimo unesite ispravan K-Faktor (npr. 0.25).", "Neispravan Unos", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (!KFactorManager.AddNewRule(material, thickness, kFactor, out string error))
            {
                System.Windows.MessageBox.Show(error, "Dodavanje Nije Dozvoljeno", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            txtNewThickness.Clear();
            txtNewKFactor.Clear();
            RefreshStatus();
            System.Windows.MessageBox.Show($"Novi K-Faktor ({material}, {thickness} mm => {kFactor}) je uspješno dodan!", "Uspjeh", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void btnSacuvajKFactors_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                KFactorManager.SaveAllRules();
                RefreshStatus();
                System.Windows.MessageBox.Show("Sva pravila K-Faktora su uspješno sačuvana!", "Uspjeh", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"Greška pri čuvanju K-Faktora: {ex.Message}", "Greška", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void btnObrisiKFactor_Click(object sender, RoutedEventArgs e)
        {
            var selected = lvKFactors.SelectedItem as KFactorRuleItem;
            if (selected == null)
            {
                System.Windows.MessageBox.Show("Molimo odaberite K-Faktor iz tabele koji želite obrisati.", "Informacija", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            if (System.Windows.MessageBox.Show($"Da li ste sigurni da želite obrisati K-Faktor za {selected.Material} ({selected.Thickness} mm)?", "Potvrda brisanja", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
            {
                KFactorManager.DeleteRule(selected);
                RefreshStatus();
                System.Windows.MessageBox.Show("K-Faktor je uspješno obrisan.", "Uspjeh", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }
    }
}
