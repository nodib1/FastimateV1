using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace FASTIMATE
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class LOGIN : Window
    {
        public LOGIN()
        {
            InitializeComponent();
        }

        private void btnSnimit_Click(object sender, RoutedEventArgs e)
        {
            string username = txtUsername.Text != null ? txtUsername.Text.Trim() : "";
            string password = txtPasw.Password;

            bool isAdmin = username.Equals("admin", StringComparison.OrdinalIgnoreCase) && (password == "p!ms\"=\"&" || password == "mtp2025" || password == "mtp2026" || password == "admin");
            bool isUser = username.Equals("pims", StringComparison.OrdinalIgnoreCase) && password == "mtp2026";

            if (isAdmin || isUser)
            {
                if (!isAdmin)
                {
                    var licInfo = LicenseManager.CheckLicense();
                    if (!licInfo.IsValid)
                    {
                        lblError.Text = "Licenca za FASTIMATE je istekla ili nije aktivirana!\nObratite se administratoru za aktivaciju.";
                        lblError.Visibility = Visibility.Visible;
                        return;
                    }

                    if (licInfo.Show15DaysWarning)
                    {
                        LicenseWarningWindow warningWin = new LicenseWarningWindow(licInfo.RemainingDays, licInfo.ExpirationDate);
                        warningWin.Owner = this;
                        warningWin.ShowDialog();
                    }
                }

                lblError.Visibility = Visibility.Collapsed;

                string displayUser = isAdmin ? "admin" : "Dino Bratić";
                MainNovi mainWindow = new MainNovi(displayUser, isAdmin);
                mainWindow.Show();

                this.Close();
            }
            else
            {
                lblError.Text = "Pogrešno korisničko ime ili lozinka!";
                lblError.Visibility = Visibility.Visible;
                txtPasw.Clear();
                txtPasw.Focus();
            }
        }

        private void btnOdustati_Click(object sender, RoutedEventArgs e)
        {
            Application.Current.Shutdown();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            txtUsername.Focus();
        }

        private void txtUsername_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                txtPasw.Focus();
            }
        }

        private void txtPasw_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                btnSnimit.Focus(); // Fokus na dugme
            }
        }
        private void btnSnimit_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                btnSnimit_Click(sender, e); // Simulira klik
            }
        }
    }
}
