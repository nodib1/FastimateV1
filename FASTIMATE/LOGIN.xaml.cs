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

            UserManager.LoginError loginError;
            AppUser user = UserManager.Authenticate(username, password, out loginError);

            if (user != null)
            {
                bool isAdmin = user.IsAdmin;

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

                MainNovi mainWindow = new MainNovi(user.PunoIme, isAdmin, user.Username);
                mainWindow.Show();

                this.Close();
            }
            else
            {
                lblError.Text = UserManager.LoginErrorMessage(loginError, username);
                lblError.Visibility = Visibility.Visible;

                if (loginError == UserManager.LoginError.PraznoKorisnickoIme
                    || loginError == UserManager.LoginError.NepoznatKorisnik
                    || loginError == UserManager.LoginError.Deaktiviran)
                {
                    txtPasw.Clear();
                    txtUsername.Focus();
                    txtUsername.SelectAll();
                }
                else
                {
                    txtPasw.Clear();
                    txtPasw.Focus();
                }
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
