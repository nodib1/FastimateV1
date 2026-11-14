using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace FASTIMATE
{
    public partial class KorisniciView : UserControl
    {
        private readonly string _logovaniUsername;
        private string _editUsername; // null = unos novog korisnika

        public KorisniciView(string logovaniUsername = "")
        {
            InitializeComponent();
            _logovaniUsername = logovaniUsername ?? "";
        }

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            RefreshList();
            ClearForm();
        }

        private void RefreshList(string selectUsername = null)
        {
            bool signatureValid;
            var users = UserManager.LoadUsers(out signatureValid).OrderBy(u => u.Username, StringComparer.OrdinalIgnoreCase).ToList();
            gridKorisnici.ItemsSource = users;
            bannerMijenjan.Visibility = signatureValid ? Visibility.Collapsed : Visibility.Visible;

            if (selectUsername != null)
            {
                gridKorisnici.SelectedItem = users.FirstOrDefault(u => u.Username.Equals(selectUsername, StringComparison.OrdinalIgnoreCase));
            }
        }

        private void ClearForm()
        {
            _editUsername = null;
            gridKorisnici.SelectedItem = null;

            lblFormaNaslov.Text = "NOVI KORISNIK";
            lblLozinka.Text = "Lozinka:";
            txtUsername.Text = "";
            txtIme.Text = "";
            txtPrezime.Text = "";
            txtLozinka.Clear();
            txtLozinka2.Clear();
            chkAdmin.IsChecked = false;
            chkAktivan.IsChecked = true;
            btnObrisi.IsEnabled = false;
            txtUsername.Focus();
        }

        private void gridKorisnici_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var user = gridKorisnici.SelectedItem as AppUser;
            if (user == null) return;

            _editUsername = user.Username;
            lblFormaNaslov.Text = "IZMJENA: " + user.Username;
            lblLozinka.Text = "Nova lozinka (prazno = ostaje):";
            txtUsername.Text = user.Username;
            txtIme.Text = user.Ime;
            txtPrezime.Text = user.Prezime;
            txtLozinka.Clear();
            txtLozinka2.Clear();
            chkAdmin.IsChecked = user.IsAdmin;
            chkAktivan.IsChecked = user.Aktivan;
            btnObrisi.IsEnabled = true;
            ShowMessage("", true);
        }

        private void btnNovi_Click(object sender, RoutedEventArgs e)
        {
            ClearForm();
            ShowMessage("", true);
        }

        private void btnSacuvaj_Click(object sender, RoutedEventArgs e)
        {
            if (txtLozinka.Password != txtLozinka2.Password)
            {
                ShowMessage("Lozinke se ne podudaraju.", false);
                return;
            }

            var data = new AppUser
            {
                Username = txtUsername.Text,
                Ime = txtIme.Text,
                Prezime = txtPrezime.Text,
                IsAdmin = chkAdmin.IsChecked == true,
                Aktivan = chkAktivan.IsChecked == true
            };

            try
            {
                UserManager.SaveUser(_editUsername, data, txtLozinka.Password);
                string saved = data.Username.Trim();
                RefreshList(saved);
                if (gridKorisnici.SelectedItem == null) ClearForm();
                ShowMessage("Korisnik '" + saved + "' je sačuvan.", true);
            }
            catch (ArgumentException ex)
            {
                ShowMessage(ex.Message, false);
            }
            catch (Exception ex)
            {
                ShowMessage("Greška pri snimanju: " + ex.Message, false);
            }
        }

        private void btnObrisi_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(_editUsername)) return;

            if (_editUsername.Equals(_logovaniUsername, StringComparison.OrdinalIgnoreCase))
            {
                ShowMessage("Ne možete obrisati korisnika s kojim ste trenutno prijavljeni.", false);
                return;
            }

            var res = MessageBox.Show("Obrisati korisnika '" + _editUsername + "'?", "Brisanje korisnika",
                MessageBoxButton.YesNo, MessageBoxImage.Warning);
            if (res != MessageBoxResult.Yes) return;

            try
            {
                string deleted = _editUsername;
                UserManager.DeleteUser(deleted);
                RefreshList();
                ClearForm();
                ShowMessage("Korisnik '" + deleted + "' je obrisan.", true);
            }
            catch (Exception ex)
            {
                ShowMessage("Greška pri brisanju: " + ex.Message, false);
            }
        }

        private void ShowMessage(string text, bool ok)
        {
            lblMessage.Text = text;
            lblMessage.Foreground = (Brush)new BrushConverter().ConvertFromString(ok ? "#81C784" : "#FF5252");
        }
    }
}
