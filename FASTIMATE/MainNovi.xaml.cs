using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
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
    public partial class MainNovi : Window
    {
        public bool IsAdmin { get; private set; } = false;

        /// <summary>Ime i prezime logovanog korisnika (ide u Excel/SolidWorks).</summary>
        public string PunoIme { get; private set; }

        /// <summary>Korisničko ime logovanog korisnika.</summary>
        public string LoginUsername { get; private set; }

        public MainNovi(string punoIme = "admin", bool isAdmin = true, string loginUsername = "admin")
        {
            InitializeComponent();
            IsAdmin = isAdmin;
            PunoIme = punoIme;
            LoginUsername = loginUsername;

            Visibility adminOnly = isAdmin ? Visibility.Visible : Visibility.Collapsed;
            labelkorisnik.Text = isAdmin ? punoIme + " (ADMIN PANEL)" : punoIme;
            if (btnUvozSablon != null) btnUvozSablon.Visibility = adminOnly;
            if (btnLicenca != null) btnLicenca.Visibility = adminOnly;
            if (btnKorisnici != null) btnKorisnici.Visibility = adminOnly;

            LoadHomeView();
        }

        private void Window_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left)
                this.DragMove();
        }

        private void SetActiveSidebarButton(Button activeButton)
        {
            btnHome.Style = (Style)FindResource("SidebarButtonStyle");
            btnIzvozDokumentacije.Style = (Style)FindResource("SidebarButtonStyle");
            btnUvozSablon.Style = (Style)FindResource("SidebarButtonStyle");
            btnLicenca.Style = (Style)FindResource("SidebarButtonStyle");
            btnKorisnici.Style = (Style)FindResource("SidebarButtonStyle");
            btnPomoc.Style = (Style)FindResource("SidebarButtonStyle");

            if (activeButton != null)
            {
                activeButton.Style = (Style)FindResource("SidebarActiveButtonStyle");
            }
        }

        private void BtnHome_Click(object sender, RoutedEventArgs e)
        {
            SetActiveSidebarButton(btnHome);
            LoadHomeView();
        }

        private void BtnOsnovni_Click(object sender, RoutedEventArgs e)
        {

        }

        public void NavigateToPomoc()
        {
            SetActiveSidebarButton(btnPomoc);
            LoadPomocView();
        }

        private void LoadPomocView()
        {
            var pomocView = new PomocView
            {
                HorizontalAlignment = HorizontalAlignment.Stretch,
                VerticalAlignment = VerticalAlignment.Stretch
            };

            var container = new Grid();
            container.Children.Add(pomocView);

            MainContent.Content = container;
        }

        private void BtnPomoc_Click(object sender, RoutedEventArgs e)
        {
            NavigateToPomoc();
        }

        private void LoadHomeView()
        {
            var homeView = new HomeView
            {
                HorizontalAlignment = HorizontalAlignment.Stretch,
                VerticalAlignment = VerticalAlignment.Stretch
            };
            homeView.sLogovaniKorisnik = PunoIme;

            var container = new Grid();
            container.Children.Add(homeView);

            MainContent.Content = container;
        }

        private void BtnOdjava_Click(object sender, RoutedEventArgs e)
        {
            // Zatvara aplikaciju
            Application.Current.Shutdown();
        }

        public void NavigateToIzvozDokumentacije(string initialPath = "")
        {
            SetActiveSidebarButton(btnIzvozDokumentacije);
            LoadIzvozDokumentacijeView(initialPath);
        }

        private void LoadIzvozDokumentacijeView(string initialPath = "")
        {
            var exportView = new IzvozDokumentacijeView(initialPath)
            {
                HorizontalAlignment = HorizontalAlignment.Stretch,
                VerticalAlignment = VerticalAlignment.Stretch
            };

            var container = new Grid();
            container.Children.Add(exportView);

            MainContent.Content = container;
        }

        private void BtnIzvozDokumentacije_Click(object sender, RoutedEventArgs e)
        {
            NavigateToIzvozDokumentacije();
        }

        public void NavigateToUvozSablon()
        {
            SetActiveSidebarButton(btnUvozSablon);
            LoadUvozSablonView();
        }

        private void LoadUvozSablonView()
        {
            var uvozView = new UvozSablonView
            {
                HorizontalAlignment = HorizontalAlignment.Stretch,
                VerticalAlignment = VerticalAlignment.Stretch
            };

            var container = new Grid();
            container.Children.Add(uvozView);

            MainContent.Content = container;
        }

        private void BtnUvozSablon_Click(object sender, RoutedEventArgs e)
        {
            NavigateToUvozSablon();
        }

        public void NavigateToLicenca()
        {
            SetActiveSidebarButton(btnLicenca);
            LoadLicencaView();
        }

        private void LoadLicencaView()
        {
            var licView = new LicencaView
            {
                HorizontalAlignment = HorizontalAlignment.Stretch,
                VerticalAlignment = VerticalAlignment.Stretch
            };

            var container = new Grid();
            container.Children.Add(licView);

            MainContent.Content = container;
        }

        private void BtnLicenca_Click(object sender, RoutedEventArgs e)
        {
            NavigateToLicenca();
        }

        private void BtnKorisnici_Click(object sender, RoutedEventArgs e)
        {
            if (!IsAdmin) return;
            SetActiveSidebarButton(btnKorisnici);

            var container = new Grid();
            container.Children.Add(new KorisniciView(LoginUsername)
            {
                HorizontalAlignment = HorizontalAlignment.Stretch,
                VerticalAlignment = VerticalAlignment.Stretch
            });
            MainContent.Content = container;
        }


        private void BtnFastimate_Click(object sender, RoutedEventArgs e)
        {
            PomocView.OpenHelpPdfSafely();
        }
    }
}
