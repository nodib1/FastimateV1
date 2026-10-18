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

        public MainNovi(string username = "admin", bool isAdmin = true)
        {
            InitializeComponent();
            IsAdmin = isAdmin;

            if (isAdmin)
            {
                labelkorisnik.Text = username + " (ADMIN PANEL)";
                if (btnUvozSablon != null) btnUvozSablon.Visibility = Visibility.Visible;
                if (btnLicenca != null) btnLicenca.Visibility = Visibility.Visible;
            }
            else
            {
                labelkorisnik.Text = username;
                if (btnUvozSablon != null) btnUvozSablon.Visibility = Visibility.Collapsed;
                if (btnLicenca != null) btnLicenca.Visibility = Visibility.Collapsed;
            }

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
            homeView.sLogovaniKorisnik = labelkorisnik.Text;

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


        private void BtnFastimate_Click(object sender, RoutedEventArgs e)
        {
            PomocView.OpenHelpPdfSafely();
        }
    }
}
