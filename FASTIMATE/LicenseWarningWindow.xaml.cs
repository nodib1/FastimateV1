using System;
using System.Windows;

namespace FASTIMATE
{
    public partial class LicenseWarningWindow : Window
    {
        public LicenseWarningWindow(int remainingDays, DateTime expDate)
        {
            InitializeComponent();
            txtWarningDetail.Text = $"Vaša FASTIMATE licenca ističe za {remainingDays} dana (dana {expDate:dd.MM.yyyy.}).";
        }

        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void btnOk_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}
