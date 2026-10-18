using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace FASTIMATE
{
    public partial class LicencaView : UserControl
    {
        public LicencaView()
        {
            InitializeComponent();
        }

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            RefreshLicenseStatus();
        }

        private void RefreshLicenseStatus()
        {
            var info = LicenseManager.CheckLicense();

            txtCompanyNameVal.Text = string.IsNullOrWhiteSpace(info.CompanyName) ? "Nije uneseno" : info.CompanyName;

            if (info.IsValid)
            {
                txtStatusVal.Text = "AKTIVNA";
                txtStatusVal.Foreground = (Brush)new BrushConverter().ConvertFromString("#81C784");

                txtExpDateVal.Text = info.ExpirationDate.ToString("dd.MM.yyyy.");
                txtRemainingDaysVal.Text = $"{info.RemainingDays} dana";

                if (info.Show15DaysWarning)
                {
                    txtBadge.Text = "ISTIČE USKORO!";
                    txtBadge.Foreground = (Brush)new BrushConverter().ConvertFromString("#FFB74D");
                    badgeBorder.Background = (Brush)new BrushConverter().ConvertFromString("#332A00");
                    badgeBorder.BorderBrush = (Brush)new BrushConverter().ConvertFromString("#FFB74D");
                }
                else
                {
                    txtBadge.Text = "VAŽEĆA LICENCA";
                    txtBadge.Foreground = (Brush)new BrushConverter().ConvertFromString("#81C784");
                    badgeBorder.Background = (Brush)new BrushConverter().ConvertFromString("#1B3320");
                    badgeBorder.BorderBrush = (Brush)new BrushConverter().ConvertFromString("#81C784");
                }
            }
            else
            {
                txtStatusVal.Text = info.IsExpired ? "ISTEKLA" : "NEAKTIVNA";
                txtStatusVal.Foreground = (Brush)new BrushConverter().ConvertFromString("#FF5252");

                txtExpDateVal.Text = info.ExpirationDate > DateTime.MinValue ? info.ExpirationDate.ToString("dd.MM.yyyy.") : "--";
                txtRemainingDaysVal.Text = "0 dana";

                txtBadge.Text = info.IsExpired ? "ISTEKLA" : "NIJE AKTIVIRANA";
                txtBadge.Foreground = (Brush)new BrushConverter().ConvertFromString("#FF5252");
                badgeBorder.Background = (Brush)new BrushConverter().ConvertFromString("#331B1B");
                badgeBorder.BorderBrush = (Brush)new BrushConverter().ConvertFromString("#FF5252");
            }

            if (txtInputCompany != null && string.IsNullOrWhiteSpace(txtInputCompany.Text))
            {
                txtInputCompany.Text = info.CompanyName != "Nije uneseno" && info.CompanyName != "Glavna Licenca" ? info.CompanyName : "";
            }

            if (txtInputDays != null && string.IsNullOrWhiteSpace(txtInputDays.Text))
            {
                txtInputDays.Text = info.RemainingDays > 0 ? info.RemainingDays.ToString() : "365";
            }
        }

        private void btnSacuvaj_Click(object sender, RoutedEventArgs e)
        {
            string pin = txtMasterPin.Password != null ? txtMasterPin.Password.Trim() : "";
            if (string.IsNullOrWhiteSpace(pin) || !pin.EndsWith("2023"))
            {
                lblMessage.Foreground = (Brush)new BrushConverter().ConvertFromString("#FF5252");
                lblMessage.Text = "Neispravan Autentifikacijski kod! Nemate autorizaciju za izmjenu licence.";
                txtMasterPin.Clear();
                txtMasterPin.Focus();
                return;
            }

            string company = txtInputCompany.Text != null ? txtInputCompany.Text.Trim() : "";
            string daysStr = txtInputDays.Text != null ? txtInputDays.Text.Trim() : "";

            if (string.IsNullOrWhiteSpace(company))
            {
                lblMessage.Foreground = (Brush)new BrushConverter().ConvertFromString("#FF5252");
                lblMessage.Text = "Unesite naziv firme / klijenta.";
                return;
            }

            if (!int.TryParse(daysStr, out int days) || days <= 0)
            {
                lblMessage.Foreground = (Brush)new BrushConverter().ConvertFromString("#FF5252");
                lblMessage.Text = "Unesite ispravan broj dana (npr. 365).";
                return;
            }

            if (LicenseManager.CreateAndSaveLicense(company, days, out string msg))
            {
                lblMessage.Foreground = (Brush)new BrushConverter().ConvertFromString("#81C784");
                lblMessage.Text = msg;
                txtMasterPin.Clear();
                RefreshLicenseStatus();
            }
            else
            {
                lblMessage.Foreground = (Brush)new BrushConverter().ConvertFromString("#FF5252");
                lblMessage.Text = msg;
            }
        }
    }
}
