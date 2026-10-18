using System;
using System.IO;
using System.Windows;
using Microsoft.Win32;

namespace FASTIMATE
{
    public static class TemplateManager
    {
        private static readonly string AppDataFolder = AppPaths.GetDbDirectory();

        private static readonly string CustomTemplatePath = Path.Combine(
            AppDataFolder,
            "PRORACUN_VREMENA_TEHNOLOGIJE_CUSTOM.xlsx"
        );

        /// <summary>
        /// Vraća bajtove aktivnog Excel šablona (uvozni šablon ako postoji, ili ugrađeni resurs za Ponudu/Tehnologiju).
        /// </summary>
        public static byte[] GetActiveTemplateBytes(int tip = 2)
        {
            try
            {
                if (File.Exists(CustomTemplatePath))
                {
                    return File.ReadAllBytes(CustomTemplatePath);
                }
            }
            catch { }

            if (tip == 1)
            {
                try
                {
                    byte[] ponudaBytes = Properties.Resources.PRORACUN_VREMENA_PONUDE;
                    if (ponudaBytes != null && ponudaBytes.Length > 0)
                    {
                        return ponudaBytes;
                    }
                }
                catch { }
            }

            return Properties.Resources.PRORACUN_VREMENA_TEHNOLOGIJE;
        }

        /// <summary>
        /// Omogućava odabir i uvoz novog sirovog Excel šablona u bazu resursa aplikacije.
        /// </summary>
        public static bool ImportNewTemplate()
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Title = "Odaberite novi sirovi Excel šablon za Tehnologiju",
                Filter = "Excel fajlovi (*.xlsx;*.xlsm)|*.xlsx;*.xlsm|Svi fajlovi (*.*)|*.*",
                Multiselect = false
            };

            if (openFileDialog.ShowDialog() == true)
            {
                try
                {
                    if (!Directory.Exists(AppDataFolder))
                    {
                        Directory.CreateDirectory(AppDataFolder);
                    }

                    byte[] newBytes = File.ReadAllBytes(openFileDialog.FileName);
                    File.WriteAllBytes(CustomTemplatePath, newBytes);

                    MessageBox.Show("Novi sirovi Excel šablon je uspješno uvezen u resurse aplikacije!\n\nSvi budući obračuni će automatski koristiti ovaj novi šablon.",
                        "Uvoz Uspješan", MessageBoxButton.OK, MessageBoxImage.Information);
                    return true;
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Greška pri uvozu Excel šablona: {ex.Message}",
                        "Greška", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }

            return false;
        }

        /// <summary>
        /// Vraća šablon na fabrički ugrađeni resurs.
        /// </summary>
        public static void ResetToDefaultTemplate()
        {
            try
            {
                if (File.Exists(CustomTemplatePath))
                {
                    File.Delete(CustomTemplatePath);
                    MessageBox.Show("Excel šablon je uspješno vraćen na fabrički ugrađeni resurs.",
                        "Fabrička Podešavanja", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                else
                {
                    MessageBox.Show("Trenutno je u upotrebi fabrički ugrađeni resurs.",
                        "Informacija", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Greška pri vraćanju fabričkog šablona: {ex.Message}",
                    "Greška", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// Provjerava da li je uvezen prilagođeni šablon.
        /// </summary>
        public static bool HasCustomTemplate()
        {
            return File.Exists(CustomTemplatePath);
        }
    }
}
