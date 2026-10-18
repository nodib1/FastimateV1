using System;
using System.IO;

namespace FASTIMATE
{
    public static class AppPaths
    {
        /// <summary>
        /// Vraća zajednički folder sa bazama podataka koji dijele SVI korisnici na Remote Desktop (RDP / Terminal) serveru:
        /// C:\ProgramData\FASTIMATE\
        /// </summary>
        public static string GetDbDirectory()
        {
            string commonPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData),
                "FASTIMATE"
            );

            try
            {
                if (!Directory.Exists(commonPath))
                {
                    Directory.CreateDirectory(commonPath);
                }

                // Migracija postojecih fajlova iz AppData korisnika u ProgramData ako noviji fajl još ne postoji
                string userAppDataPath = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "FASTIMATE"
                );

                if (Directory.Exists(userAppDataPath))
                {
                    string[] filesToMigrate = new string[]
                    {
                        "sheet_metal_raw_materials.txt",
                        "erp_fasteners.txt",
                        "k_factor_rules.txt",
                        "PRORACUN_VREMENA_TEHNOLOGIJE_CUSTOM.xlsx",
                        "PRORACUN_PONUDE_CUSTOM.xlsx"
                    };

                    foreach (string fileName in filesToMigrate)
                    {
                        string oldFile = Path.Combine(userAppDataPath, fileName);
                        string newFile = Path.Combine(commonPath, fileName);
                        if (File.Exists(oldFile) && !File.Exists(newFile))
                        {
                            try { File.Copy(oldFile, newFile, true); } catch { }
                        }
                    }
                }

                return commonPath;
            }
            catch
            {
                string fallback = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "FASTIMATE"
                );
                if (!Directory.Exists(fallback)) Directory.CreateDirectory(fallback);
                return fallback;
            }
        }
    }
}
