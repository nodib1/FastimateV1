using System;
using System.IO;
using System.Security.AccessControl;
using System.Security.Principal;

namespace FASTIMATE
{
    public static class AppPaths
    {
        private static bool _permissionsChecked;

        /// <summary>
        /// Daje grupi "Users" pravo izmjene (Modify) na folder/fajl, da svi RDP korisnici mogu snimati
        /// zajedničke fajlove, bez obzira koji Windows korisnik ih je prvi napravio.
        /// Uspije samo ako trenutni korisnik ima pravo mijenjati dozvole (vlasnik ili admin); inače se tiho preskače.
        /// </summary>
        public static void GrantUsersModify(string path, bool isDirectory)
        {
            try
            {
                var users = new SecurityIdentifier(WellKnownSidType.BuiltinUsersSid, null);

                if (isDirectory)
                {
                    var dirSec = Directory.GetAccessControl(path);
                    var rule = new FileSystemAccessRule(users, FileSystemRights.Modify,
                        InheritanceFlags.ContainerInherit | InheritanceFlags.ObjectInherit,
                        PropagationFlags.None, AccessControlType.Allow);

                    if (HasRule(dirSec.GetAccessRules(true, true, typeof(SecurityIdentifier)), users, rule.InheritanceFlags)) return;

                    dirSec.AddAccessRule(rule);
                    Directory.SetAccessControl(path, dirSec);
                }
                else
                {
                    var fileSec = File.GetAccessControl(path);
                    if (HasRule(fileSec.GetAccessRules(true, true, typeof(SecurityIdentifier)), users, InheritanceFlags.None)) return;

                    fileSec.AddAccessRule(new FileSystemAccessRule(users, FileSystemRights.Modify, AccessControlType.Allow));
                    File.SetAccessControl(path, fileSec);
                }
            }
            catch { }
        }

        private static bool HasRule(AuthorizationRuleCollection rules, SecurityIdentifier sid, InheritanceFlags requiredInheritance)
        {
            foreach (FileSystemAccessRule r in rules)
            {
                if (r.AccessControlType == AccessControlType.Allow
                    && sid.Equals(r.IdentityReference)
                    && (r.FileSystemRights & FileSystemRights.Modify) == FileSystemRights.Modify
                    && (r.InheritanceFlags & requiredInheritance) == requiredInheritance)
                {
                    return true;
                }
            }
            return false;
        }

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

                if (!_permissionsChecked)
                {
                    _permissionsChecked = true;
                    GrantUsersModify(commonPath, true);
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
