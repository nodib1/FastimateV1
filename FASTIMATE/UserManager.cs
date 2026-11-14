using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;

namespace FASTIMATE
{
    public class AppUser
    {
        public string Username { get; set; }
        public string Ime { get; set; }
        public string Prezime { get; set; }
        public bool IsAdmin { get; set; }
        public bool Aktivan { get; set; } = true;
        public string Salt { get; set; }
        public string PasswordHash { get; set; }

        public string PunoIme
        {
            get
            {
                string full = ((Ime ?? "") + " " + (Prezime ?? "")).Trim();
                return string.IsNullOrEmpty(full) ? Username : full;
            }
        }

        public string AdminText { get { return IsAdmin ? "Da" : "Ne"; } }
        public string AktivanText { get { return Aktivan ? "Da" : "Ne"; } }
    }

    /// <summary>
    /// Korisnici se čuvaju u C:\ProgramData\FASTIMATE\fastimate_u.txt (jedan red = jedan korisnik).
    /// Format: Username;Ime;Prezime;IsAdmin(1/0);Aktivan(1/0);Salt;PasswordHash
    /// Lozinke su hashovane (PBKDF2-SHA256 + nasumični salt), nikad se ne čuvaju kao tekst.
    /// </summary>
    public static class UserManager
    {
        private const int SaltSize = 16;
        private const int HashSize = 32;
        private const int Iterations = 100000;

        // Ugrađeni (rezervni) admin - radi SAMO ako u fajlu ne postoji nijedan aktivan admin.
        private const string FallbackAdminUsername = "admin";
        private static readonly string[] FallbackAdminPasswords = { "mtp\"=\"&!" };

        private static string FilePath
        {
            get { return Path.Combine(AppPaths.GetDbDirectory(), "fastimate_u.txt"); }
        }

        // Potpis fajla (HMAC-SHA256): svaka ručna izmjena fajla poništava potpis.
        private const string SignaturePrefix = "#SIG;";
        private static readonly byte[] SignatureKey = Convert.FromBase64String("yoIyjnptsqQ7e202D/bNW2isVA/cuc9M1J0ZhbD7/mo=");

        private static string ComputeSignature(List<string> dataLines)
        {
            using (var hmac = new HMACSHA256(SignatureKey))
            {
                byte[] bytes = Encoding.UTF8.GetBytes(string.Join("\n", dataLines));
                return Convert.ToBase64String(hmac.ComputeHash(bytes));
            }
        }

        /// <summary>
        /// True ako je fajl ručno mijenjan (potpis ne odgovara sadržaju).
        /// </summary>
        public static bool IsFileTampered()
        {
            bool valid;
            LoadUsers(out valid);
            return !valid;
        }

        public static List<AppUser> LoadUsers()
        {
            bool valid;
            return LoadUsers(out valid);
        }

        public static List<AppUser> LoadUsers(out bool signatureValid)
        {
            EnsureFileExists();
            var users = new List<AppUser>();
            var dataLines = new List<string>();
            string signature = null;
            signatureValid = false;

            try
            {
                foreach (string raw in File.ReadAllLines(FilePath, Encoding.UTF8))
                {
                    string line = raw.Trim();
                    if (line.StartsWith(SignaturePrefix)) { signature = line.Substring(SignaturePrefix.Length).Trim(); continue; }
                    if (line.Length == 0 || line.StartsWith("#")) continue;

                    dataLines.Add(line);

                    string[] p = line.Split(';');
                    if (p.Length < 7 || string.IsNullOrWhiteSpace(p[0])) continue;

                    users.Add(new AppUser
                    {
                        Username = p[0].Trim(),
                        Ime = p[1].Trim(),
                        Prezime = p[2].Trim(),
                        IsAdmin = p[3].Trim() == "1",
                        Aktivan = p[4].Trim() == "1",
                        Salt = p[5].Trim(),
                        PasswordHash = p[6].Trim()
                    });
                }

                if (signature != null)
                    signatureValid = signature == ComputeSignature(dataLines);
                else
                    signatureValid = dataLines.Count == 0; // stari prazan fajl bez potpisa je u redu
            }
            catch { }

            return users;
        }

        private static void SaveUsers(List<AppUser> users)
        {
            var dataLines = new List<string>();
            foreach (var u in users)
            {
                dataLines.Add(string.Join(";",
                    Clean(u.Username), Clean(u.Ime), Clean(u.Prezime),
                    u.IsAdmin ? "1" : "0", u.Aktivan ? "1" : "0",
                    (u.Salt ?? "").Trim(), (u.PasswordHash ?? "").Trim()));
            }

            var lines = new List<string>
            {
                "# FASTIMATE korisnici - ne uređivati ručno (fajl je potpisan, ručna izmjena blokira prijavu).",
                "# Username;Ime;Prezime;IsAdmin;Aktivan;Salt;PasswordHash"
            };
            lines.AddRange(dataLines);
            lines.Add(SignaturePrefix + ComputeSignature(dataLines));

            // Direktan upis (zadržava postojeće dozvole na fajlu).
            File.WriteAllLines(FilePath, lines, Encoding.UTF8);
        }

        private static bool _fileChecked;

        private static void EnsureFileExists()
        {
            if (!File.Exists(FilePath))
            {
                // Prvo pokretanje: prazan fajl. Prijava ide preko ugrađenog admina dok se ne doda admin u fajl.
                try { SaveUsers(new List<AppUser>()); } catch { }
            }

            // Ko god je napravio fajl (vlasnik), daje svim korisnicima na serveru pravo izmjene.
            if (!_fileChecked && File.Exists(FilePath))
            {
                _fileChecked = true;
                AppPaths.GrantUsersModify(FilePath, false);
            }
        }

        public enum LoginError { None, PraznoKorisnickoIme, PraznaLozinka, NepoznatKorisnik, Deaktiviran, PogresnaLozinka, FajlMijenjan }

        public static AppUser Authenticate(string username, string password)
        {
            LoginError err;
            return Authenticate(username, password, out err);
        }

        /// <summary>
        /// Vraća korisnika ako su podaci ispravni, inače null i razlog u 'error'.
        /// </summary>
        public static AppUser Authenticate(string username, string password, out LoginError error)
        {
            username = (username ?? "").Trim();
            if (username.Length == 0) { error = LoginError.PraznoKorisnickoIme; return null; }
            if (string.IsNullOrEmpty(password)) { error = LoginError.PraznaLozinka; return null; }

            bool signatureValid;
            var users = LoadUsers(out signatureValid);

            if (!signatureValid)
            {
                // Fajl je ručno mijenjan: ne vjerujemo mu. Ulaz ima samo ugrađeni admin, da može popraviti korisnike.
                if (username.Equals(FallbackAdminUsername, StringComparison.OrdinalIgnoreCase) && FallbackAdminPasswords.Contains(password))
                {
                    error = LoginError.None;
                    return new AppUser { Username = FallbackAdminUsername, Ime = "admin", Prezime = "", IsAdmin = true, Aktivan = true };
                }
                error = LoginError.FajlMijenjan;
                return null;
            }

            var user = users.FirstOrDefault(u => u.Username.Equals(username, StringComparison.OrdinalIgnoreCase));

            if (user == null) error = LoginError.NepoznatKorisnik;
            else if (!user.Aktivan) error = LoginError.Deaktiviran;
            else if (!VerifyPassword(user, password)) error = LoginError.PogresnaLozinka;
            else { error = LoginError.None; return user; }

            // Rezervni admin iz koda - samo dok u fajlu nema nijednog aktivnog admina.
            bool hasActiveAdmin = users.Any(u => u.IsAdmin && u.Aktivan);
            if (!hasActiveAdmin && username.Equals(FallbackAdminUsername, StringComparison.OrdinalIgnoreCase))
            {
                if (FallbackAdminPasswords.Contains(password))
                {
                    error = LoginError.None;
                    return new AppUser { Username = FallbackAdminUsername, Ime = "admin", Prezime = "", IsAdmin = true, Aktivan = true };
                }
                error = LoginError.PogresnaLozinka;
            }

            return null;
        }

        public static string LoginErrorMessage(LoginError error, string username)
        {
            switch (error)
            {
                case LoginError.PraznoKorisnickoIme: return "Unesite korisničko ime.";
                case LoginError.PraznaLozinka: return "Unesite lozinku.";
                case LoginError.NepoznatKorisnik: return "Korisnik '" + (username ?? "").Trim() + "' ne postoji.";
                case LoginError.Deaktiviran: return "Korisnik '" + (username ?? "").Trim() + "' je deaktiviran.\nObratite se administratoru.";
                case LoginError.PogresnaLozinka: return "Pogrešna lozinka!";
                case LoginError.FajlMijenjan: return "Fajl s korisnicima je ručno mijenjan i nije važeći.\nObratite se administratoru.";
                default: return "";
            }
        }

        /// <summary>
        /// Dodaje novog ili mijenja postojećeg korisnika. Ako je password prazan kod izmjene, stara lozinka ostaje.
        /// originalUsername = null za novog korisnika.
        /// </summary>
        public static void SaveUser(string originalUsername, AppUser data, string newPassword)
        {
            string username = (data.Username ?? "").Trim();
            if (username.Length == 0) throw new ArgumentException("Korisničko ime je obavezno.");
            if (username.Contains(";")) throw new ArgumentException("Korisničko ime ne smije sadržavati znak ';'.");

            var users = LoadUsers(); // uvijek svježe sa diska (više RDP korisnika)

            bool isNew = string.IsNullOrEmpty(originalUsername);
            AppUser existing = isNew ? null : users.FirstOrDefault(u => u.Username.Equals(originalUsername, StringComparison.OrdinalIgnoreCase));

            if (!isNew && existing == null) throw new InvalidOperationException("Korisnik više ne postoji u fajlu.");

            bool nameTaken = users.Any(u => u.Username.Equals(username, StringComparison.OrdinalIgnoreCase) && u != existing);
            if (nameTaken) throw new ArgumentException("Korisničko ime '" + username + "' već postoji.");

            if (isNew && string.IsNullOrEmpty(newPassword)) throw new ArgumentException("Lozinka je obavezna za novog korisnika.");

            var target = existing ?? new AppUser();
            target.Username = username;
            target.Ime = (data.Ime ?? "").Trim();
            target.Prezime = (data.Prezime ?? "").Trim();
            target.IsAdmin = data.IsAdmin;
            target.Aktivan = data.Aktivan;
            if (!string.IsNullOrEmpty(newPassword)) SetPassword(target, newPassword);

            if (isNew) users.Add(target);
            SaveUsers(users);
        }

        public static void DeleteUser(string username)
        {
            var users = LoadUsers();
            users.RemoveAll(u => u.Username.Equals(username, StringComparison.OrdinalIgnoreCase));
            SaveUsers(users);
        }

        private static void SetPassword(AppUser user, string password)
        {
            byte[] salt = new byte[SaltSize];
            using (var rng = RandomNumberGenerator.Create()) rng.GetBytes(salt);

            user.Salt = Convert.ToBase64String(salt);
            user.PasswordHash = Convert.ToBase64String(Hash(password, salt));
        }

        private static bool VerifyPassword(AppUser user, string password)
        {
            try
            {
                byte[] salt = Convert.FromBase64String(user.Salt);
                byte[] expected = Convert.FromBase64String(user.PasswordHash);
                byte[] actual = Hash(password, salt);

                if (expected.Length != actual.Length) return false;
                int diff = 0;
                for (int i = 0; i < actual.Length; i++) diff |= expected[i] ^ actual[i];
                return diff == 0;
            }
            catch
            {
                return false;
            }
        }

        private static byte[] Hash(string password, byte[] salt)
        {
            using (var pbkdf2 = new Rfc2898DeriveBytes(password, salt, Iterations, HashAlgorithmName.SHA256))
            {
                return pbkdf2.GetBytes(HashSize);
            }
        }

        private static string Clean(string s)
        {
            return (s ?? "").Replace(";", ",").Replace("\r", " ").Replace("\n", " ").Trim();
        }
    }
}
