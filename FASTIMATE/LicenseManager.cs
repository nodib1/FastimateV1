using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;

namespace FASTIMATE
{
    public class LicenseStatusInfo
    {
        public bool IsValid { get; set; }
        public bool IsExpired { get; set; }
        public DateTime ExpirationDate { get; set; }
        public int RemainingDays { get; set; }
        public string CompanyName { get; set; }
        public string StatusText { get; set; }
        public bool Show15DaysWarning { get; set; }
    }

    public static class LicenseManager
    {
        private static readonly string SecretSalt = "FASTIMATE_MTP_2026_SECRET_SALT_KEY_8829";

        private static readonly string LicenseFilePath = Path.Combine(
            AppPaths.GetDbDirectory(),
            "sys_config.dat"
        );

        public static LicenseStatusInfo CheckLicense()
        {
            var info = new LicenseStatusInfo
            {
                IsValid = false,
                IsExpired = true,
                ExpirationDate = DateTime.MinValue,
                RemainingDays = 0,
                CompanyName = "Nije uneseno",
                StatusText = "Licenca nije aktivirana",
                Show15DaysWarning = false
            };

            if (!File.Exists(LicenseFilePath))
            {
                return info;
            }

            try
            {
                byte[] fileBytes = File.ReadAllBytes(LicenseFilePath);
                string payload = DecryptStringFromBytes(fileBytes, SecretSalt);

                if (string.IsNullOrWhiteSpace(payload) || !payload.Contains("|"))
                {
                    info.StatusText = "Nevažeća licenca (oštećen fajl)";
                    return info;
                }

                string[] parts = payload.Split('|');
                if (parts.Length < 3)
                {
                    if (parts.Length == 2)
                    {
                        string expStrOld = parts[0].Trim();
                        string sigOld = parts[1].Trim();
                        if (DateTime.TryParseExact(expStrOld, "yyyy-MM-dd", System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out DateTime dOld))
                        {
                            info.ExpirationDate = dOld;
                            info.RemainingDays = (int)(dOld.Date - DateTime.Now.Date).TotalDays;
                            info.IsValid = info.RemainingDays >= 0;
                            info.IsExpired = !info.IsValid;
                            info.CompanyName = "Glavna Licenca";
                            info.StatusText = info.IsValid ? $"AKTIVNA (Važi do {dOld:dd.MM.yyyy.})" : $"ISTEKLA ({dOld:dd.MM.yyyy.})";
                            info.Show15DaysWarning = info.IsValid && info.RemainingDays <= 15;
                            return info;
                        }
                    }
                    info.StatusText = "Nevažeći format licence";
                    return info;
                }

                string expStr = parts[0].Trim();
                string company = parts[1].Trim();
                string sig = parts[2].Trim();

                string expectedSig = GenerateSignature($"{expStr}|{company}", SecretSalt);
                if (!string.Equals(sig, expectedSig, StringComparison.OrdinalIgnoreCase))
                {
                    info.StatusText = "Nevažeći digitalni potpis licence";
                    return info;
                }

                if (!DateTime.TryParseExact(expStr, "yyyy-MM-dd", System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out DateTime expDate))
                {
                    info.StatusText = "Nevažeći datum licence";
                    return info;
                }

                info.ExpirationDate = expDate;
                info.CompanyName = string.IsNullOrWhiteSpace(company) ? "Nije uneseno" : company;
                info.RemainingDays = (int)(expDate.Date - DateTime.Now.Date).TotalDays;

                if (info.RemainingDays < 0)
                {
                    info.IsValid = false;
                    info.IsExpired = true;
                    info.StatusText = $"Licenca je ISTEKLA ({expDate:dd.MM.yyyy.})";
                }
                else
                {
                    info.IsValid = true;
                    info.IsExpired = false;
                    info.StatusText = $"AKTIVNA (Važi do {expDate:dd.MM.yyyy.})";
                    if (info.RemainingDays <= 15)
                    {
                        info.Show15DaysWarning = true;
                    }
                }

                return info;
            }
            catch
            {
                info.StatusText = "Greška pri provjeri licence";
                return info;
            }
        }

        public static bool CreateAndSaveLicense(string companyName, int durationDays, out string message)
        {
            message = "";
            if (string.IsNullOrWhiteSpace(companyName))
            {
                message = "Unesite naziv firme / klijenta.";
                return false;
            }

            if (durationDays <= 0)
            {
                message = "Unesite važeći broj dana (veći od 0).";
                return false;
            }

            try
            {
                DateTime expDate = DateTime.Now.AddDays(durationDays);
                string expStr = expDate.ToString("yyyy-MM-dd");
                string company = companyName.Trim();
                string sig = GenerateSignature($"{expStr}|{company}", SecretSalt);

                string payload = $"{expStr}|{company}|{sig}";
                byte[] encryptedData = EncryptStringToBytes(payload, SecretSalt);

                string dir = Path.GetDirectoryName(LicenseFilePath);
                if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);

                File.WriteAllBytes(LicenseFilePath, encryptedData);
                message = $"Licenca za '{company}' je uspješno kreirana na {durationDays} dana (važi do {expDate:dd.MM.yyyy.})!";
                return true;
            }
            catch (Exception ex)
            {
                message = $"Greška pri čuvanju licence: {ex.Message}";
                return false;
            }
        }

        public static bool ActivateLicenseKey(string inputKey, out string message)
        {
            message = "";
            if (string.IsNullOrWhiteSpace(inputKey))
            {
                message = "Unesite kod licence.";
                return false;
            }

            try
            {
                string cleanKey = inputKey.Trim();
                string payload = "";

                try
                {
                    byte[] tokenBytes = Convert.FromBase64String(cleanKey);
                    payload = DecryptStringFromBytes(tokenBytes, SecretSalt);
                }
                catch { }

                if (string.IsNullOrWhiteSpace(payload) && cleanKey.StartsWith("FST-", StringComparison.OrdinalIgnoreCase))
                {
                    string[] parts = cleanKey.Split('-');
                    if (parts.Length >= 3)
                    {
                        string datePart = parts[1].Trim();
                        if (DateTime.TryParseExact(datePart, "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out DateTime d))
                        {
                            string formattedDateStr = d.ToString("yyyy-MM-dd");
                            string defaultCompany = "Glavna Licenca";
                            string sig = GenerateSignature($"{formattedDateStr}|{defaultCompany}", SecretSalt);
                            payload = $"{formattedDateStr}|{defaultCompany}|{sig}";
                        }
                    }
                }

                if (string.IsNullOrWhiteSpace(payload) || !payload.Contains("|"))
                {
                    message = "Neispravan kod licence. Provjerite uneseni kod.";
                    return false;
                }

                string[] p = payload.Split('|');
                if (p.Length < 3)
                {
                    message = "Neispravan format licence.";
                    return false;
                }

                string expStr = p[0];
                string company = p[1];
                string sigStr = p[2];

                string expSig = GenerateSignature($"{expStr}|{company}", SecretSalt);
                if (!string.Equals(sigStr, expSig, StringComparison.OrdinalIgnoreCase))
                {
                    message = "Neispravan digitalni potpis licence.";
                    return false;
                }

                byte[] encryptedData = EncryptStringToBytes(payload, SecretSalt);
                string dir = Path.GetDirectoryName(LicenseFilePath);
                if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);

                File.WriteAllBytes(LicenseFilePath, encryptedData);
                message = "Licenca je uspješno aktivirana!";
                return true;
            }
            catch (Exception ex)
            {
                message = $"Greška pri aktivaciji licence: {ex.Message}";
                return false;
            }
        }

        private static string GenerateSignature(string text, string salt)
        {
            using (SHA256 sha = SHA256.Create())
            {
                byte[] bytes = sha.ComputeHash(Encoding.UTF8.GetBytes(text + "_SALT_" + salt));
                StringBuilder sb = new StringBuilder();
                foreach (byte b in bytes)
                {
                    sb.Append(b.ToString("x2"));
                }
                return sb.ToString();
            }
        }

        private static byte[] EncryptStringToBytes(string plainText, string salt)
        {
            byte[] key = SHA256.Create().ComputeHash(Encoding.UTF8.GetBytes(salt));
            byte[] iv = new byte[16];
            for (int i = 0; i < 16; i++) iv[i] = key[i];

            using (Aes aes = Aes.Create())
            {
                aes.Key = key;
                aes.IV = iv;
                ICryptoTransform encryptor = aes.CreateEncryptor(aes.Key, aes.IV);
                using (MemoryStream ms = new MemoryStream())
                {
                    using (CryptoStream cs = new CryptoStream(ms, encryptor, CryptoStreamMode.Write))
                    {
                        using (StreamWriter sw = new StreamWriter(cs))
                        {
                            sw.Write(plainText);
                        }
                        return ms.ToArray();
                    }
                }
            }
        }

        private static string DecryptStringFromBytes(byte[] cipherText, string salt)
        {
            byte[] key = SHA256.Create().ComputeHash(Encoding.UTF8.GetBytes(salt));
            byte[] iv = new byte[16];
            for (int i = 0; i < 16; i++) iv[i] = key[i];

            using (Aes aes = Aes.Create())
            {
                aes.Key = key;
                aes.IV = iv;
                ICryptoTransform decryptor = aes.CreateDecryptor(aes.Key, aes.IV);
                using (MemoryStream ms = new MemoryStream(cipherText))
                {
                    using (CryptoStream cs = new CryptoStream(ms, decryptor, CryptoStreamMode.Read))
                    {
                        using (StreamReader sr = new StreamReader(cs))
                        {
                            return sr.ReadToEnd();
                        }
                    }
                }
            }
        }
    }
}
