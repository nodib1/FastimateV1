using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows;

namespace FASTIMATE
{
    public class ErpFastenerItem
    {
        public string ErpIdent { get; set; }
        public string Standard { get; set; }
        public string Naziv { get; set; }
        public string Velicina { get; set; }
    }

    public static class FastenerDbManager
    {
        private static readonly List<ErpFastenerItem> Fasteners = new List<ErpFastenerItem>();
        private static readonly ObservableCollection<ErpFastenerItem> FastenersCollection = new ObservableCollection<ErpFastenerItem>();

        private static readonly string DbFilePath = Path.Combine(
            AppPaths.GetDbDirectory(),
            "erp_fasteners.txt"
        );

        static FastenerDbManager()
        {
            EnsureDirectory();
            LoadData();
        }

        private static void EnsureDirectory()
        {
            string dir = Path.GetDirectoryName(DbFilePath);
            if (!Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }
        }

        public static ObservableCollection<ErpFastenerItem> GetAllItems()
        {
            return FastenersCollection;
        }

        public static void LoadData(bool forceReload = false)
        {
            if (!forceReload && FastenersCollection.Count > 0)
            {
                return;
            }

            Fasteners.Clear();
            FastenersCollection.Clear();
            if (!File.Exists(DbFilePath))
            {
                return;
            }

            try
            {
                byte[] rawBytes = File.ReadAllBytes(DbFilePath);
                if (rawBytes.Length > 2 && rawBytes[0] == 0x50 && rawBytes[1] == 0x4B)
                {
                    File.Delete(DbFilePath);
                    return;
                }

                string[] lines = File.ReadAllLines(DbFilePath, System.Text.Encoding.UTF8);
                foreach (string line in lines)
                {
                    if (string.IsNullOrWhiteSpace(line) || line.StartsWith("#")) continue;
                    string[] parts = line.Split('|');
                    if (parts.Length >= 2)
                    {
                        string id = parts[0].Trim();
                        string naz = parts.Length > 1 ? parts[1].Trim() : "";
                        string std = parts.Length > 2 ? parts[2].Trim() : "";
                        string vel = parts.Length > 3 ? parts[3].Trim() : "";

                        if (!string.IsNullOrWhiteSpace(id) && !string.IsNullOrWhiteSpace(naz) &&
                            Regex.IsMatch(id, @"^[a-zA-Z0-9_-]+$"))
                        {
                            var item = new ErpFastenerItem
                            {
                                ErpIdent = id,
                                Naziv = naz,
                                Standard = std,
                                Velicina = vel
                            };
                            Fasteners.Add(item);
                            FastenersCollection.Add(item);
                        }
                    }
                }
            }
            catch { }
        }

        public static void SaveData()
        {
            try
            {
                EnsureDirectory();
                List<string> lines = new List<string>();
                foreach (var item in FastenersCollection)
                {
                    if (!string.IsNullOrWhiteSpace(item.ErpIdent) || !string.IsNullOrWhiteSpace(item.Naziv))
                    {
                        lines.Add($"{item.ErpIdent}|{item.Naziv}|{item.Standard}|{item.Velicina}");
                    }
                }
                File.WriteAllLines(DbFilePath, lines, System.Text.Encoding.UTF8);

                Fasteners.Clear();
                Fasteners.AddRange(FastenersCollection);
            }
            catch { }
        }

        public static void ClearData()
        {
            Fasteners.Clear();
            FastenersCollection.Clear();
            try
            {
                if (File.Exists(DbFilePath))
                {
                    File.Delete(DbFilePath);
                }
            }
            catch { }
        }

        /// <summary>
        /// Uvoz ERP baze iz Excel (.xlsx/.xls) ili CSV/TXT fajla
        /// Podržava proizvoljan broj kolona (automatski skenira i detektuje Ident i Naziv bez obzira na šemu kolona)
        /// </summary>
        public static bool ImportFile(string filePath)
        {
            if (!File.Exists(filePath)) return false;

            string ext = Path.GetExtension(filePath).ToLowerInvariant();
            if (ext == ".xlsx" || ext == ".xls")
            {
                return ImportFromExcel(filePath);
            }
            else
            {
                return ImportFromCsv(filePath);
            }
        }

        private static bool ImportFromExcel(string excelPath)
        {
            dynamic excelApp = null;
            dynamic workbook = null;
            try
            {
                Type excelType = Type.GetTypeFromProgID("Excel.Application");
                if (excelType == null) return false;

                excelApp = Activator.CreateInstance(excelType);
                excelApp.Visible = false;
                excelApp.DisplayAlerts = false;

                workbook = excelApp.Workbooks.Open(excelPath);
                dynamic sheet = workbook.Sheets[1];

                int lastRow = sheet.Cells[sheet.Rows.Count, 1].End(-4162).Row;
                if (lastRow < 2)
                {
                    lastRow = sheet.UsedRange.Rows.Count;
                }

                int lastCol = sheet.UsedRange.Columns.Count;
                if (lastCol < 2) lastCol = 15;

                int colIdent = -1;
                int colNaziv = -1;

                for (int c = 1; c <= lastCol; c++)
                {
                    string hText = Convert.ToString(sheet.Cells[1, c].Value)?.Trim()?.ToLowerInvariant() ?? "";
                    if (string.IsNullOrWhiteSpace(hText)) continue;

                    if (hText.Contains("ident") || hText.Contains("sifra") || hText.Contains("šifra") || hText.Contains("erp") || hText == "id")
                    {
                        colIdent = c;
                    }
                    else if (hText.Contains("naziv") || hText.Contains("artikal") || hText.Contains("opis") || hText.Contains("standard") || hText.Contains("name"))
                    {
                        colNaziv = c;
                    }
                }

                List<ErpFastenerItem> newItems = new List<ErpFastenerItem>();

                for (int r = 2; r <= lastRow; r++)
                {
                    string identVal = "";
                    string nazivVal = "";

                    if (colIdent > 0 && colNaziv > 0)
                    {
                        identVal = Convert.ToString(sheet.Cells[r, colIdent].Value)?.Trim() ?? "";
                        nazivVal = Convert.ToString(sheet.Cells[r, colNaziv].Value)?.Trim() ?? "";
                    }

                    if (string.IsNullOrWhiteSpace(identVal) || string.IsNullOrWhiteSpace(nazivVal))
                    {
                        for (int c = 1; c <= lastCol; c++)
                        {
                            string cellVal = Convert.ToString(sheet.Cells[r, c].Value)?.Trim() ?? "";
                            if (string.IsNullOrWhiteSpace(cellVal)) continue;

                            if (string.IsNullOrWhiteSpace(identVal) && Regex.IsMatch(cellVal, @"^\d{3,12}$"))
                            {
                                identVal = cellVal;
                            }
                            else if (string.IsNullOrWhiteSpace(nazivVal) && cellVal.Length >= 3 && !Regex.IsMatch(cellVal, @"^\d+$"))
                            {
                                nazivVal = cellVal;
                            }
                        }
                    }

                    if (!string.IsNullOrWhiteSpace(identVal) && !string.IsNullOrWhiteSpace(nazivVal))
                    {
                        newItems.Add(new ErpFastenerItem
                        {
                            ErpIdent = identVal,
                            Naziv = nazivVal,
                            Standard = ExtractStandard(nazivVal),
                            Velicina = ExtractSize(nazivVal)
                        });
                    }
                }

                workbook.Close(false);
                excelApp.Quit();

                workbook = null;
                excelApp = null;

                if (newItems.Count > 0)
                {
                    foreach (var item in newItems)
                    {
                        if (!FastenersCollection.Any(f => f.ErpIdent.Equals(item.ErpIdent, StringComparison.OrdinalIgnoreCase) &&
                                                f.Naziv.Equals(item.Naziv, StringComparison.OrdinalIgnoreCase)))
                        {
                            FastenersCollection.Add(item);
                        }
                    }
                    SaveData();
                    return true;
                }

                return false;
            }
            catch
            {
                try { workbook?.Close(false); } catch { }
                try { excelApp?.Quit(); } catch { }
                return false;
            }
        }

        public static bool ImportFromCsv(string csvPath)
        {
            if (!File.Exists(csvPath)) return false;

            try
            {
                string[] lines = File.ReadAllLines(csvPath);
                List<ErpFastenerItem> newItems = new List<ErpFastenerItem>();

                foreach (string rawLine in lines)
                {
                    string line = rawLine.Trim();
                    if (string.IsNullOrWhiteSpace(line) || line.StartsWith("#") || line.ToLower().StartsWith("r.br")) continue;

                    string[] parts = line.Split(new char[] { ';', ',', '\t', '|' });
                    if (parts.Length >= 2)
                    {
                        string identVal = "";
                        string nazivVal = "";

                        foreach (string p in parts)
                        {
                            string item = p.Trim();
                            if (string.IsNullOrWhiteSpace(item)) continue;

                            if (string.IsNullOrWhiteSpace(identVal) && Regex.IsMatch(item, @"^\d{3,12}$"))
                            {
                                identVal = item;
                            }
                            else if (string.IsNullOrWhiteSpace(nazivVal) && item.Length >= 3 && !Regex.IsMatch(item, @"^\d+$"))
                            {
                                nazivVal = item;
                            }
                        }

                        if (!string.IsNullOrWhiteSpace(identVal) && !string.IsNullOrWhiteSpace(nazivVal))
                        {
                            newItems.Add(new ErpFastenerItem
                            {
                                ErpIdent = identVal,
                                Naziv = nazivVal,
                                Standard = ExtractStandard(nazivVal),
                                Velicina = ExtractSize(nazivVal)
                            });
                        }
                    }
                }

                if (newItems.Count > 0)
                {
                    foreach (var item in newItems)
                    {
                        if (!FastenersCollection.Any(f => f.ErpIdent.Equals(item.ErpIdent, StringComparison.OrdinalIgnoreCase) &&
                                                f.Naziv.Equals(item.Naziv, StringComparison.OrdinalIgnoreCase)))
                        {
                            FastenersCollection.Add(item);
                        }
                    }
                    SaveData();
                    return true;
                }

                return false;
            }
            catch
            {
                return false;
            }
        }

        private class CachedFastener
        {
            public ErpFastenerItem Item;
            public string NormNaziv;
            public string NormStandard;
            public string NormVelicina;
            public string[] Tokens;
        }

        private static readonly List<CachedFastener> CachedFasteners = new List<CachedFastener>();

        private static void RebuildCache()
        {
            CachedFasteners.Clear();
            foreach (var item in FastenersCollection)
            {
                if (item == null) continue;
                string normNaziv = NormalizeToken(item.Naziv ?? "");
                string normStd = NormalizeToken(item.Standard ?? "");
                string normVel = NormalizeToken(item.Velicina ?? "");
                string[] tokens = !string.IsNullOrWhiteSpace(item.Naziv)
                    ? item.Naziv.Split(new char[] { '_', '-', ' ' }, StringSplitOptions.RemoveEmptyEntries)
                    : new string[0];

                for (int i = 0; i < tokens.Length; i++)
                {
                    tokens[i] = tokens[i].ToUpperInvariant();
                }

                CachedFasteners.Add(new CachedFastener
                {
                    Item = item,
                    NormNaziv = normNaziv,
                    NormStandard = normStd,
                    NormVelicina = normVel,
                    Tokens = tokens
                });
            }
        }

        private static string ExtractStandard(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return "";
            var match = Regex.Match(text, @"(DIN\s*\d+|ISO\s*\d+)", RegexOptions.IgnoreCase);
            return match.Success ? match.Value.ToUpperInvariant() : text;
        }

        private static string ExtractSize(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return "";
            var match = Regex.Match(text, @"(M\d+(\.\d+)?(\s*[xX×]\s*\d+(\.\d+)?)?)", RegexOptions.IgnoreCase);
            return match.Success ? match.Value.ToUpperInvariant() : "";
        }

        public static bool IsFastener(string partName, string configName = "", string description = "")
        {
            string fullText = $"{partName} {configName} {description}".ToUpperInvariant();

            string[] keywords = new string[] {
                "DIN", "ISO", "BOLT", "SCHRAUBE", "VIJAK", "MATICA", "PODLOŠKA", "PODLOSKA",
                "MUTTER", "WASHER", "UNTERLEGSCHEIBE", "SCREW", "THREAD", "PEM"
            };

            foreach (string kw in keywords)
            {
                if (fullText.Contains(kw)) return true;
            }

            if (Regex.IsMatch(fullText, @"\bDIN\s*\d+") || Regex.IsMatch(fullText, @"\bISO\s*\d+") || Regex.IsMatch(fullText, @"\bM\d+X\d+\b"))
            {
                return true;
            }

            return false;
        }

        public static string FindErpIdent(string partName, string configName = "", string description = "")
        {
            if (FastenersCollection.Count == 0) return "";

            if (CachedFasteners.Count != FastenersCollection.Count)
            {
                RebuildCache();
            }

            string rawText = $"{partName} {configName} {description}".Trim();
            string normText = NormalizeToken(rawText);
            if (string.IsNullOrWhiteSpace(normText)) return "";

            // 1. Provera tačnog ili pod-naziva
            for (int i = 0; i < CachedFasteners.Count; i++)
            {
                var cached = CachedFasteners[i];
                if (!string.IsNullOrWhiteSpace(cached.NormNaziv))
                {
                    if (normText.Contains(cached.NormNaziv) || cached.NormNaziv.Contains(normText))
                    {
                        return cached.Item.ErpIdent;
                    }
                }
            }

            // 2. Provera Standard + Veličina
            for (int i = 0; i < CachedFasteners.Count; i++)
            {
                var cached = CachedFasteners[i];
                if (!string.IsNullOrWhiteSpace(cached.NormStandard) && normText.Contains(cached.NormStandard))
                {
                    if (!string.IsNullOrWhiteSpace(cached.NormVelicina) && normText.Contains(cached.NormVelicina))
                    {
                        return cached.Item.ErpIdent;
                    }
                }
            }

            // 3. Token pretraga
            for (int i = 0; i < CachedFasteners.Count; i++)
            {
                var cached = CachedFasteners[i];
                if (cached.Tokens.Length > 0)
                {
                    int matches = 0;
                    for (int t = 0; t < cached.Tokens.Length; t++)
                    {
                        string token = cached.Tokens[t];
                        if (token.Length >= 2 && normText.Contains(token))
                        {
                            matches++;
                        }
                    }

                    if (matches >= 2 && matches >= cached.Tokens.Length - 1)
                    {
                        return cached.Item.ErpIdent;
                    }
                }
            }

            return "";
        }

        private static string NormalizeToken(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return "";
            return text.Replace("_", " ").Replace("-", " ").Replace("  ", " ").Trim().ToUpperInvariant();
        }

        public static int GetFastenerCount()
        {
            return FastenersCollection.Count;
        }
    }
}
