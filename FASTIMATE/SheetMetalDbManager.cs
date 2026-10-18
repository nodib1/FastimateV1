using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace FASTIMATE
{
    public class SheetMetalRawItem
    {
        public string ErpIdent { get; set; }
        public string Material { get; set; }
        public string ThicknessStr { get; set; }
        public string Description { get; set; }

        public double Thickness
        {
            get
            {
                string s = (ThicknessStr ?? "").Trim().Replace(',', '.');
                return double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out double val) ? val : 0.0;
            }
        }
    }

    public static class SheetMetalDbManager
    {
        private static readonly List<SheetMetalRawItem> Items = new List<SheetMetalRawItem>();
        private static readonly ObservableCollection<SheetMetalRawItem> ItemsCollection = new ObservableCollection<SheetMetalRawItem>();

        private static readonly string DbFilePath = Path.Combine(
            AppPaths.GetDbDirectory(),
            "sheet_metal_raw_materials.txt"
        );

        static SheetMetalDbManager()
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

        public static ObservableCollection<SheetMetalRawItem> GetAllItems()
        {
            return ItemsCollection;
        }

        public static void LoadData(bool forceReload = false)
        {
            if (!forceReload && ItemsCollection.Count > 0)
            {
                return;
            }

            Items.Clear();
            ItemsCollection.Clear();
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
                foreach (string rawLine in lines)
                {
                    string line = rawLine.Trim();
                    if (string.IsNullOrWhiteSpace(line) || line.StartsWith("#")) continue;

                    string[] parts = line.Split('|');
                    if (parts.Length >= 2)
                    {
                        string id = parts[0].Trim();
                        string mat = parts.Length > 1 ? parts[1].Trim() : "";
                        string thStr = parts.Length > 2 ? parts[2].Trim() : "0";
                        string desc = parts.Length > 3 ? parts[3].Trim() : "";

                        if (!string.IsNullOrWhiteSpace(id))
                        {
                            var item = new SheetMetalRawItem
                            {
                                ErpIdent = id,
                                Material = mat,
                                ThicknessStr = thStr,
                                Description = desc
                            };
                            Items.Add(item);
                            ItemsCollection.Add(item);
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
                foreach (var item in ItemsCollection)
                {
                    if (!string.IsNullOrWhiteSpace(item.ErpIdent))
                    {
                        lines.Add($"{item.ErpIdent}|{item.Material}|{item.ThicknessStr}|{item.Description}");
                    }
                }
                File.WriteAllLines(DbFilePath, lines, System.Text.Encoding.UTF8);

                Items.Clear();
                Items.AddRange(ItemsCollection);
            }
            catch { }
        }

        public static void ClearData()
        {
            Items.Clear();
            ItemsCollection.Clear();
            try
            {
                if (File.Exists(DbFilePath))
                {
                    File.Delete(DbFilePath);
                }
            }
            catch { }
        }

        public static bool AddNewItem(string id, string mat, double thickness, string desc, out string error)
        {
            error = "";
            id = (id ?? "").Trim();
            mat = (mat ?? "").Trim();
            desc = (desc ?? "").Trim();

            if (string.IsNullOrWhiteSpace(id))
            {
                error = "Unesite važeći ERP Ident sirovine limova.";
                return false;
            }

            // Uslov: ERP Ident mora biti brojčana vrijednost i imati TAČNO 6 cifara
            if (!Regex.IsMatch(id, @"^\d{6}$"))
            {
                error = "ERP Ident mora biti brojčana vrijednost sa tačno 6 cifara (npr. 900123).";
                return false;
            }

            if (ItemsCollection.Any(i => i.ErpIdent.Equals(id, StringComparison.OrdinalIgnoreCase)))
            {
                error = $"Sirovina limova sa ERP Identom '{id}' već postoji u bazi!";
                return false;
            }

            var newItem = new SheetMetalRawItem
            {
                ErpIdent = id,
                Material = mat,
                ThicknessStr = thickness.ToString("G", CultureInfo.InvariantCulture),
                Description = desc
            };

            ItemsCollection.Add(newItem);
            SaveData();
            return true;
        }

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
                if (lastRow < 2) lastRow = sheet.UsedRange.Rows.Count;

                int lastCol = sheet.UsedRange.Columns.Count;
                if (lastCol < 2) lastCol = 15;

                int colIdent = -1, colMat = -1, colTh = -1, colDesc = -1;

                for (int c = 1; c <= lastCol; c++)
                {
                    string hText = Convert.ToString(sheet.Cells[1, c].Value)?.Trim()?.ToLowerInvariant() ?? "";
                    if (string.IsNullOrWhiteSpace(hText)) continue;

                    if (hText.Contains("ident") || hText.Contains("sifra") || hText.Contains("šifra") || hText == "id")
                        colIdent = c;
                    else if (hText.Contains("materijal") || hText.Contains("material") || hText.Contains("vrsta"))
                        colMat = c;
                    else if (hText.Contains("debljina") || hText.Contains("thick") || hText.Contains("deb"))
                        colTh = c;
                    else if (hText.Contains("opis") || hText.Contains("naziv") || hText.Contains("dimenzija") || hText.Contains("format") || hText.Contains("standard"))
                        colDesc = c;
                }

                List<SheetMetalRawItem> newItems = new List<SheetMetalRawItem>();

                for (int r = 2; r <= lastRow; r++)
                {
                    string identVal = colIdent > 0 ? Convert.ToString(sheet.Cells[r, colIdent].Value)?.Trim() ?? "" : "";
                    string matVal = colMat > 0 ? Convert.ToString(sheet.Cells[r, colMat].Value)?.Trim() ?? "" : "";
                    string thVal = colTh > 0 ? Convert.ToString(sheet.Cells[r, colTh].Value)?.Trim() ?? "" : "";
                    string descVal = colDesc > 0 ? Convert.ToString(sheet.Cells[r, colDesc].Value)?.Trim() ?? "" : "";

                    if (string.IsNullOrWhiteSpace(identVal))
                    {
                        for (int c = 1; c <= lastCol; c++)
                        {
                            string cellVal = Convert.ToString(sheet.Cells[r, c].Value)?.Trim() ?? "";
                            if (!string.IsNullOrWhiteSpace(cellVal) && Regex.IsMatch(cellVal, @"^\d{6}$"))
                            {
                                identVal = cellVal;
                                break;
                            }
                        }
                    }

                    if (!string.IsNullOrWhiteSpace(identVal) && Regex.IsMatch(identVal, @"^\d{6}$"))
                    {
                        newItems.Add(new SheetMetalRawItem
                        {
                            ErpIdent = identVal,
                            Material = !string.IsNullOrWhiteSpace(matVal) ? matVal : "Crni lim",
                            ThicknessStr = !string.IsNullOrWhiteSpace(thVal) ? thVal.Replace(',', '.') : "1.0",
                            Description = descVal
                        });
                    }
                }

                workbook.Close(false);
                excelApp.Quit();

                if (newItems.Count > 0)
                {
                    foreach (var item in newItems)
                    {
                        if (!ItemsCollection.Any(i => i.ErpIdent.Equals(item.ErpIdent, StringComparison.OrdinalIgnoreCase)))
                        {
                            ItemsCollection.Add(item);
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

        private static bool ImportFromCsv(string csvPath)
        {
            if (!File.Exists(csvPath)) return false;

            try
            {
                string[] lines = File.ReadAllLines(csvPath);
                List<SheetMetalRawItem> newItems = new List<SheetMetalRawItem>();

                foreach (string rawLine in lines)
                {
                    string line = rawLine.Trim();
                    if (string.IsNullOrWhiteSpace(line) || line.StartsWith("#") || line.ToLower().StartsWith("r.br")) continue;

                    string[] parts = line.Split(new char[] { ';', ',', '\t', '|' });
                    if (parts.Length >= 2)
                    {
                        string id = parts[0].Trim();
                        string mat = parts.Length > 1 ? parts[1].Trim() : "Crni lim";
                        string thStr = parts.Length > 2 ? parts[2].Trim().Replace(',', '.') : "1.0";
                        string desc = parts.Length > 3 ? parts[3].Trim() : "";

                        if (!string.IsNullOrWhiteSpace(id) && Regex.IsMatch(id, @"^\d{6}$"))
                        {
                            newItems.Add(new SheetMetalRawItem
                            {
                                ErpIdent = id,
                                Material = mat,
                                ThicknessStr = thStr,
                                Description = desc
                            });
                        }
                    }
                }

                if (newItems.Count > 0)
                {
                    foreach (var item in newItems)
                    {
                        if (!ItemsCollection.Any(i => i.ErpIdent.Equals(item.ErpIdent, StringComparison.OrdinalIgnoreCase)))
                        {
                            ItemsCollection.Add(item);
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

        public static int GetItemCount()
        {
            return ItemsCollection.Count;
        }

        /// <summary>
        /// Pronalazi tačan ERP Ident sirovine limova na osnovu materijala i debljine.
        /// Određuje kategoriju (Crni lim, Inox, Pocinčani, Aluminijum).
        /// Za Crni lim: debljine < 4mm su DC01, a >= 4mm su DD11/S235.
        /// </summary>
        public static string FindRawMaterialIdent(string materialName, double thickness)
        {
            if (ItemsCollection.Count == 0) return "";
            string matNorm = (materialName ?? "").ToLowerInvariant();

            // 1. Kategorizacija materijala
            string category = "Crni lim";
            if (matNorm.Contains("inox") || matNorm.Contains("1.4301") || matNorm.Contains("1.4404") || matNorm.Contains("rostfrei"))
            {
                category = "Inox";
            }
            else if (matNorm.Contains("pocinčan") || matNorm.Contains("pocincan") || matNorm.Contains("dx51d") || matNorm.Contains("galvaniz"))
            {
                category = "Pocinčani";
            }
            else if (matNorm.Contains("alu") || matNorm.Contains("en-aw") || matNorm.Contains("almg"))
            {
                category = "Aluminijum";
            }

            // Target standard za crni lim po debljini: < 4mm -> DC01, >= 4mm -> DD11
            string targetStandard = "";
            if (category == "Crni lim")
            {
                targetStandard = thickness < 4.0 ? "dc01" : "dd11";
            }

            // 2. Pretraga po kategoriji i debljini + specifičnom standardu u Opisu (DC01 / DD11)
            foreach (var item in ItemsCollection)
            {
                if (Math.Abs(item.Thickness - thickness) < 0.05)
                {
                    string itemMat = (item.Material ?? "").ToLowerInvariant();
                    string itemDesc = (item.Description ?? "").ToLowerInvariant();

                    if (itemMat.Contains(category.ToLowerInvariant()) || category.ToLowerInvariant().Contains(itemMat))
                    {
                        if (!string.IsNullOrWhiteSpace(targetStandard) && itemDesc.Contains(targetStandard))
                        {
                            return item.ErpIdent;
                        }
                    }
                }
            }

            // 3. Pretraga po kategoriji i debljini
            foreach (var item in ItemsCollection)
            {
                if (Math.Abs(item.Thickness - thickness) < 0.05)
                {
                    string itemMat = (item.Material ?? "").ToLowerInvariant();
                    if (itemMat.Contains(category.ToLowerInvariant()) || category.ToLowerInvariant().Contains(itemMat))
                    {
                        return item.ErpIdent;
                    }
                }
            }

            // 4. Fallback na najbližu debljinu
            foreach (var item in ItemsCollection)
            {
                if (Math.Abs(item.Thickness - thickness) < 0.05)
                {
                    return item.ErpIdent;
                }
            }

            return "";
        }
    }
}
