using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.Linq;

namespace FASTIMATE
{
    public class KFactorRuleItem
    {
        public string Material { get; set; }  // "Čelik", "Inox", "Aluminijum"
        public string ThicknessStr { get; set; }
        public string KFactorStr { get; set; }

        public double Thickness
        {
            get
            {
                string s = (ThicknessStr ?? "").Trim().Replace(',', '.');
                return double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out double val) ? val : 0.0;
            }
        }

        public double KFactor
        {
            get
            {
                string s = (KFactorStr ?? "").Trim().Replace(',', '.');
                return double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out double val) ? val : 0.0;
            }
        }
    }

    public static class KFactorManager
    {
        private static readonly Dictionary<double, double> SteelKFactors = new Dictionary<double, double>();
        private static readonly Dictionary<double, double> InoxKFactors = new Dictionary<double, double>();
        private static readonly Dictionary<double, double> AluminiumKFactors = new Dictionary<double, double>();

        private static readonly ObservableCollection<KFactorRuleItem> KFactorRulesCollection = new ObservableCollection<KFactorRuleItem>();

        private static readonly string ConfigFilePath = Path.Combine(
            AppPaths.GetDbDirectory(),
            "k_factor_rules.txt"
        );

        static KFactorManager()
        {
            EnsureDirectory();
            LoadData();
        }

        public static void LoadData(bool forceReload = false)
        {
            if (!forceReload && KFactorRulesCollection.Count > 0)
            {
                return;
            }
            LoadDefaultRules();
            LoadFromFile();
        }

        private static void EnsureDirectory()
        {
            string dir = Path.GetDirectoryName(ConfigFilePath);
            if (!Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }
        }

        private static void LoadDefaultRules()
        {
            SteelKFactors.Clear();
            SteelKFactors[1.0] = 0.115;
            SteelKFactors[1.5] = 0.22;
            SteelKFactors[2.0] = 0.3;
            SteelKFactors[3.0] = 0.25;
            SteelKFactors[4.0] = 0.215;
            SteelKFactors[5.0] = 0.171;
            SteelKFactors[6.0] = 0.143;

            InoxKFactors.Clear();
            InoxKFactors[1.0] = 0.221;
            InoxKFactors[1.5] = 0.133;
            InoxKFactors[2.0] = 0.197;
            InoxKFactors[3.0] = 0.199;
            InoxKFactors[4.0] = 0.0005;
            InoxKFactors[5.0] = 0.0985;

            AluminiumKFactors.Clear();
            AluminiumKFactors[1.0] = 0.22;
            AluminiumKFactors[1.5] = 0.25;
            AluminiumKFactors[2.0] = 0.30;
            AluminiumKFactors[3.0] = 0.35;
            AluminiumKFactors[4.0] = 0.40;
            AluminiumKFactors[5.0] = 0.45;
        }

        public static ObservableCollection<KFactorRuleItem> GetAllRules()
        {
            SyncCollectionFromDicts();
            return KFactorRulesCollection;
        }

        private static void SyncCollectionFromDicts()
        {
            KFactorRulesCollection.Clear();
            foreach (var kvp in SteelKFactors.OrderBy(k => k.Key))
            {
                KFactorRulesCollection.Add(new KFactorRuleItem
                {
                    Material = "Čelik",
                    ThicknessStr = kvp.Key.ToString("G", CultureInfo.InvariantCulture),
                    KFactorStr = kvp.Value.ToString("G", CultureInfo.InvariantCulture)
                });
            }

            foreach (var kvp in InoxKFactors.OrderBy(k => k.Key))
            {
                KFactorRulesCollection.Add(new KFactorRuleItem
                {
                    Material = "Inox",
                    ThicknessStr = kvp.Key.ToString("G", CultureInfo.InvariantCulture),
                    KFactorStr = kvp.Value.ToString("G", CultureInfo.InvariantCulture)
                });
            }

            foreach (var kvp in AluminiumKFactors.OrderBy(k => k.Key))
            {
                KFactorRulesCollection.Add(new KFactorRuleItem
                {
                    Material = "Aluminijum",
                    ThicknessStr = kvp.Key.ToString("G", CultureInfo.InvariantCulture),
                    KFactorStr = kvp.Value.ToString("G", CultureInfo.InvariantCulture)
                });
            }
        }

        public static void LoadFromFile()
        {
            if (!File.Exists(ConfigFilePath))
            {
                SaveToFile();
                return;
            }

            try
            {
                string[] lines = File.ReadAllLines(ConfigFilePath, System.Text.Encoding.UTF8);
                string currentMaterial = "";

                SteelKFactors.Clear();
                InoxKFactors.Clear();
                AluminiumKFactors.Clear();

                foreach (string rawLine in lines)
                {
                    string line = rawLine.Trim();
                    if (string.IsNullOrEmpty(line) || line.StartsWith("#")) continue;

                    if (line.StartsWith("[") && line.EndsWith("]"))
                    {
                        currentMaterial = line.Substring(1, line.Length - 2).ToLowerInvariant();
                        continue;
                    }

                    string[] parts = line.Split('=');
                    if (parts.Length == 2)
                    {
                        string p0 = parts[0].Trim().Replace(',', '.');
                        string p1 = parts[1].Trim().Replace(',', '.');

                        if (double.TryParse(p0, NumberStyles.Any, CultureInfo.InvariantCulture, out double thickness) &&
                            double.TryParse(p1, NumberStyles.Any, CultureInfo.InvariantCulture, out double kFactor))
                        {
                            if (currentMaterial.Contains("inox") || currentMaterial.Contains("rostfrei"))
                            {
                                InoxKFactors[thickness] = kFactor;
                            }
                            else if (currentMaterial.Contains("alu"))
                            {
                                AluminiumKFactors[thickness] = kFactor;
                            }
                            else
                            {
                                SteelKFactors[thickness] = kFactor;
                            }
                        }
                    }
                }

                SyncCollectionFromDicts();
            }
            catch { }
        }

        public static bool AddNewRule(string material, double thickness, double kFactor, out string error)
        {
            error = "";
            string matNorm = (material ?? "").Trim();
            if (string.IsNullOrWhiteSpace(matNorm))
            {
                error = "Odaberite važeći materijal (Čelik, Inox ili Aluminijum).";
                return false;
            }

            if (matNorm.Equals("inox", StringComparison.OrdinalIgnoreCase))
            {
                matNorm = "Inox";
            }
            else if (matNorm.IndexOf("alu", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                matNorm = "Aluminijum";
            }
            else
            {
                matNorm = "Čelik";
            }

            Dictionary<double, double> targetDict = SteelKFactors;
            if (matNorm == "Inox") targetDict = InoxKFactors;
            else if (matNorm == "Aluminijum") targetDict = AluminiumKFactors;

            foreach (var key in targetDict.Keys)
            {
                if (Math.Abs(key - thickness) < 0.01)
                {
                    error = $"K-Faktor za materijal '{matNorm}' i debljinu {thickness} mm već postoji u bazi!";
                    return false;
                }
            }

            targetDict[thickness] = kFactor;
            SaveToFile();
            SyncCollectionFromDicts();
            return true;
        }

        public static void SaveAllRules()
        {
            SteelKFactors.Clear();
            InoxKFactors.Clear();
            AluminiumKFactors.Clear();

            foreach (var item in KFactorRulesCollection)
            {
                string m = (item.Material ?? "").Trim().ToLowerInvariant();
                double th = item.Thickness;
                double kf = item.KFactor;

                if (th <= 0) continue;

                if (m.Contains("inox"))
                {
                    InoxKFactors[th] = kf;
                }
                else if (m.Contains("alu"))
                {
                    AluminiumKFactors[th] = kf;
                }
                else
                {
                    SteelKFactors[th] = kf;
                }
            }

            SaveToFile();
            SyncCollectionFromDicts();
        }

        public static void DeleteRule(KFactorRuleItem rule)
        {
            if (rule == null) return;
            string m = (rule.Material ?? "").Trim().ToLowerInvariant();
            double th = rule.Thickness;

            if (m.Contains("inox"))
            {
                InoxKFactors.Remove(th);
            }
            else if (m.Contains("alu"))
            {
                AluminiumKFactors.Remove(th);
            }
            else
            {
                SteelKFactors.Remove(th);
            }

            SaveToFile();
            SyncCollectionFromDicts();
        }

        public static void SaveToFile()
        {
            try
            {
                EnsureDirectory();
                List<string> lines = new List<string>
                {
                    "# K-Factor configuration rules for FASTIMATE",
                    "[STEEL]"
                };

                foreach (var kvp in SteelKFactors.OrderBy(k => k.Key))
                {
                    lines.Add($"{kvp.Key.ToString(CultureInfo.InvariantCulture)}={kvp.Value.ToString(CultureInfo.InvariantCulture)}");
                }

                lines.Add("\n[INOX]");
                foreach (var kvp in InoxKFactors.OrderBy(k => k.Key))
                {
                    lines.Add($"{kvp.Key.ToString(CultureInfo.InvariantCulture)}={kvp.Value.ToString(CultureInfo.InvariantCulture)}");
                }

                lines.Add("\n[ALUMINIUM]");
                foreach (var kvp in AluminiumKFactors.OrderBy(k => k.Key))
                {
                    lines.Add($"{kvp.Key.ToString(CultureInfo.InvariantCulture)}={kvp.Value.ToString(CultureInfo.InvariantCulture)}");
                }

                File.WriteAllLines(ConfigFilePath, lines, System.Text.Encoding.UTF8);
            }
            catch { }
        }

        public static bool IsKFactorValid(string materialName, double thickness, double actualKFactor)
        {
            return IsValidKFactor(materialName, thickness, actualKFactor, out _);
        }

        public static bool IsValidKFactor(string materialName, double thickness, double actualKFactor, out double expectedKFactor)
        {
            expectedKFactor = GetKFactor(materialName, thickness);
            if (thickness <= 0) return true;

            if (expectedKFactor > 0)
            {
                return Math.Abs(expectedKFactor - actualKFactor) < 0.02;
            }

            return true;
        }

        public static double GetKFactor(string materialName, double thickness)
        {
            if (thickness <= 0) return 0.25;
            string mat = (materialName ?? "").ToLowerInvariant();
            Dictionary<double, double> targetDict = SteelKFactors;

            if (mat.Contains("inox") || mat.Contains("rostfrei") || mat.Contains("1.4301") || mat.Contains("1.4404"))
            {
                targetDict = InoxKFactors;
            }
            else if (mat.Contains("alu") || mat.Contains("alumin"))
            {
                targetDict = AluminiumKFactors;
            }

            if (targetDict.Count == 0) return 0.25;

            double closestDiff = double.MaxValue;
            double closestKFactor = 0.25;

            foreach (var kvp in targetDict)
            {
                double diff = Math.Abs(kvp.Key - thickness);
                if (diff < closestDiff)
                {
                    closestDiff = diff;
                    closestKFactor = kvp.Value;
                }
            }

            return closestKFactor;
        }
    }
}
