using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace FASTIMATE
{
    /// <summary>
    /// Proračun tehnologije: čita podatke iz otvorenog SolidWorks sklopa i upisuje ih u Excel.
    ///
    /// Zamjenjuje raniji Macro_Tehnologija.swp i naknadnu doradu (FixMainAssemblyData) jednim prolazom.
    /// Aplikacija radi izvan SolidWorks procesa, pa je svaki API poziv skup (COM između procesa).
    /// Zato se: svaki model obrađuje samo jednom, propertyji čitaju odjednom (GetAll3),
    /// stablo feature-a prolazi jednom, a Excel se piše u blokovima na kraju.
    /// </summary>
    public class TehnologijaProcessor
    {
        private const int FirstRow = 10;
        private const int LastScanRow = 1000;

        // Kolone u Excel šablonu
        private const int ColNaziv = 1;                 // A
        private const int ColIdent = 2;                 // B
        private const int ColOsnovniNaziv = 3;          // C
        private const int ColKolicina = 4;              // D
        private const int ColDebljina = 5;              // E
        private const int ColRazvijenaDuzina = 6;       // F
        private const int ColRazvijenaSirina = 7;       // G
        private const int ColFinish = 10;               // J
        private const int ColNavoji = 53;
        private const int ColUpustanja = 58;
        private const int ColPemUSklopu = 64;
        private const int ColSavijanja = 67;
        private const int ColNavarneCelik = 70;
        private const int ColZavarKomadaCelik = 73;
        private const int ColZavarCelik = 74;
        private const int ColNavarneInox = 83;
        private const int ColZavarKomadaInox = 86;
        private const int ColZavarInox = 87;
        private const int ColPlastifikacija = 104;      // CZ – formula, samo se čita
        private const int ColSchaumdicht = 107;
        private const int ColVijci = 114;               // DJ
        private const int ColMatice = 115;              // DK
        private const int ColBrojKomponenti = 120;      // DP
        private const int ColPovrsina = 121;
        private const int ColVolumen = 124;
        private const int ColUpravljackiCentar = 132;
        private const int ColMaterijal = 134;           // ED
        private const int ColPovrsinskaZastita = 136;   // EF
        private const int ColKFaktor = 138;             // EH

        private const int KFaktorOdstupanjeBoja = 13815295; // blago crvena (#FFC7CE)

        private static readonly string[] CutListKKeys = {
            "Bend Allowance", "Bend Allowance Standard", "K-Factor", "K Factor", "K_Faktor", "KFactor",
            "Bend Deduction", "Dodatak za savijanje", "Bend_Allowance"
        };

        private static readonly string[] ModelKKeys = {
            "Bend Allowance", "Bend Allowance Standard", "K-Factor", "K Factor", "K_Faktor", "Bend Deduction", "Dodatak za savijanje"
        };

        [DllImport("user32.dll")]
        private static extern bool EnableWindow(IntPtr hWnd, bool bEnable);

        private readonly SldWorks swApp;
        private readonly ModelDoc2 swModel;

        // Količina po nazivu (A kolona) – isti ključ kao u makrou (osjetljiv na velika/mala slova)
        private readonly Dictionary<string, int> dictQuantity = new Dictionary<string, int>(StringComparer.Ordinal);
        // Putanja fajla -> naziv reda (null = model nema GS_Ident i preskače se)
        private readonly Dictionary<string, string> dictPathToName = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        private readonly List<RowData> rows = new List<RowData>();
        private readonly List<string> warnings = new List<string>();
        private RowData mainRow;
        private int nextRow = FirstRow;

        private TehnologijaProcessor(SldWorks app, ModelDoc2 model)
        {
            swApp = app;
            swModel = model;
        }

        public static bool RunProcess(SldWorks swApp, ModelDoc2 swModel, string excelFile, string ident, string logovaniKorisnik, string datumUpisa)
        {
            try
            {
                if (swApp == null || swModel == null)
                {
                    MessageBox.Show("Nema otvorenog SolidWorks dokumenta.", "Greška", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return false;
                }

                if (IsFileLocked(excelFile))
                {
                    MessageBox.Show($"Excel fajl '{Path.GetFileName(excelFile)}' je trenutno otvoren.\n\nZatvorite ga i pokrenite obradu ponovo.", "Fajl je otvoren", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return false;
                }

                return new TehnologijaProcessor(swApp, swModel).Execute(excelFile, ident, logovaniKorisnik, datumUpisa);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Greška pri obradi Tehnologije: {ex.Message}", "Greška", MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }
        }

        private bool Execute(string excelFile, string ident, string korisnik, string datum)
        {
            Stopwatch stopwatch = Stopwatch.StartNew();
            System.Windows.Forms.Form procesBox = ProcesUTokuProzor.Prikazi();
            try
            {
                return ExecuteCore(excelFile, ident, korisnik, datum, stopwatch, procesBox);
            }
            finally
            {
                ProcesUTokuProzor.Zatvori(procesBox);
            }
        }

        private bool ExecuteCore(string excelFile, string ident, string korisnik, string datum, Stopwatch stopwatch, System.Windows.Forms.Form procesBox)
        {
            try { KFactorManager.LoadData(); } catch { }

            IntPtr swHwnd = GetSolidWorksWindow();
            if (swHwnd != IntPtr.Zero) { try { EnableWindow(swHwnd, false); } catch { } }

            // CommandInProgress znatno ubrzava veliki broj API poziva iz vanjske aplikacije
            try { swApp.CommandInProgress = true; } catch { }
            ModelView view = null;
            try
            {
                view = swModel.ActiveView as ModelView;
                if (view != null) view.EnableGraphicsUpdate = false;
            }
            catch { view = null; }

            MainAssemblyData main;
            try
            {
                CollectFromSolidWorks();
                main = ReadMainAssemblyData();
            }
            finally
            {
                try { if (view != null) view.EnableGraphicsUpdate = true; } catch { }
                try { swApp.CommandInProgress = false; } catch { }
                if (swHwnd != IntPtr.Zero) { try { EnableWindow(swHwnd, true); } catch { } }
            }

            WriteExcel(excelFile, ident, korisnik, datum, main);

            stopwatch.Stop();

            string message = $"Obrada završena uspješno.\n\nBroj pozicija: {rows.Count}\nTrajanje: {stopwatch.Elapsed.TotalSeconds.ToString("0.0", CultureInfo.InvariantCulture)} s";
            if (warnings.Count > 0)
            {
                message += "\n\nNeke pozicije nisu u potpunosti obrađene:\n" + string.Join("\n", warnings.GetRange(0, Math.Min(warnings.Count, 10)));
                if (warnings.Count > 10) message += $"\n... i još {warnings.Count - 10}";
            }
            ProcesUTokuProzor.Zatvori(procesBox);
            MessageBox.Show(message, "Uspjeh", MessageBoxButton.OK, warnings.Count > 0 ? MessageBoxImage.Warning : MessageBoxImage.Information);
            return true;
        }

        #region SolidWorks – prikupljanje podataka

        private void CollectFromSolidWorks()
        {
            Configuration conf = swModel.ConfigurationManager?.ActiveConfiguration;
            Component2 root = conf?.GetRootComponent3(true);

            if (root != null)
            {
                ProcessComponent(root, true);
            }
            else
            {
                // Aktivni dokument je dio (part), a ne sklop
                ProcessModel(swModel, null, true);
            }
        }

        /// <summary>
        /// Rekurzija kroz stablo kao u makrou: prvo djeca, pa sama komponenta.
        /// Glavni sklop se zato upisuje u zadnji red.
        /// </summary>
        private void ProcessComponent(Component2 comp, bool isRoot)
        {
            if (comp.GetChildren() is object[] children)
            {
                foreach (object child in children)
                {
                    if (child is Component2 childComp) ProcessComponent(childComp, false);
                }
            }

            ModelDoc2 model = isRoot ? swModel : comp.GetModelDoc2() as ModelDoc2;
            if (model == null) return; // potisnuta ili lightweight komponenta – kao u makrou

            string path = "";
            try { path = isRoot ? swModel.GetPathName() : comp.GetPathName(); } catch { }

            if (!string.IsNullOrEmpty(path) && dictPathToName.TryGetValue(path, out string knownName))
            {
                if (knownName != null) dictQuantity[knownName]++;
                return;
            }

            string finalName = ProcessModel(model, isRoot ? null : comp, isRoot);
            if (!string.IsNullOrEmpty(path)) dictPathToName[path] = finalName;
        }

        /// <summary>Vraća naziv reda ili null ako model nema GS_Ident.</summary>
        private string ProcessModel(ModelDoc2 model, Component2 comp, bool isMain)
        {
            PropSet globalProps = PropSet.Read(model.Extension.get_CustomPropertyManager(""));

            string gsIdent = globalProps.Resolved("GS_Ident");
            if (gsIdent.Trim() == "") return null;

            string gsNaziv = globalProps.Resolved("GS_Osnovni_naziv");
            string gsDodatni = globalProps.Resolved("GS_Dodatni_opis");
            string finalName = gsDodatni.Trim() != "" ? gsDodatni : gsNaziv;

            if (dictQuantity.ContainsKey(finalName))
            {
                dictQuantity[finalName]++;
                return finalName;
            }
            dictQuantity[finalName] = 1;

            RowData row = new RowData { Row = nextRow++, FinalName = finalName, IsMain = isMain };
            rows.Add(row);
            if (isMain) mainRow = row;

            row.Set(ColNaziv, finalName);
            row.Set(ColIdent, gsIdent);
            row.Set(ColOsnovniNaziv, gsNaziv);

            try
            {
                AnalyzeModel(model, comp, row, globalProps);
            }
            catch (Exception ex)
            {
                warnings.Add($"{finalName}: {ex.Message}");
            }

            return finalName;
        }

        private void AnalyzeModel(ModelDoc2 model, Component2 comp, RowData row, PropSet globalProps)
        {
            int docType = model.GetType();
            bool isPart = docType == (int)swDocumentTypes_e.swDocPART;
            bool isAssembly = docType == (int)swDocumentTypes_e.swDocASSEMBLY;

            string uc = globalProps.Resolved("GS_Upravljacki_centar").Trim();
            if (uc == "66") row.Set(ColUpravljackiCentar, "PMPR");
            else if (uc == "68") row.Set(ColUpravljackiCentar, "NMPR");

            // --- MATERIJAL ---
            string swMaterialName = "";
            string materialEval = "";
            string materialForExcel = "";
            if (isPart)
            {
                swMaterialName = GetPartMaterial(model);
                materialEval = swMaterialName;
                materialForExcel = swMaterialName;
            }
            else if (isAssembly)
            {
                materialForExcel = "--";
                materialEval = GetFirstPartMaterialInAssembly(model);
            }
            bool uvecajPerimetar = IsUvecanPerimetar(materialEval);

            // --- DJECA SKLOPA (jedno čitanje, koristi se za montažu, PEM i zavarivanje) ---
            List<string> kidNames = isAssembly ? GetTopLevelChildNames(model) : null;
            if (kidNames != null)
            {
                long mCount = 0, vCount = 0;
                foreach (string nRaw in kidNames)
                {
                    string nClean = nRaw.Replace(" ", "").Replace("_", "");
                    if (!nClean.Contains("din929") && !nClean.Contains("iso13918"))
                    {
                        if (nRaw.Contains("matica")) mCount++;
                        else if (nRaw.Contains("vijak")) vCount++;
                    }
                }
                if (vCount > 0) row.Set(ColVijci, vCount);
                if (mCount > 0) row.Set(ColMatice, mCount);
            }

            // --- JEDAN PROLAZ KROZ STABLO FEATURE-A ---
            FeatureScan scan = ScanFeatures(model, isPart);
            if (!string.IsNullOrWhiteSpace(scan.CutListMaterial))
            {
                materialEval = scan.CutListMaterial;
                materialForExcel = scan.CutListMaterial;
                if (IsUvecanPerimetar(materialEval)) uvecajPerimetar = true;
            }

            double thickness = scan.Thickness;
            row.Thickness = Math.Round(thickness, 2, MidpointRounding.AwayFromZero);
            row.Material = materialForExcel;
            row.IsAssembly = isAssembly;

            // --- SCHAUMDICHT (rezultat u metrima) ---
            if (row.FinalName.ToLowerInvariant().Contains("schaumdicht"))
            {
                Array box = null;
                if (isPart) box = (model as PartDoc)?.GetPartBox(true) as Array;
                else if (isAssembly) box = (model as AssemblyDoc)?.GetBox(0) as Array;

                if (box != null && box.Length >= 6)
                {
                    double[] d = {
                        Math.Abs(Convert.ToDouble(box.GetValue(3)) - Convert.ToDouble(box.GetValue(0))),
                        Math.Abs(Convert.ToDouble(box.GetValue(4)) - Convert.ToDouble(box.GetValue(1))),
                        Math.Abs(Convert.ToDouble(box.GetValue(5)) - Convert.ToDouble(box.GetValue(2)))
                    };
                    Array.Sort(d);
                    Array.Reverse(d);
                    row.Set(ColSchaumdicht, Math.Round((d[0] * 2) + (d[1] * 2) + d[2], 3, MidpointRounding.AwayFromZero));
                }
            }

            // --- AUTOMATSKO AŽURIRANJE "Material" PROPERTYJA U MODELU ---
            AzurirajMaterijalUProperties(model, isAssembly, swMaterialName, thickness);

            // --- OSNOVNI PODACI ---
            // Kao makro (Format "0.00"): tekst sa zarezom – formule u šablonu računaju s tim
            row.Set(ColDebljina, row.Thickness.ToString("0.00", CultureInfo.InvariantCulture).Replace('.', ','));
            row.Set(ColRazvijenaDuzina, ToCellValue(scan.FlattenLength));
            row.Set(ColRazvijenaSirina, ToCellValue(scan.FlattenWidth));
            row.Set(ColSavijanja, scan.IsSheetMetal ? ToCellValue(scan.Bends) : null);
            row.Set(ColMaterijal, materialForExcel);

            // --- PERIMETAR PO DEBLJINI ---
            if (scan.IsSheetMetal)
            {
                double finalPerimeter = uvecajPerimetar ? scan.FlattenPerimeter * 1.5 : scan.FlattenPerimeter;
                int perimeterCol = GetPerimeterColumn(thickness);
                if (perimeterCol > 0) row.Set(perimeterCol, finalPerimeter);
            }

            // --- RUPE ---
            if (scan.ThreadCount > 0) row.Set(ColNavoji, scan.ThreadCount);
            if (scan.CountersinkCount > 0) row.Set(ColUpustanja, scan.CountersinkCount);

            if (kidNames != null)
            {
                long pemCount = 0;
                foreach (string n in kidNames)
                {
                    if (n.Contains("pem") || n.Contains("fh")) pemCount++;
                }
                if (pemCount > 0) row.Set(ColPemUSklopu, pemCount);
            }

            // --- ZAVARIVANJE ---
            WriteWeldData(model, uvecajPerimetar, row, kidNames);

            // --- POVRŠINSKA ZAŠTITA, FINISH I K-FAKTOR ---
            List<PropSet> confProps = ReadConfigurationProps(model, row.IsMain ? null : comp);
            row.PovrsinskaZastita = LookupProp(confProps, globalProps, "Površinska zaštita");
            if (string.IsNullOrWhiteSpace(row.PovrsinskaZastita)) row.PovrsinskaZastita = LookupProp(confProps, globalProps, "Povrsinska zastita");
            row.Finish = LookupProp(confProps, globalProps, "Finish");

            if (isPart && row.Thickness > 0 && !materialForExcel.Equals("--", StringComparison.OrdinalIgnoreCase))
            {
                PropSet activeConfProps = confProps.Count > 0 ? confProps[confProps.Count - 1] : null;
                ExtractKFactor(scan, activeConfProps, globalProps, out row.KFactorStr, out row.KFactorVal);
            }
        }

        private FeatureScan ScanFeatures(ModelDoc2 model, bool isPart)
        {
            FeatureScan scan = new FeatureScan();

            string title = (model.GetTitle() ?? "").ToLower();
            bool checkThreads = !title.Contains("din") && !title.Contains("iso") && !title.Contains("vijak") && !title.Contains("matica");
            bool checkCsk = checkThreads && !title.Contains("pem") && !title.Contains("zeger") && !title.Contains("podloška");

            // Sklop bez navoja/upuštanja – nema potrebe prolaziti stablo
            if (!isPart && !checkThreads) return scan;

            long threadCount = 0;
            Feature feat = model.FirstFeature() as Feature;
            while (feat != null)
            {
                string typeName = feat.GetTypeName2() ?? "";
                string typeLower = typeName.ToLowerInvariant();

                if (isPart)
                {
                    if (typeName == "SheetMetal")
                    {
                        scan.IsSheetMetal = true;
                        try
                        {
                            if (feat.GetDefinition() is SheetMetalFeatureData smData) scan.Thickness = smData.Thickness * 1000;
                        }
                        catch { }
                    }
                    if (typeLower.StartsWith("sheetmetal") || typeLower.StartsWith("sm") || typeLower.Contains("bend") || typeLower.Contains("flange") || typeLower == "flatpattern")
                    {
                        scan.SheetMetalFeatures.Add(feat);
                    }
                }

                bool isHole = typeName == "WinHole" || typeName == "HoleWzd";
                if (checkThreads && isHole)
                {
                    dynamic holeData = null;
                    try { holeData = feat.GetDefinition(); } catch { }
                    if (holeData != null)
                    {
                        try
                        {
                            int holeType = Convert.ToInt32(holeData.Type);
                            if (holeType == 2 || holeType == 3 || holeType == 5)
                            {
                                Array refLocs = null;
                                try { refLocs = holeData.ReferenceLocations as Array; } catch { }
                                double holes = (refLocs != null && refLocs.Length > 0) ? refLocs.Length / 3.0 : 1;
                                threadCount = (long)Math.Round(threadCount + holes, MidpointRounding.ToEven); // kao Long u makrou
                            }
                        }
                        catch { }

                        if (checkCsk && (feat.Name ?? "").ToLower().Contains("csk"))
                        {
                            Array points = null;
                            try { points = holeData.ReferenceLocations as Array; } catch { }
                            if (points == null || points.Length == 0)
                            {
                                try { points = holeData.GetSketchPoints() as Array; } catch { }
                            }
                            scan.CountersinkCount += (points != null && points.Length > 0) ? points.Length : 1;
                        }
                    }
                }
                else if (checkThreads && typeName == "CosmeticThread")
                {
                    threadCount++;
                }

                bool isBodyFolder = isPart && (typeName == "SolidBodyFolder" || typeLower.Contains("bodyfolder") || typeLower.Contains("cutlist"));
                if (checkThreads || isBodyFolder)
                {
                    Feature sub = feat.GetFirstSubFeature() as Feature;
                    while (sub != null)
                    {
                        string subType = sub.GetTypeName2() ?? "";

                        if (checkThreads && subType == "CosmeticThread") threadCount++;

                        if (isBodyFolder)
                        {
                            if (typeName == "SolidBodyFolder" && subType == "CutListFolder")
                            {
                                ReadCutListData(sub.CustomPropertyManager, scan);
                            }
                            if (subType == "CutListFolder" || subType == "SubWeldFolder" || subType.ToLowerInvariant().Contains("cutlist"))
                            {
                                CustomPropertyManager mgr = sub.CustomPropertyManager;
                                if (mgr != null) scan.CutListManagers.Add(mgr);
                            }
                        }

                        sub = sub.GetNextSubFeature() as Feature;
                    }
                }

                feat = feat.GetNextFeature() as Feature;
            }

            scan.ThreadCount = threadCount;
            return scan;
        }

        private static void ReadCutListData(CustomPropertyManager propMgr, FeatureScan scan)
        {
            if (propMgr == null) return;

            scan.Bends = GetResolved(propMgr, "Bends");
            scan.FlattenLength = GetResolved(propMgr, "Bounding Box Width");
            scan.FlattenWidth = GetResolved(propMgr, "Bounding Box Length");

            string cutMat = GetResolved(propMgr, "Material");
            if (cutMat.Trim() != "") scan.CutListMaterial = cutMat;

            string cutOuter = GetResolved(propMgr, "Cutting Length-Outer");
            string cutInner = GetResolved(propMgr, "Cutting Length-Inner");
            if (TryParseNumber(cutOuter, out double outer) && TryParseNumber(cutInner, out double inner))
            {
                scan.FlattenPerimeter = outer + inner;
            }
        }

        private static string GetResolved(CustomPropertyManager propMgr, string name)
        {
            propMgr.Get4(name, true, out string _, out string resolved);
            return resolved ?? "";
        }

        private static void WriteWeldData(ModelDoc2 model, bool isStainless, RowData row, List<string> kidNames)
        {
            bool hasWeld = false;
            double suma2mm = 0, suma5mm = 0, sumaOstalo = 0;
            HashSet<string> doneFolders = new HashSet<string>();

            ModelDocExtension ext = model.Extension;
            int k = 1;
            while (ext.SelectByID2("Weld Bead" + k, "COSMETIC_WELD", 0, 0, 0, false, 0, null, 0))
            {
                hasWeld = true;
                try
                {
                    SelectionMgr selMgr = model.SelectionManager as SelectionMgr;
                    Feature weldFeat = selMgr?.GetSelectedObject6(1, 0) as Feature;
                    if (weldFeat != null)
                    {
                        dynamic weldData = weldFeat.GetDefinition();
                        dynamic weldFolder = weldData?.GetWeldBeadFolder();
                        if (weldFolder != null)
                        {
                            double totalLength = Convert.ToDouble(weldFolder.TotalLength);
                            if (doneFolders.Add("F_" + totalLength.ToString("R", CultureInfo.InvariantCulture)))
                            {
                                double beadMm = Math.Round(Convert.ToDouble(weldData.BeadSize) * 1000.0, 0);
                                if (beadMm < 4) suma2mm += totalLength;
                                else if (beadMm <= 6) suma5mm += totalLength;
                                else sumaOstalo += totalLength;
                            }
                        }
                    }
                }
                catch { }
                k++;
            }
            if (hasWeld) { try { model.ClearSelection2(true); } catch { } }

            double rezultat = suma2mm + (suma5mm * 2.0) + (sumaOstalo * 3.0);
            if (rezultat > 0)
            {
                row.Set(isStainless ? ColZavarInox : ColZavarCelik, Math.Round(rezultat, 2));
            }

            if (kidNames == null || kidNames.Count == 0) return;

            long pCelik = 0, pInox = 0, wCount = 0;
            foreach (string n in kidNames)
            {
                if (!n.Contains("vijak") && !n.Contains("matica") && !n.Contains("din") && !n.Contains("iso"))
                {
                    wCount++;
                }
                if (n.Contains("din929") || n.Contains("iso13918"))
                {
                    if (n.Contains("a2k")) pCelik++;
                    else if (n.Contains("a2")) pInox++;
                    else pCelik++;
                }
            }

            if (hasWeld && wCount > 1)
            {
                row.Set(isStainless ? ColZavarKomadaInox : ColZavarKomadaCelik, wCount - 1);
            }
            if (pCelik > 0) row.Set(ColNavarneCelik, pCelik);
            if (pInox > 0) row.Set(ColNavarneInox, pInox);
        }

        /// <summary>
        /// K-faktor redoslijedom: cut-list propertyji, definicija sheet metal feature-a, propertyji modela.
        /// Koristi podatke prikupljene u ScanFeatures, pa se stablo ne prolazi ponovo.
        /// </summary>
        private static void ExtractKFactor(FeatureScan scan, PropSet activeConfProps, PropSet globalProps, out string kFactorStr, out double kFactorVal)
        {
            kFactorStr = "";
            kFactorVal = 0;

            // 1. Cut-list (Bend Allowance gdje je i Bends)
            foreach (CustomPropertyManager mgr in scan.CutListManagers)
            {
                PropSet cutProps;
                try { cutProps = PropSet.Read(mgr); } catch { continue; }

                foreach (string key in CutListKKeys)
                {
                    string val = cutProps.Value(key);
                    if (string.IsNullOrWhiteSpace(val) || val.Contains("@@")) continue;

                    string cleaned = CleanNumber(val);
                    if (double.TryParse(cleaned, NumberStyles.Any, CultureInfo.InvariantCulture, out double parsed) && parsed > 0)
                    {
                        if (parsed < 0.01) parsed *= 1000.0;
                        kFactorVal = Math.Round(parsed, 4);
                        kFactorStr = kFactorVal.ToString("0.####", CultureInfo.InvariantCulture);
                        return;
                    }
                    if (string.IsNullOrWhiteSpace(kFactorStr) && cleaned != "0" && cleaned != "0.00" && cleaned != "0,00")
                    {
                        kFactorStr = val.Trim();
                    }
                }

                foreach (PropEntry p in cutProps.Entries)
                {
                    string nameLower = p.Name.Trim().ToLower();
                    if (!(nameLower.Contains("bend allowance") || nameLower.Contains("k-factor") || nameLower.Contains("k factor") ||
                          nameLower.Contains("k_faktor") || nameLower.Contains("bend deduction") || nameLower.Contains("dodatak za savijanje")))
                    {
                        continue;
                    }

                    string val = p.Value;
                    if (string.IsNullOrWhiteSpace(val) || val.Contains("@@")) continue;

                    string cleaned = CleanNumber(val);
                    if (double.TryParse(cleaned, NumberStyles.Any, CultureInfo.InvariantCulture, out double parsed) && parsed > 0)
                    {
                        if (parsed < 0.01) parsed *= 1000.0;
                        kFactorVal = Math.Round(parsed, 4);
                        kFactorStr = cleaned;
                        return;
                    }
                }
            }

            // 2. Definicija sheet metal feature-a
            foreach (Feature feat in scan.SheetMetalFeatures)
            {
                dynamic def;
                try { def = feat.GetDefinition(); } catch { continue; }
                if (def == null) continue;

                try
                {
                    dynamic cba = def.GetCustomBendAllowance();
                    if (cba != null)
                    {
                        if (TryPositive(() => cba.BendAllowance, out double ba))
                        {
                            SetK(ba < 0.1 ? ba * 1000.0 : ba, out kFactorStr, out kFactorVal);
                            return;
                        }
                        if (TryPositive(() => cba.KFactor, out double ck))
                        {
                            SetK(ck, out kFactorStr, out kFactorVal);
                            return;
                        }
                        if (TryPositive(() => cba.BendDeduction, out double bd))
                        {
                            SetK(bd < 0.1 ? bd * 1000.0 : bd, out kFactorStr, out kFactorVal);
                            return;
                        }
                    }
                }
                catch { }

                if (TryPositive(() => def.BendAllowance, out double defBa))
                {
                    SetK(defBa < 0.1 ? defBa * 1000.0 : defBa, out kFactorStr, out kFactorVal);
                    return;
                }
                if (TryPositive(() => def.KFactor, out double defK))
                {
                    SetK(defK, out kFactorStr, out kFactorVal);
                    return;
                }
            }

            // 3. Propertyji konfiguracije i globalni propertyji dijela
            foreach (PropSet props in new[] { activeConfProps, globalProps })
            {
                if (props == null) continue;
                foreach (string key in ModelKKeys)
                {
                    string val = props.Value(key);
                    if (string.IsNullOrWhiteSpace(val) || val.Contains("@@")) continue;

                    string cleaned = CleanNumber(val);
                    if (double.TryParse(cleaned, NumberStyles.Any, CultureInfo.InvariantCulture, out double parsed) && parsed > 0)
                    {
                        kFactorVal = Math.Round(parsed, 4);
                        kFactorStr = cleaned;
                        return;
                    }
                }
            }
        }

        private static void SetK(double value, out string kStr, out double kVal)
        {
            kVal = Math.Round(value, 4);
            kStr = kVal.ToString("0.####", CultureInfo.InvariantCulture);
        }

        private static bool TryPositive(Func<object> getter, out double value)
        {
            value = 0;
            try
            {
                value = Convert.ToDouble(getter());
                return value > 0;
            }
            catch
            {
                return false;
            }
        }

        private static string CleanNumber(string val)
        {
            return val.ToLower().Replace("mm", "").Trim().Replace(',', '.');
        }

        private MainAssemblyData ReadMainAssemblyData()
        {
            MainAssemblyData main = new MainAssemblyData();

            main.ChildCount = GetMainAssemblyChildCount(swModel);
            if (main.ChildCount <= 0) main.ChildCount = 1;

            try
            {
                int warning = 0;
                if (swModel.Extension.GetMassProperties(1, ref warning) is Array massProps)
                {
                    if (massProps.Length >= 4) main.VolumeM3 = Convert.ToDouble(massProps.GetValue(3));
                    if (massProps.Length >= 5) main.AreaM2 = Convert.ToDouble(massProps.GetValue(4));
                }
            }
            catch { }

            if (main.VolumeM3 == 0 && swModel is AssemblyDoc asm)
            {
                try
                {
                    if (asm.GetBox(0) is Array box && box.Length >= 6)
                    {
                        double bX = Math.Abs(Convert.ToDouble(box.GetValue(3)) - Convert.ToDouble(box.GetValue(0)));
                        double bY = Math.Abs(Convert.ToDouble(box.GetValue(4)) - Convert.ToDouble(box.GetValue(1)));
                        double bZ = Math.Abs(Convert.ToDouble(box.GetValue(5)) - Convert.ToDouble(box.GetValue(2)));
                        main.VolumeM3 = bX * bY * bZ;
                    }
                }
                catch { }
            }

            PropSet globalProps = PropSet.Read(swModel.Extension.get_CustomPropertyManager(""));
            List<PropSet> confProps = ReadConfigurationProps(swModel, null);
            main.PovrsinskaZastita = LookupProp(confProps, globalProps, "Površinska zaštita");
            if (string.IsNullOrWhiteSpace(main.PovrsinskaZastita)) main.PovrsinskaZastita = LookupProp(confProps, globalProps, "Povrsinska zastita");
            main.Finish = LookupProp(confProps, globalProps, "Finish");

            return main;
        }

        private static long GetMainAssemblyChildCount(ModelDoc2 model)
        {
            long count = 0;
            try
            {
                if (model is AssemblyDoc asm && asm.GetComponents(true) is object[] comps)
                {
                    foreach (object c in comps)
                    {
                        Component2 child = c as Component2;
                        if (child == null || child.IsSuppressed()) continue;

                        string name = (child.Name2 ?? "").ToLower();
                        if (!name.Contains("vijak") && !name.Contains("matica") &&
                            !name.Contains("din") && !name.Contains("iso") &&
                            !name.Contains("zeger") && !name.Contains("podloška") && !name.Contains("podloska") &&
                            !IsPressFitNutOrFastener(name) && !IsAssemblyNut(name))
                        {
                            count++;
                        }
                    }
                }
            }
            catch { }
            return count;
        }

        private static List<string> GetTopLevelChildNames(ModelDoc2 model)
        {
            List<string> names = new List<string>();
            try
            {
                Component2 root = model.ConfigurationManager?.ActiveConfiguration?.GetRootComponent3(true);
                if (root?.GetChildren() is object[] kids)
                {
                    foreach (object k in kids)
                    {
                        if (k is Component2 kid) names.Add((kid.Name2 ?? "").ToLower());
                    }
                }
            }
            catch { }
            return names;
        }

        private static string GetPartMaterial(ModelDoc2 model)
        {
            try
            {
                return (model as PartDoc)?.GetMaterialPropertyName2("", out string _) ?? "";
            }
            catch
            {
                return "";
            }
        }

        private static string GetFirstPartMaterialInAssembly(ModelDoc2 model)
        {
            try
            {
                if (model is AssemblyDoc asm && asm.GetComponents(true) is object[] comps)
                {
                    foreach (object c in comps)
                    {
                        ModelDoc2 sub = (c as Component2)?.GetModelDoc2() as ModelDoc2;
                        if (sub != null && sub.GetType() == (int)swDocumentTypes_e.swDocPART)
                        {
                            string mat = GetPartMaterial(sub);
                            if (mat.Trim() != "") return mat;
                        }
                    }
                }
            }
            catch { }
            return "";
        }

        private static bool IsUvecanPerimetar(string material)
        {
            string m = (material ?? "").ToUpper();
            return m.Contains("1.4301") || m.Contains("EN-AW") || m.Contains("GALVANIZED STEEL");
        }

        private static int GetPerimeterColumn(double thickness)
        {
            double[] debljine = { 1, 1.5, 2, 3, 4, 5, 8, 10, 12, 15, 20 };
            int[] kolone = { 17, 20, 23, 26, 29, 35, 38, 41, 44, 47, 50 };
            for (int i = 0; i < debljine.Length; i++)
            {
                if (Math.Abs(thickness - debljine[i]) < 1e-6) return kolone[i];
            }
            return 0;
        }

        private static void AzurirajMaterijalUProperties(ModelDoc2 model, bool isAssembly, string swMaterialName, double thickness)
        {
            string finalStandard = "";
            if (isAssembly)
            {
                finalStandard = "--";
            }
            else
            {
                string mat = (swMaterialName ?? "").ToUpper();
                if (mat.Contains("PLAIN CARBON STEEL"))
                {
                    if (thickness >= 0.8 && thickness <= 3.5) finalStandard = "EN 10130-DC01-A-m";
                    else if (thickness >= 4 && thickness <= 10.5) finalStandard = "EN 10111-DD11";
                    else finalStandard = "PLAIN CARBON STEEL";
                }
                else if (mat.Contains("S235")) finalStandard = mat;
                else if (mat.Contains("S355")) finalStandard = mat;
                else if (mat.Contains("GALVANIZED STEEL")) finalStandard = "DX51D+Z275";
                else if (mat.Contains("1.4301")) finalStandard = "EN 10088-2-1.4301";
            }

            if (finalStandard != "")
            {
                try
                {
                    model.Extension.get_CustomPropertyManager("")?.Add3("Material", (int)swCustomInfoType_e.swCustomInfoText, finalStandard, (int)swCustomPropertyAddOption_e.swCustomPropertyOnlyIfNew);
                }
                catch { }
            }
        }

        /// <summary>Propertyji konfiguracija redoslijedom: referencirana konfiguracija komponente, pa aktivna.</summary>
        private static List<PropSet> ReadConfigurationProps(ModelDoc2 model, Component2 comp)
        {
            List<string> confNames = new List<string>();
            if (comp != null)
            {
                try
                {
                    string refConf = comp.ReferencedConfiguration;
                    if (!string.IsNullOrWhiteSpace(refConf)) confNames.Add(refConf);
                }
                catch { }
            }
            try
            {
                string activeConf = model.ConfigurationManager?.ActiveConfiguration?.Name;
                if (!string.IsNullOrEmpty(activeConf) && !confNames.Contains(activeConf)) confNames.Add(activeConf);
            }
            catch { }

            List<PropSet> result = new List<PropSet>();
            foreach (string conf in confNames)
            {
                try { result.Add(PropSet.Read(model.Extension.get_CustomPropertyManager(conf))); } catch { }
            }
            return result;
        }

        private static string LookupProp(List<PropSet> confProps, PropSet globalProps, string name)
        {
            foreach (PropSet props in confProps)
            {
                string v = props.Value(name);
                if (!string.IsNullOrWhiteSpace(v)) return v.Trim(' ', '"', '\'');
            }
            string g = globalProps?.Value(name);
            return !string.IsNullOrWhiteSpace(g) ? g.Trim(' ', '"', '\'') : "";
        }

        private IntPtr GetSolidWorksWindow()
        {
            try
            {
                dynamic frame = swApp.IFrameObject();
                if (frame != null)
                {
                    IntPtr hwnd = new IntPtr(Convert.ToInt64(frame.GetHWnd()));
                    if (hwnd != IntPtr.Zero) return hwnd;
                }
            }
            catch { }

            try
            {
                Process[] processes = Process.GetProcessesByName("SLDWORKS");
                if (processes.Length > 0) return processes[0].MainWindowHandle;
            }
            catch { }

            return IntPtr.Zero;
        }

        #endregion

        #region Excel – upis

        private void WriteExcel(string excelFile, string ident, string korisnik, string datum, MainAssemblyData main)
        {
            Type excelType = Type.GetTypeFromProgID("Excel.Application");
            if (excelType == null) throw new InvalidOperationException("Microsoft Excel nije instaliran na ovom računaru.");

            dynamic excelApp = null;
            dynamic workbook = null;
            dynamic sheet = null;
            try
            {
                excelApp = Activator.CreateInstance(excelType);
                excelApp.Visible = false;
                excelApp.DisplayAlerts = false;

                workbook = excelApp.Workbooks.Open(excelFile);
                sheet = workbook.Sheets[1];

                try { excelApp.ScreenUpdating = false; } catch { }
                try { excelApp.EnableEvents = false; } catch { }
                try { excelApp.Calculation = -4135; } catch { } // xlCalculationManual

                excelApp.DecimalSeparator = ",";
                excelApp.ThousandsSeparator = ".";
                excelApp.UseSystemSeparators = false;

                sheet.Cells[2, 1].Value = "IDENT: " + ident;
                sheet.Cells[4, 1].Value = "Ime i Prezime: " + korisnik;
                sheet.Cells[5, 1].Value = "Datum: " + datum;

                // Postojeće vrijednosti čitamo jednom – za ručno unesene podatke i čišćenje matica
                CellBuffer buffer = new CellBuffer(FirstRow, LastScanRow);
                foreach (int col in new[] { ColNaziv, ColIdent, ColOsnovniNaziv, ColKolicina, ColDebljina, ColFinish,
                                            ColPlastifikacija, ColVijci, ColMatice, ColMaterijal, ColPovrsinskaZastita })
                {
                    buffer.LoadColumn((object)sheet, col);
                }

                foreach (RowData row in rows)
                {
                    foreach (KeyValuePair<int, object> cell in row.Cells) buffer.Set(row.Row, cell.Key, cell.Value);
                    buffer.Set(row.Row, ColKolicina, dictQuantity[row.FinalName]);
                }

                int gpRow = mainRow != null ? mainRow.Row : (rows.Count > 0 ? rows[rows.Count - 1].Row : FirstRow);
                Dictionary<int, RowData> rowsByIndex = new Dictionary<int, RowData>();
                foreach (RowData row in rows) rowsByIndex[row.Row] = row;

                List<int> highlightRows = new List<int>();
                for (int r = FirstRow; r <= gpRow; r++)
                {
                    rowsByIndex.TryGetValue(r, out RowData row);
                    bool isGp = r == gpRow;

                    if (isGp) WriteSurfaceProtection(buffer, r, main.PovrsinskaZastita, main.Finish);
                    else WriteSurfaceProtection(buffer, r, row?.PovrsinskaZastita, row?.Finish);
                    buffer.Set(r, ColBrojKomponenti, isGp ? (object)main.ChildCount : null);

                    object kValue = null;
                    if (!isGp && row != null && row.Thickness > 0 && !row.Material.Equals("--", StringComparison.OrdinalIgnoreCase))
                    {
                        kValue = GetKFactorCellValue(row, out bool highlight);
                        if (highlight) highlightRows.Add(r);
                    }
                    buffer.Set(r, ColKFaktor, kValue);
                }

                if (main.AreaM2 > 0) buffer.Set(gpRow, ColPovrsina, Math.Round(main.AreaM2, 4));
                if (main.VolumeM3 > 0) buffer.Set(gpRow, ColVolumen, Math.Round(main.VolumeM3, 6));

                SanitizeNutAndFastenerColumns(buffer);

                try
                {
                    sheet.Columns["EH"].ColumnWidth = 14.0;
                    sheet.Cells[8, ColKFaktor].Value = "K - faktor";
                }
                catch { }
                try { sheet.Range[$"EH{FirstRow}:EH{gpRow}"].NumberFormat = "General"; } catch { }

                buffer.Flush((object)sheet);

                RestorePlastifikacijaFormula((object)sheet, gpRow);

                try { sheet.Range[$"EH{FirstRow}:EH{gpRow}"].Interior.ColorIndex = -4142; } catch { }
                foreach (int r in highlightRows)
                {
                    try { sheet.Cells[r, ColKFaktor].Interior.Color = KFaktorOdstupanjeBoja; } catch { }
                }

                PostaviValidaciju((object)sheet);

                try { excelApp.Calculation = -4105; } catch { } // xlCalculationAutomatic
                try { excelApp.EnableEvents = true; } catch { }
                try { excelApp.ScreenUpdating = true; } catch { }

                workbook.Save();
                workbook.Close(false);
                workbook = null;
            }
            finally
            {
                if (workbook != null) { try { workbook.Close(false); } catch { } }
                if (excelApp != null) { try { excelApp.Quit(); } catch { } }
                sheet = null;
                workbook = null;
                excelApp = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
            }
        }

        /// <summary>
        /// EF = "Površinska zaštita" tačno kako piše u modelu (i "-", "--"); ako je prazna – EF prazno.
        /// J = "Finish" iz modela, ali samo ako površinska zaštita nije prazna, "-" ili "--"; inače J prazno.
        /// </summary>
        private static void WriteSurfaceProtection(CellBuffer buffer, int r, string pzFromModel, string finishFromModel)
        {
            string pz = (pzFromModel ?? "").Trim();
            buffer.Set(r, ColPovrsinskaZastita, pz == "" ? null : pz);

            if (pz == "" || pz == "-" || pz == "--")
            {
                buffer.Set(r, ColFinish, null);
                return;
            }

            string finish = (finishFromModel ?? "").Trim();
            if (finish == "")
            {
                buffer.Set(r, ColFinish, null);
            }
            else if (double.TryParse(finish.Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out double finishNum))
            {
                buffer.Set(r, ColFinish, finishNum);
            }
            else
            {
                buffer.Set(r, ColFinish, finish);
            }
        }

        /// <summary>
        /// K-faktor iz modela; ako ga nema – vrijednost iz baze K-faktora.
        /// Ako se model razlikuje od baze za više od 0,005 – ćelija se boji.
        /// </summary>
        private static object GetKFactorCellValue(RowData row, out bool highlight)
        {
            highlight = false;

            double modelK = row.KFactorVal;
            if (modelK == 0 && !string.IsNullOrWhiteSpace(row.KFactorStr) &&
                double.TryParse(row.KFactorStr.Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out double parsed) && parsed > 0)
            {
                modelK = parsed;
            }

            double expectedK = KFactorManager.GetKFactor(row.Material, row.Thickness);
            if (modelK > 0)
            {
                highlight = expectedK > 0 && Math.Abs(modelK - expectedK) > 0.005;
                return Math.Round(modelK, 4);
            }
            if (expectedK > 0) return Math.Round(expectedK, 4);
            return null;
        }

        /// <summary>
        /// Obične matice za montažu idu u DK; uprešne (PEM) matice se čiste iz DK/DJ.
        /// Kolona CZ (Plastifikacija) je formula i ovdje se nikad ne upisuje.
        /// Prolazi kroz sve popunjene redove (do 3 uzastopna prazna).
        /// </summary>
        private static void SanitizeNutAndFastenerColumns(CellBuffer buffer)
        {
            int emptyCount = 0;
            for (int r = FirstRow; r <= LastScanRow; r++)
            {
                string pName = Convert.ToString(buffer.Get(r, ColNaziv))?.Trim() ?? "";
                string pIdent = Convert.ToString(buffer.Get(r, ColIdent))?.Trim() ?? "";
                string pDesc = Convert.ToString(buffer.Get(r, ColOsnovniNaziv))?.Trim() ?? "";

                if (pName == "" && pIdent == "" && pDesc == "")
                {
                    emptyCount++;
                    if (emptyCount >= 3) break;
                    continue;
                }
                emptyCount = 0;

                string checkStr = $"{pIdent} {pName} {pDesc}".ToLower();

                int qty = 1;
                object qtyVal = buffer.Get(r, ColKolicina);
                if (qtyVal != null && int.TryParse(Convert.ToString(qtyVal), out int parsedQty) && parsedQty > 0) qty = parsedQty;

                if (IsPressFitNutOrFastener(checkStr))
                {
                    buffer.Set(r, ColMatice, null);
                    if (!checkStr.Contains("vijak")) buffer.Set(r, ColVijci, null);
                }
                else if (IsAssemblyNut(checkStr))
                {
                    buffer.Set(r, ColMatice, qty);
                }
                else
                {
                    object val104 = buffer.Get(r, ColPlastifikacija);
                    if (val104 != null && !string.IsNullOrWhiteSpace(Convert.ToString(val104)))
                    {
                        buffer.Set(r, ColMatice, null);
                    }

                    double thicknessVal = 0;
                    object thObj = buffer.Get(r, ColDebljina);
                    if (thObj != null) double.TryParse(Convert.ToString(thObj, CultureInfo.InvariantCulture).Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out thicknessVal);

                    string materialStr = Convert.ToString(buffer.Get(r, ColMaterijal))?.Trim() ?? "";
                    if (thicknessVal > 0 && !materialStr.Equals("--", StringComparison.OrdinalIgnoreCase))
                    {
                        buffer.Set(r, ColMatice, null);
                    }
                }
            }
        }

        /// <summary>
        /// Ranije verzije su u CZ (Plastifikacija) upisivale količinu PEM matica i brisale formulu.
        /// Ako u redu nema formule, vraća se formula iz šablona.
        /// </summary>
        private static void RestorePlastifikacijaFormula(dynamic sheet, int lastRow)
        {
            try
            {
                if (!(sheet.Range[$"CZ{FirstRow}:CZ{lastRow}"].Formula is object[,] formulas)) return;

                int lower = formulas.GetLowerBound(0);
                for (int i = 0; i <= lastRow - FirstRow; i++)
                {
                    string f = Convert.ToString(formulas[lower + i, formulas.GetLowerBound(1)]) ?? "";
                    if (f.StartsWith("=")) continue;

                    int r = FirstRow + i;
                    sheet.Range[$"CZ{r}"].Formula = $"=IF(J{r}=0,\"--\",J{r}/$CZ$3/(24)*DA{r})";
                }
            }
            catch { }
        }

        private static void PostaviValidaciju(dynamic sheet)
        {
            try
            {
                dynamic daValidation = sheet.Range["DA10:DA300"].Validation;
                daValidation.Delete();
                daValidation.Add(Type: 3, AlertStyle: 1, Formula1: "=$ER$1000:$EU$1000");
                daValidation.IgnoreBlank = true;
                daValidation.InCellDropdown = true;

                dynamic cqValidation = sheet.Range["CQ10:CQ300"].Validation;
                cqValidation.Delete();
                cqValidation.Add(Type: 3, AlertStyle: 1, Formula1: "=$ER$999:$ES$999");
                cqValidation.IgnoreBlank = true;
                cqValidation.InCellDropdown = true;
            }
            catch { }
        }

        #endregion

        #region Pomoćne funkcije

        /// <summary>Čita property sa aktivne konfiguracije, pa globalno (koristi HomeView za IDENT i PUTANJA).</summary>
        public static string GetModelProperty(ModelDoc2 doc, string propName)
        {
            if (doc == null || string.IsNullOrWhiteSpace(propName)) return "";
            try
            {
                PropSet globalProps = PropSet.Read(doc.Extension.get_CustomPropertyManager(""));
                return LookupProp(ReadConfigurationProps(doc, null), globalProps, propName);
            }
            catch
            {
                return "";
            }
        }

        /// <summary>Tekst iz SolidWorksa upisuje se kakav jeste (kao u makrou) – Excel ga sam pretvara u broj.</summary>
        private static object ToCellValue(string value)
        {
            return string.IsNullOrWhiteSpace(value) ? null : value;
        }

        private static bool TryParseNumber(string value, out double number)
        {
            number = 0;
            if (string.IsNullOrWhiteSpace(value)) return false;
            return double.TryParse(value.Trim().Replace(',', '.'), NumberStyles.Float, CultureInfo.InvariantCulture, out number);
        }

        private static bool IsFileLocked(string path)
        {
            if (!File.Exists(path)) return false;
            try
            {
                using (new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None)) { }
                return false;
            }
            catch (IOException)
            {
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static string ColumnLetter(int col)
        {
            StringBuilder sb = new StringBuilder();
            while (col > 0)
            {
                int rem = (col - 1) % 26;
                sb.Insert(0, (char)('A' + rem));
                col = (col - 1) / 26;
            }
            return sb.ToString();
        }

        internal static bool IsPressFitNutOrFastener(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return false;
            string n = name.ToLower();

            if (n.Contains("pem") || n.Contains("upres") || n.Contains("upreš") || n.Contains("press") || n.Contains("plm") || n.Contains("upresavanje") || n.Contains("uprešavanje"))
            {
                return true;
            }

            if (n.Contains("s-m") || n.Contains("s_m") || n.Contains("ss-m") || n.Contains("sp-m") || n.Contains("so-m") ||
                n.Contains("cls-m") || n.Contains("clss-m") || n.Contains("css-m") || n.Contains("fh-m") || n.Contains("fhs-m") ||
                n.Contains("bso-m") || n.Contains("bsos-m") || n.Contains("tso-m") || n.Contains("tsos-m"))
            {
                return true;
            }

            string[] multiLetterPressPrefixes = { "bsos", "bso", "tsos", "tso", "clss", "cls", "fhs", "fha", "kss" };

            string[] tokens = n.Split(new[] { ' ', '_', '-', '.', '<', '>', '/', '\\', '(', ')', ':', ';', '=', '+' }, StringSplitOptions.RemoveEmptyEntries);
            for (int i = 0; i < tokens.Length; i++)
            {
                string t = tokens[i];
                if (t == "pem" || t == "upres" || t == "upresna" || t == "uprešna" || t == "press" || t == "plm" || t == "upresnamatica" || t == "upresna_matica" || t == "upresavanje" || t == "uprešavanje") return true;

                if ((t == "s" || t == "ss" || t == "sp" || t == "so" || t == "ps" || t == "fh" || t == "cls") &&
                    i < tokens.Length - 1 && tokens[i + 1].StartsWith("m") && tokens[i + 1].Length >= 2 && char.IsDigit(tokens[i + 1][1]))
                {
                    return true;
                }

                foreach (string p in multiLetterPressPrefixes)
                {
                    if (t == p) return true;
                    if (t.StartsWith(p))
                    {
                        string sub = t.Substring(p.Length);
                        if (sub.Length == 0 || (sub.StartsWith("m") && sub.Length >= 2 && char.IsDigit(sub[1])))
                        {
                            return true;
                        }
                    }
                }
            }

            return false;
        }

        internal static bool IsAssemblyNut(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return false;
            if (IsPressFitNutOrFastener(name)) return false;

            string n = name.ToLower();
            if (n.Contains("din929") || n.Contains("iso13918") || n.Contains("din 929") || n.Contains("iso 13918")) return false;

            if (n.Contains("kvt") || n.Contains("huc") || n.Contains("feks") || n.Contains("riksg") || n.Contains("fef") || n.Contains("hut"))
            {
                return true;
            }

            if (n.Contains("din934") || n.Contains("din985") || n.Contains("iso4032") || n.Contains("din 934") || n.Contains("din 985") || n.Contains("iso 4032"))
            {
                return true;
            }

            if (n.Contains("matica") || n.Contains("nut"))
            {
                if (!n.Contains("upres") && !n.Contains("upreš") && !n.Contains("pem") && !n.Contains("press") && !n.Contains("plm"))
                {
                    return true;
                }
            }

            return false;
        }

        #endregion

        #region Pomoćne klase

        private class RowData
        {
            public int Row;
            public string FinalName;
            public bool IsMain;
            public bool IsAssembly;
            public double Thickness;
            public string Material = "";
            public string PovrsinskaZastita = "";
            public string Finish = "";
            public string KFactorStr = "";
            public double KFactorVal;
            public readonly Dictionary<int, object> Cells = new Dictionary<int, object>();

            public void Set(int col, object value) { Cells[col] = value; }
        }

        private class FeatureScan
        {
            public bool IsSheetMetal;
            public double Thickness;
            public string Bends = "";
            public string FlattenLength = "";
            public string FlattenWidth = "";
            public double FlattenPerimeter;
            public string CutListMaterial = "";
            public long ThreadCount;
            public long CountersinkCount;
            public readonly List<CustomPropertyManager> CutListManagers = new List<CustomPropertyManager>();
            public readonly List<Feature> SheetMetalFeatures = new List<Feature>();
        }

        private class MainAssemblyData
        {
            public long ChildCount;
            public double AreaM2;
            public double VolumeM3;
            public string PovrsinskaZastita = "";
            public string Finish = "";
        }

        private class PropEntry
        {
            public string Name;
            public string Value;
        }

        /// <summary>
        /// Propertyji jednog CustomPropertyManager-a.
        /// GetAll3 (jedan poziv) služi samo za spisak imena – njegov "Resolved" niz su statusni kodovi, a ne vrijednosti.
        /// Vrijednost se čita sa Get4 (kao u makrou), samo za propertyje koji postoje, i kešira se.
        /// </summary>
        private class PropSet
        {
            private readonly CustomPropertyManager mgr;
            private readonly List<string> names = new List<string>();
            private readonly Dictionary<string, string> exactName = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            private readonly Dictionary<string, string> cache = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            private PropSet(CustomPropertyManager mgr)
            {
                this.mgr = mgr;
            }

            public static PropSet Read(CustomPropertyManager mgr)
            {
                PropSet set = new PropSet(mgr);
                if (mgr == null) return set;

                object names = null, types = null, values = null, resolved = null, links = null;
                mgr.GetAll3(ref names, ref types, ref values, ref resolved, ref links);

                if (names is Array nameArr)
                {
                    foreach (object n in nameArr)
                    {
                        string name = Convert.ToString(n);
                        if (string.IsNullOrWhiteSpace(name) || set.exactName.ContainsKey(name.Trim())) continue;
                        set.exactName[name.Trim()] = name;
                        set.names.Add(name.Trim());
                    }
                }
                return set;
            }

            /// <summary>Svi propertyji redoslijedom kao u SolidWorksu (vrijednosti se čitaju po potrebi).</summary>
            public IEnumerable<PropEntry> Entries
            {
                get
                {
                    foreach (string name in names) yield return new PropEntry { Name = name, Value = Value(name) };
                }
            }

            /// <summary>Izračunata vrijednost (Get4); ako je prazna – unesena vrijednost.</summary>
            public string Value(string name)
            {
                if (mgr == null || !exactName.TryGetValue(name.Trim(), out string realName)) return "";
                if (cache.TryGetValue(realName, out string cached)) return cached;

                string result = "";
                try
                {
                    mgr.Get4(realName, true, out string raw, out string resolvedVal);
                    result = !string.IsNullOrWhiteSpace(resolvedVal) ? resolvedVal : (raw ?? "");
                }
                catch { }
                cache[realName] = result;
                return result;
            }

            public string Resolved(string name)
            {
                return Value(name);
            }
        }

        /// <summary>
        /// Skuplja sve upise i šalje ih u Excel na kraju: susjedne ćelije u istoj koloni idu jednim pozivom.
        /// Upisuju se samo ćelije koje je obrada postavila – formule u ostalim ćelijama ostaju netaknute.
        /// </summary>
        private class CellBuffer
        {
            private readonly int firstRow;
            private readonly int lastRow;
            private readonly Dictionary<int, object[]> original = new Dictionary<int, object[]>();
            private readonly Dictionary<int, SortedDictionary<int, object>> pending = new Dictionary<int, SortedDictionary<int, object>>();

            public CellBuffer(int firstRow, int lastRow)
            {
                this.firstRow = firstRow;
                this.lastRow = lastRow;
            }

            public void LoadColumn(dynamic sheet, int col)
            {
                object[] values = new object[lastRow - firstRow + 1];
                try
                {
                    string letter = ColumnLetter(col);
                    object raw = sheet.Range[$"{letter}{firstRow}:{letter}{lastRow}"].Value2;
                    if (raw is object[,] arr)
                    {
                        int lower = arr.GetLowerBound(0);
                        for (int i = 0; i < values.Length && i + lower <= arr.GetUpperBound(0); i++)
                        {
                            values[i] = arr[i + lower, arr.GetLowerBound(1)];
                        }
                    }
                }
                catch { }
                original[col] = values;
            }

            public void Set(int row, int col, object value)
            {
                if (!pending.TryGetValue(col, out SortedDictionary<int, object> colCells))
                {
                    colCells = new SortedDictionary<int, object>();
                    pending[col] = colCells;
                }
                colCells[row] = value;
            }

            public object Get(int row, int col)
            {
                if (pending.TryGetValue(col, out SortedDictionary<int, object> colCells) && colCells.TryGetValue(row, out object v)) return v;
                if (original.TryGetValue(col, out object[] values) && row >= firstRow && row <= lastRow) return values[row - firstRow];
                return null;
            }

            public void Flush(dynamic sheet)
            {
                foreach (KeyValuePair<int, SortedDictionary<int, object>> colEntry in pending)
                {
                    int col = colEntry.Key;
                    string letter = ColumnLetter(col);

                    List<KeyValuePair<int, object>> run = new List<KeyValuePair<int, object>>();
                    foreach (KeyValuePair<int, object> cell in colEntry.Value)
                    {
                        if (run.Count > 0 && cell.Key != run[run.Count - 1].Key + 1)
                        {
                            WriteRun((object)sheet, letter, run);
                            run.Clear();
                        }
                        run.Add(cell);
                    }
                    if (run.Count > 0) WriteRun((object)sheet, letter, run);
                }
                pending.Clear();
            }

            private static void WriteRun(dynamic sheet, string letter, List<KeyValuePair<int, object>> run)
            {
                int start = run[0].Key;
                int end = run[run.Count - 1].Key;
                if (run.Count == 1)
                {
                    sheet.Range[$"{letter}{start}"].Value = run[0].Value;
                    return;
                }

                object[,] data = new object[run.Count, 1];
                for (int i = 0; i < run.Count; i++) data[i, 0] = run[i].Value;
                sheet.Range[$"{letter}{start}:{letter}{end}"].Value = data;
            }
        }

        #endregion
    }
}
