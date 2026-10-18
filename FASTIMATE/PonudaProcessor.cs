using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace FASTIMATE
{
    public class PonudaPartInfo
    {
        public string CompKey { get; set; }
        public string PartName { get; set; }
        public string Material { get; set; }
        public double Thickness { get; set; }
        public double MassSingleKg { get; set; }
        public bool IsSheetMetal { get; set; }
    }

    public class PonudaProcessor
    {
        private SldWorks swApp;
        private ModelDoc2 swModel;
        private dynamic excelApp;
        private dynamic excelWorkbook;
        private dynamic excelSheet;

        private int rowIndex;
        private string ident;
        private string putanja;
        private string sLogovaniKorisnik;

        [DllImport("user32.dll")]
        private static extern bool EnableWindow(IntPtr hWnd, bool bEnable);

        private Dictionary<string, int> dictQuantity = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        private Dictionary<string, int> dictRowIndex = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        private Dictionary<string, string> dictVisitedModels = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        private Dictionary<string, string> dictModelMaterial = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        private Dictionary<string, bool> dictModelIsStainless = new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);

        public static bool RunProcess(SldWorks activeSwApp, ModelDoc2 activeSwModel, string logovaniKorisnik, string uiIdent = "", string uiPutanja = "")
        {
            try
            {
                var processor = new PonudaProcessor();
                return processor.Execute(activeSwApp, activeSwModel, logovaniKorisnik, uiIdent, uiPutanja);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Greška pri obradi Ponude: {ex.Message}", "Greška", MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }
        }

        public bool Execute(SldWorks activeSwApp, ModelDoc2 activeSwModel, string logovaniKorisnik, string uiIdent = "", string uiPutanja = "")
        {
            swApp = activeSwApp;
            if (swApp == null)
            {
                try
                {
                    swApp = (SldWorks)Marshal.GetActiveObject("SldWorks.Application");
                }
                catch
                {
                    MessageBox.Show("SolidWorks aplikacija nije pokrenuta.", "Greška", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return false;
                }
            }

            swModel = activeSwModel ?? swApp.ActiveDoc as ModelDoc2;
            if (swModel == null)
            {
                MessageBox.Show("Nema otvorenog SolidWorks dokumenta.", "Greška", MessageBoxButton.OK, MessageBoxImage.Warning);
                return false;
            }

            sLogovaniKorisnik = logovaniKorisnik;

            if (!string.IsNullOrWhiteSpace(uiIdent))
            {
                ident = uiIdent.Trim();
            }
            else
            {
                CustomPropertyManager custPropMgr = swModel.Extension.get_CustomPropertyManager("");
                string valOut = "", resValOut = "";
                custPropMgr.Get4("IDENT", true, out valOut, out resValOut);
                ident = !string.IsNullOrWhiteSpace(resValOut) ? resValOut : valOut;
            }

            if (!string.IsNullOrWhiteSpace(uiPutanja))
            {
                putanja = uiPutanja.Trim();
            }
            else
            {
                CustomPropertyManager custPropMgr = swModel.Extension.get_CustomPropertyManager("");
                string valOut = "", resValOut = "";
                custPropMgr.Get4("PUTANJA", true, out valOut, out resValOut);
                putanja = !string.IsNullOrWhiteSpace(resValOut) ? resValOut : valOut;
            }

            if (string.IsNullOrWhiteSpace(putanja) && swModel != null)
            {
                string mPath = swModel.GetPathName();
                if (!string.IsNullOrWhiteSpace(mPath))
                    putanja = Path.GetDirectoryName(mPath);
            }

            if (string.IsNullOrWhiteSpace(ident) && swModel != null)
            {
                string mPath = swModel.GetPathName();
                if (!string.IsNullOrWhiteSpace(mPath))
                    ident = Path.GetFileNameWithoutExtension(mPath);
                else
                    ident = swModel.GetTitle();
            }

            if (string.IsNullOrWhiteSpace(putanja))
                putanja = System.Environment.GetFolderPath(System.Environment.SpecialFolder.Desktop);

            if (string.IsNullOrWhiteSpace(ident))
                ident = "DOKUMENT";

            try
            {
                if (!Directory.Exists(putanja))
                {
                    Directory.CreateDirectory(putanja);
                }
            }
            catch { }

            string excelFile = Path.Combine(putanja, $"PRORACUN_VREMENA_PONUDE_{ident}.xlsx");
            if (!File.Exists(excelFile))
            {
                byte[] templateBytes = TemplateManager.GetActiveTemplateBytes(1);
                if (templateBytes != null && templateBytes.Length > 0)
                {
                    try
                    {
                        File.WriteAllBytes(excelFile, templateBytes);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Ne mogu kreirati fajl ponude na lokaciji:\n{excelFile}\nGreška: {ex.Message}", "Greška", MessageBoxButton.OK, MessageBoxImage.Error);
                        return false;
                    }
                }
                else
                {
                    MessageBox.Show($"Excel fajl ne postoji na putanji:\n{excelFile}", "Greška", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return false;
                }
            }

            try
            {
                SheetMetalDbManager.LoadData();
                FastenerDbManager.LoadData();
                KFactorManager.LoadData();
            }
            catch { }

            Type excelType = Type.GetTypeFromProgID("Excel.Application");
            if (excelType == null)
            {
                MessageBox.Show("Microsoft Excel nije instaliran na ovom računaru.", "Greška", MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }

            excelApp = Activator.CreateInstance(excelType);
            excelApp.Visible = false;
            excelApp.DisplayAlerts = false;

            IntPtr swHwnd = IntPtr.Zero;
            try
            {
                swApp.Visible = true;
                dynamic frame = swApp.IFrameObject();
                if (frame != null)
                {
                    swHwnd = (IntPtr)frame.GetHWnd();
                }
            }
            catch { }

            if (swHwnd == IntPtr.Zero)
            {
                try
                {
                    var swProcesses = System.Diagnostics.Process.GetProcessesByName("SLDWORKS");
                    if (swProcesses.Length > 0)
                    {
                        swHwnd = swProcesses[0].MainWindowHandle;
                    }
                }
                catch { }
            }

            if (swHwnd != IntPtr.Zero)
            {
                try { EnableWindow(swHwnd, false); } catch { }
            }

            System.Windows.Forms.Form swLoadingBox = null;
            try
            {
                swLoadingBox = new System.Windows.Forms.Form
                {
                    Text = "SolidWorks - Proces u toku",
                    Width = 420,
                    Height = 130,
                    FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog,
                    MaximizeBox = false,
                    MinimizeBox = false,
                    ControlBox = false,
                    StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen,
                    TopMost = true,
                    BackColor = System.Drawing.Color.White
                };

                var lblText = new System.Windows.Forms.Label
                {
                    Text = "Proces u toku !",
                    ForeColor = System.Drawing.Color.FromArgb(2, 136, 209),
                    Font = new System.Drawing.Font("Segoe UI", 16, System.Drawing.FontStyle.Bold),
                    Dock = System.Windows.Forms.DockStyle.Fill,
                    TextAlign = System.Drawing.ContentAlignment.MiddleCenter
                };

                swLoadingBox.Controls.Add(lblText);
                swLoadingBox.Show();
                System.Windows.Forms.Application.DoEvents();
            }
            catch { }

            try
            {
                excelWorkbook = excelApp.Workbooks.Open(excelFile);
                excelSheet = excelWorkbook.Sheets[1];

                string datumUpisa = DateTime.Now.ToString("dd.MM.yyyy");
                excelSheet.Cells[2, 1].Value = "IDENT: " + ident;
                excelSheet.Cells[4, 1].Value = "Ime i Prezime: " + sLogovaniKorisnik;
                excelSheet.Cells[5, 1].Value = "Datum: " + datumUpisa;

                try
                {
                    excelSheet.Columns["EH"].ColumnWidth = 14.0;
                    excelSheet.Cells[8, 138].Value = "K - faktor";
                }
                catch { }

                excelApp.DecimalSeparator = ",";
                excelApp.ThousandsSeparator = ".";
                excelApp.UseSystemSeparators = false;

                try { excelApp.ScreenUpdating = false; } catch { }
                try { excelApp.EnableEvents = false; } catch { }
                try { excelApp.Calculation = -4135; } catch { } // xlCalculationManual

                rowIndex = 10;
                dictQuantity = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                dictRowIndex = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                dictVisitedModels = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                dictModelMaterial = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                dictModelIsStainless = new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);

                Configuration swConf = swModel.ConfigurationManager.ActiveConfiguration;
                Component2 swRootComp = swConf.GetRootComponent3(true);

                // Prolaz kroz sve podkomponente sklopa (glavni sklop se upisuje jednom na kraju kao gpRow)
                if (swRootComp != null)
                {
                    object vChildObj = swRootComp.GetChildren();
                    if (vChildObj is Array childrenArray)
                    {
                        foreach (object child in childrenArray)
                        {
                            Component2 swChild = child as Component2;
                            if (swChild != null)
                            {
                                ProcessComponents(swChild);
                            }
                        }
                    }
                }

                // Glavni sklop (GP) se postavlja na zadnji red nakon svih podpozicija
                int gpRow = rowIndex;

                string mainTitleStr = swModel.GetTitle() ?? "";
                string gsIdentStr = GetCustomPropValue(swModel, null, "GS_Ident");
                string gsNazivStr = GetCustomPropValue(swModel, null, "GS_Osnovni_naziv");
                string gsDodatniStr = GetCustomPropValue(swModel, null, "GS_Dodatni_opis");

                string mainFinalNameStr = !string.IsNullOrWhiteSpace(gsDodatniStr) ? gsDodatniStr.Trim() : (!string.IsNullOrWhiteSpace(gsNazivStr) ? gsNazivStr.Trim() : mainTitleStr);

                excelSheet.Cells[gpRow, 1].Value = mainFinalNameStr;
                excelSheet.Cells[gpRow, 2].Value = !string.IsNullOrWhiteSpace(gsIdentStr) ? gsIdentStr : ident;
                excelSheet.Cells[gpRow, 3].Value = gsNazivStr;
                excelSheet.Cells[gpRow, 4].Value = 1;

                // Upis konsolidovanih količina u kolonu 4 (D)
                foreach (var kvp in dictQuantity)
                {
                    if (dictRowIndex.ContainsKey(kvp.Key))
                    {
                        excelSheet.Cells[dictRowIndex[kvp.Key], 4].Value = kvp.Value;
                    }
                }

                UpisiPodatkeGlavnogSklopa(swModel, gpRow);

                bool mainIsStainless = false;
                if (swModel.GetType() == (int)swDocumentTypes_e.swDocASSEMBLY)
                {
                    mainIsStainless = DetermineIfAssemblyIsStainless(swModel, out string mainDetectedMat);
                    if (!mainIsStainless && !string.IsNullOrWhiteSpace(mainDetectedMat))
                    {
                        mainIsStainless = IsStainlessText(mainDetectedMat);
                    }
                }
                else
                {
                    string mainMat = GetCustomPropValue(swModel, null, "Material");
                    if (string.IsNullOrWhiteSpace(mainMat)) mainMat = GetCustomPropValue(swModel, null, "Materijal");
                    if (string.IsNullOrWhiteSpace(mainMat)) mainMat = GetCustomPropValue(swModel, null, "GS_Materijal");
                    if (!string.IsNullOrWhiteSpace(mainMat) && IsStainlessText(mainMat))
                    {
                        mainIsStainless = true;
                    }
                }
                bool mainHasWeld = false;
                GetModelWeldLength(swModel, mainIsStainless, gpRow, out mainHasWeld);
                if (swModel.GetType() == (int)swDocumentTypes_e.swDocASSEMBLY)
                {
                    ScanAssemblyChildren(swModel, mainIsStainless, gpRow, mainHasWeld);
                }

                SanitizeNutAndFastenerColumns(excelSheet, rowIndex + 1);
                PostaviValidaciju();

                try { excelApp.Calculate(); } catch { }
                try { excelApp.CalculateFull(); } catch { }

                UpisiZbirniSheetSumirano(rowIndex + 1);

                try { excelApp.Calculation = -4105; } catch { } // xlCalculationAutomatic
                try { excelApp.EnableEvents = true; } catch { }
                try { excelApp.ScreenUpdating = true; } catch { }

                try 
                { 
                    excelWorkbook.Save(); 
                } 
                catch 
                { 
                    MessageBox.Show($"Excel fajl '{Path.GetFileName(excelFile)}' je trenutno otvoren ili zaključan od strane drugog korisnika.\n\nZatvorite fajl u Excel-u i pritisnite OK za nastavak.", "Fajl je Otvoren", MessageBoxButton.OK, MessageBoxImage.Warning);
                    try { excelWorkbook.Save(); } catch { }
                }

                try { excelWorkbook.Close(false); } catch { }
                try { excelApp.Quit(); } catch { }

                excelSheet = null;
                excelWorkbook = null;
                excelApp = null;

                GC.Collect();
                GC.WaitForPendingFinalizers();

                MessageBox.Show("Obrada Ponude je uspješno završena.", "Uspjeh", MessageBoxButton.OK, MessageBoxImage.Information);
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Greška u radu sa Excel-om: {ex.Message}", "Greška", MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }
            finally
            {
                try
                {
                    if (swLoadingBox != null && !swLoadingBox.IsDisposed)
                    {
                        swLoadingBox.Close();
                        swLoadingBox.Dispose();
                    }
                }
                catch { }

                if (swHwnd != IntPtr.Zero)
                {
                    try { EnableWindow(swHwnd, true); } catch { }
                }

                excelSheet = null;
                excelWorkbook = null;
                if (excelApp != null)
                {
                    try { excelApp.Quit(); } catch { }
                    excelApp = null;
                }
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        private void ProcessComponents(Component2 swComp)
        {
            if (swComp == null) return;
            try
            {
                if (swComp.IsSuppressed()) return;
            }
            catch { }

            string compPath = "";
            try { compPath = swComp.GetPathName(); } catch { }
            string compConf = "";
            try { compConf = swComp.ReferencedConfiguration; } catch { }

            if (string.IsNullOrWhiteSpace(compPath))
            {
                try { compPath = swComp.Name2; } catch { }
            }

            string cacheKey = $"{compPath}|{compConf}";

            if (!string.IsNullOrWhiteSpace(compPath) && dictVisitedModels.TryGetValue(cacheKey, out string cachedCompKey))
            {
                if (cachedCompKey == "EXCLUDE") return;
                if (!string.IsNullOrWhiteSpace(cachedCompKey) && dictQuantity.ContainsKey(cachedCompKey))
                {
                    dictQuantity[cachedCompKey]++;
                    return;
                }
            }

            bool isPartFile = compPath.EndsWith(".sldprt", StringComparison.OrdinalIgnoreCase);
            if (!isPartFile)
            {
                object vChildObj = swComp.GetChildren();
                if (vChildObj is Array childrenArray)
                {
                    foreach (object child in childrenArray)
                    {
                        Component2 swChild = child as Component2;
                        if (swChild != null)
                        {
                            ProcessComponents(swChild);
                        }
                    }
                }
            }

            ModelDoc2 swModelComp = swComp.GetModelDoc2() as ModelDoc2;
            if (swModelComp == null)
            {
                try
                {
                    swComp.SetSuppression2((int)swComponentSuppressionState_e.swComponentFullyResolved);
                    swModelComp = swComp.GetModelDoc2() as ModelDoc2;
                }
                catch { }
            }

            if (swModelComp != null)
            {
                ProcessPartOrAssembly(swModelComp, swComp.Name2, cacheKey);
            }
        }

        private void ProcessPartOrAssembly(ModelDoc2 swPart, string compName, string cacheKey)
        {
            if (swPart == null) return;

            string gsIdent = GetCustomPropValue(swPart, null, "GS_Ident");
            if (string.IsNullOrWhiteSpace(gsIdent)) gsIdent = GetCustomPropValue(swPart, null, "IDENT");
            if (string.IsNullOrWhiteSpace(gsIdent)) gsIdent = GetCustomPropValue(swPart, null, "Ident");

            string gsNaziv = GetCustomPropValue(swPart, null, "GS_Osnovni_naziv");
            if (string.IsNullOrWhiteSpace(gsNaziv)) gsNaziv = GetCustomPropValue(swPart, null, "Description");
            if (string.IsNullOrWhiteSpace(gsNaziv)) gsNaziv = GetCustomPropValue(swPart, null, "NAZIV");

            string gsDodatniOpis = GetCustomPropValue(swPart, null, "GS_Dodatni_opis");

            string rawName = "";
            try { rawName = Path.GetFileNameWithoutExtension(swPart.GetPathName()); } catch { }
            if (string.IsNullOrWhiteSpace(rawName)) rawName = swPart.GetTitle();
            if (rawName.Contains("/")) rawName = rawName.Split('/')[0];

            string finalName = !string.IsNullOrWhiteSpace(gsDodatniOpis) ? gsDodatniOpis.Trim() : (!string.IsNullOrWhiteSpace(gsNaziv) ? gsNaziv.Trim() : rawName);

            // Automatsko dodjeljivanje ERP Identa za vijačnu robu ukoliko Ident nedostaje
            if (string.IsNullOrWhiteSpace(gsIdent))
            {
                string confName = "";
                try { confName = swPart.ConfigurationManager?.ActiveConfiguration?.Name ?? ""; } catch { }
                if (FastenerDbManager.IsFastener(finalName, confName, gsNaziv))
                {
                    string erpIdent = FastenerDbManager.FindErpIdent(finalName, confName, gsNaziv);
                    if (!string.IsNullOrWhiteSpace(erpIdent))
                    {
                        gsIdent = erpIdent;
                    }
                }
            }

            // ZA PONUDU: NEMA USLOVA DA GS_IDENT MORA BITI POPUNJEN!
            string compKey = !string.IsNullOrWhiteSpace(gsIdent) ? gsIdent : finalName;

            if (!string.IsNullOrWhiteSpace(cacheKey))
            {
                dictVisitedModels[cacheKey] = compKey;
            }

            if (dictQuantity.ContainsKey(compKey))
            {
                dictQuantity[compKey]++;
                return;
            }

            dictQuantity.Add(compKey, 1);
            int targetRow = 0;
            if (dictRowIndex.ContainsKey(compKey))
            {
                targetRow = dictRowIndex[compKey];
            }
            else
            {
                targetRow = rowIndex;
                dictRowIndex.Add(compKey, targetRow);
                rowIndex++;
            }

            string ucVal = GetCustomPropValue(swPart, null, "GS_Upravljacki_centar");
            if (ucVal?.Trim() == "66")
            {
                excelSheet.Cells[targetRow, 132].Value = "PMPR";
            }
            else if (ucVal?.Trim() == "68")
            {
                excelSheet.Cells[targetRow, 132].Value = "NMPR";
            }

            bool isPart = swPart.GetType() == (int)swDocumentTypes_e.swDocPART;
            bool isAssembly = swPart.GetType() == (int)swDocumentTypes_e.swDocASSEMBLY;

            bool isSheetMetal = false;
            double thickness = 0;
            string bendEval = "";
            string flattenLength = "";
            string flattenWidth = "";
            double flattenPerimeter = 0;
            double kFactorVal = 0;
            string kFactorStr = "";
            string materialEval = "";
            string materialForExcel = "";
            bool uvecajPerimetar = false;
            long tCount = 0;
            long cskCount = 0;

            string partTitleLower = (swPart.GetTitle() ?? "").ToLower();

            if (isPart)
            {
                ScanPartFeatures(swPart, partTitleLower,
                    out isSheetMetal, out thickness, out bendEval,
                    out flattenLength, out flattenWidth, out flattenPerimeter,
                    out kFactorVal, out kFactorStr, out materialEval, out materialForExcel,
                    out uvecajPerimetar, out tCount, out cskCount);
            }
            bool asmIsStainless = false;
            if (isAssembly)
            {
                materialForExcel = "--";
                asmIsStainless = DetermineIfAssemblyIsStainless(swPart, out string asmDetectedMat);
                if (asmIsStainless)
                {
                    materialEval = !string.IsNullOrWhiteSpace(asmDetectedMat) ? asmDetectedMat : "1.4301";
                }
                else if (!string.IsNullOrWhiteSpace(asmDetectedMat))
                {
                    materialEval = asmDetectedMat;
                }
            }

            // Određivanje da li je materijal Inox / obojeni metal za perimetar
            string matUpper = (materialEval ?? "").ToUpperInvariant();
            if (IsStainlessText(matUpper) || matUpper.Contains("EN-AW") || matUpper.Contains("ALU") || matUpper.Contains("GALVANIZED STEEL"))
            {
                uvecajPerimetar = true;
            }

            bool isStainless = false;
            if (isAssembly)
            {
                if (asmIsStainless)
                {
                    isStainless = true;
                    uvecajPerimetar = true;
                }
            }
            else if (isPart)
            {
                if (IsStainlessText(materialEval) || IsComponentOrModelStainless(null, swPart))
                {
                    isStainless = true;
                    uvecajPerimetar = true;
                }
            }

            // Zapamti materijal modela kako bi sklop iznad mogao odmah pročitati materijal podsklopa (isključivo pravi partovi/sklopovi, nikad vijci i matice)
            if (!IsFastenerComponent(null, swPart))
            {
                RecordModelMaterial(swPart, finalName, rawName, materialEval, isStainless);
            }

            // Schaumdicht check
            if (finalName.ToLower().Contains("schaumdicht"))
            {
                Array vBox = null;
                if (isPart)
                {
                    PartDoc partDoc = swPart as PartDoc;
                    vBox = partDoc?.GetPartBox(true) as Array;
                }
                else if (isAssembly)
                {
                    AssemblyDoc asmDoc = swPart as AssemblyDoc;
                    vBox = asmDoc?.GetBox(0) as Array;
                }

                if (vBox != null && vBox.Length >= 6)
                {
                    double bX = Math.Abs(Convert.ToDouble(vBox.GetValue(3)) - Convert.ToDouble(vBox.GetValue(0)));
                    double bY = Math.Abs(Convert.ToDouble(vBox.GetValue(4)) - Convert.ToDouble(vBox.GetValue(1)));
                    double bZ = Math.Abs(Convert.ToDouble(vBox.GetValue(5)) - Convert.ToDouble(vBox.GetValue(2)));

                    double[] d = new double[] { bX, bY, bZ };
                    Array.Sort(d);
                    Array.Reverse(d);

                    double schaumdichtVal = Math.Round((d[0] * 2) + (d[1] * 2) + d[2], 3);
                    excelSheet.Cells[targetRow, 107].Value = schaumdichtVal;
                }
            }

            // Upis u Excel
            excelSheet.Cells[targetRow, 1].Value = finalName;
            excelSheet.Cells[targetRow, 2].Value = gsIdent;
            excelSheet.Cells[targetRow, 3].Value = gsNaziv;
            excelSheet.Cells[targetRow, 5].Value = thickness > 0 ? thickness.ToString("0.00") : "0,00";
            excelSheet.Cells[targetRow, 6].Value = flattenLength;
            excelSheet.Cells[targetRow, 7].Value = flattenWidth;
            excelSheet.Cells[targetRow, 67].Value = isSheetMetal ? bendEval : "";
            excelSheet.Cells[targetRow, 134].Value = materialForExcel;

            if (isPart && (isSheetMetal || thickness > 0 || !string.IsNullOrWhiteSpace(bendEval) || kFactorVal > 0 || !string.IsNullOrWhiteSpace(kFactorStr)))
            {
                string valToWrite = "";

                if (kFactorVal > 0)
                {
                    valToWrite = kFactorVal.ToString("0.####", System.Globalization.CultureInfo.InvariantCulture);
                }
                else if (!string.IsNullOrWhiteSpace(kFactorStr) && !kFactorStr.Trim().Equals("0") && !kFactorStr.Trim().Equals("0.00") && !kFactorStr.Trim().Equals("0,00"))
                {
                    valToWrite = kFactorStr.Trim();
                }
                else if (thickness > 0)
                {
                    double validK = KFactorManager.GetKFactor(materialForExcel, thickness);
                    if (validK > 0) valToWrite = validK.ToString("0.####", System.Globalization.CultureInfo.InvariantCulture);
                }

                if (!string.IsNullOrWhiteSpace(valToWrite))
                {
                    if (double.TryParse(valToWrite, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double parsedVal))
                    {
                        excelSheet.Cells[targetRow, 138].Value = parsedVal;
                    }
                    else if (double.TryParse(valToWrite.Replace('.', ','), System.Globalization.NumberStyles.Any, new System.Globalization.CultureInfo("hr-BA"), out double parsedVal2))
                    {
                        excelSheet.Cells[targetRow, 138].Value = parsedVal2;
                    }
                    else
                    {
                        excelSheet.Cells[targetRow, 138].NumberFormat = "@";
                        excelSheet.Cells[targetRow, 138].Value = valToWrite;
                    }
                }
            }

            if (isSheetMetal)
            {
                UpisiPerimetarPoDebljinama(targetRow, thickness, flattenPerimeter, uvecajPerimetar);
            }

            string pzEval = GetCustomPropValue(swPart, null, "Površinska zaštita");
            if (string.IsNullOrWhiteSpace(pzEval)) pzEval = GetCustomPropValue(swPart, null, "Povrsinska zastita");

            string trimmedPz = (pzEval ?? "").Trim();
            if (trimmedPz == "-" || trimmedPz == "--" || trimmedPz == "---" || trimmedPz.Trim('-').Length == 0) trimmedPz = "";

            string finishEval = GetCustomPropValue(swPart, null, "Finish");
            string trimmedFinish = (finishEval ?? "").Trim();
            if (trimmedFinish == "-" || trimmedFinish == "--" || trimmedFinish == "---" || trimmedFinish.Trim('-').Length == 0) trimmedFinish = "";

            if (!string.IsNullOrWhiteSpace(trimmedPz) && !string.IsNullOrWhiteSpace(trimmedFinish))
            {
                if (double.TryParse(trimmedFinish.Replace(',', '.'), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double dFinish))
                {
                    excelSheet.Cells[targetRow, 10].Value = dFinish;
                }
                else
                {
                    excelSheet.Cells[targetRow, 10].Value = trimmedFinish;
                }
            }
            else
            {
                excelSheet.Cells[targetRow, 10].Value = null;
            }

            if (!string.IsNullOrWhiteSpace(trimmedPz))
            {
                excelSheet.Cells[targetRow, 136].Value = trimmedPz;
            }
            else
            {
                excelSheet.Cells[targetRow, 136].Value = null;
            }

            if (isPart)
            {
                if (tCount > 0) excelSheet.Cells[targetRow, 53].Value = tCount;
                if (cskCount > 0) excelSheet.Cells[targetRow, 58].Value = cskCount;
            }

            bool hasWeld = false;
            GetModelWeldLength(swPart, isStainless, targetRow, out hasWeld);

            if (isAssembly)
            {
                ScanAssemblyChildren(swPart, isStainless, targetRow, hasWeld);
            }
        }

        private void ScanPartFeatures(ModelDoc2 swPart, string partTitleLower,
            out bool isSheetMetal, out double thickness, out string bendEval,
            out string flattenLength, out string flattenWidth, out double flattenPerimeter,
            out double kFactorVal, out string kFactorStr, out string materialEval, out string materialForExcel,
            out bool uvecajPerimetar, out long tCount, out long cskCount)
        {
            isSheetMetal = false;
            thickness = 0;
            bendEval = "";
            flattenLength = "";
            flattenWidth = "";
            flattenPerimeter = 0;
            kFactorVal = 0;
            kFactorStr = "";
            materialEval = "";
            materialForExcel = "";
            uvecajPerimetar = false;
            tCount = 0;
            cskCount = 0;

            PartDoc partDoc = swPart as PartDoc;
            if (partDoc != null)
            {
                string dummy = "";
                materialEval = partDoc.GetMaterialPropertyName2("", out dummy) ?? "";
                materialForExcel = materialEval;
            }

            if (string.IsNullOrWhiteSpace(materialEval))
            {
                materialEval = GetCustomPropValue(swPart, null, "Material");
                if (string.IsNullOrWhiteSpace(materialEval)) materialEval = GetCustomPropValue(swPart, null, "Materijal");
                if (string.IsNullOrWhiteSpace(materialEval)) materialEval = GetCustomPropValue(swPart, null, "GS_Materijal");
                if (string.IsNullOrWhiteSpace(materialEval)) materialEval = GetCustomPropValue(swPart, null, "Oznaka_Materijala");
                if (!string.IsNullOrWhiteSpace(materialEval)) materialForExcel = materialEval;
            }

            string matUpper = materialEval.ToUpper();
            if (IsStainlessText(matUpper) || matUpper.Contains("1.4301") || matUpper.Contains("EN-AW") || matUpper.Contains("ALU") || matUpper.Contains("GALVANIZED STEEL"))
            {
                uvecajPerimetar = true;
            }

            bool checkThreads = !partTitleLower.Contains("din") && !partTitleLower.Contains("iso") &&
                                 !partTitleLower.Contains("vijak") && !partTitleLower.Contains("matica");

            bool checkCsk = checkThreads && !partTitleLower.Contains("pem") &&
                            !partTitleLower.Contains("zeger") && !partTitleLower.Contains("podloška");

            Feature swFeat = swPart.FirstFeature() as Feature;
            while (swFeat != null)
            {
                string typeName = swFeat.GetTypeName2() ?? "";
                string typeLower = typeName.ToLower();

                if (typeLower.StartsWith("sheetmetal") || typeLower.StartsWith("sm") || typeLower.Contains("bend") || typeLower.Contains("flange") || typeLower == "flatpattern")
                {
                    isSheetMetal = true;
                    try
                    {
                        dynamic def = swFeat.GetDefinition();
                        if (def != null)
                        {
                            if (thickness == 0)
                            {
                                try { thickness = Convert.ToDouble(def.Thickness) * 1000.0; } catch { }
                            }

                            // 1. Očitavanje K-Faktora sa definicije feature-a
                            if (kFactorVal == 0)
                            {
                                try
                                {
                                    double kVal = Convert.ToDouble(def.KFactor);
                                    if (kVal > 0) kFactorVal = Math.Round(kVal, 4);
                                }
                                catch { }
                            }

                            // 2. Očitavanje Bend Allowance / Bend Deduction sa definicije feature-a
                            if (kFactorVal == 0)
                            {
                                try
                                {
                                    double ba = Convert.ToDouble(def.BendAllowance);
                                    if (ba > 0)
                                    {
                                        double baMm = ba < 0.1 ? ba * 1000.0 : ba;
                                        kFactorVal = Math.Round(baMm, 4);
                                        kFactorStr = kFactorVal.ToString("0.####", System.Globalization.CultureInfo.InvariantCulture);
                                    }
                                }
                                catch { }

                                if (kFactorVal == 0)
                                {
                                    try
                                    {
                                        double bd = Convert.ToDouble(def.BendDeduction);
                                        if (bd > 0)
                                        {
                                            double bdMm = bd < 0.1 ? bd * 1000.0 : bd;
                                            kFactorVal = Math.Round(bdMm, 4);
                                            kFactorStr = kFactorVal.ToString("0.####", System.Globalization.CultureInfo.InvariantCulture);
                                        }
                                    }
                                    catch { }
                                }
                            }

                            // 3. Očitavanje CustomBendAllowance sa definicije feature-a
                            if (kFactorVal == 0)
                            {
                                try
                                {
                                    dynamic cba = def.GetCustomBendAllowance();
                                    if (cba != null)
                                    {
                                        try
                                        {
                                            double ck = Convert.ToDouble(cba.KFactor);
                                            if (ck > 0) kFactorVal = Math.Round(ck, 4);
                                        }
                                        catch { }

                                        if (kFactorVal == 0)
                                        {
                                            try
                                            {
                                                double cbaVal = Convert.ToDouble(cba.BendAllowance);
                                                if (cbaVal > 0)
                                                {
                                                    double baMm = cbaVal < 0.1 ? cbaVal * 1000.0 : cbaVal;
                                                    kFactorVal = Math.Round(baMm, 4);
                                                    kFactorStr = kFactorVal.ToString("0.####", System.Globalization.CultureInfo.InvariantCulture);
                                                }
                                            }
                                            catch { }
                                        }

                                        if (kFactorVal == 0)
                                        {
                                            try
                                            {
                                                double cbdVal = Convert.ToDouble(cba.BendDeduction);
                                                if (cbdVal > 0)
                                                {
                                                    double bdMm = cbdVal < 0.1 ? cbdVal * 1000.0 : cbdVal;
                                                    kFactorVal = Math.Round(bdMm, 4);
                                                    kFactorStr = kFactorVal.ToString("0.####", System.Globalization.CultureInfo.InvariantCulture);
                                                }
                                            }
                                            catch { }
                                        }
                                    }
                                }
                                catch { }
                            }
                        }
                    }
                    catch { }
                }

                if (typeName == "SolidBodyFolder" || typeLower.Contains("bodyfolder") || typeLower.Contains("cutlist"))
                {
                    Feature swSub = swFeat.GetFirstSubFeature() as Feature;
                    while (swSub != null)
                    {
                        string subType = swSub.GetTypeName2() ?? "";
                        if (subType == "CutListFolder" || subType == "SubWeldFolder")
                        {
                            CustomPropertyManager propMgr = swSub.CustomPropertyManager;
                            if (propMgr != null)
                            {
                                string v1 = "", v2 = "";
                                propMgr.Get4("Bends", true, out v1, out v2);
                                bendEval = !string.IsNullOrWhiteSpace(v2) ? v2 : v1;

                                propMgr.Get4("Bounding Box Width", true, out v1, out v2);
                                flattenLength = !string.IsNullOrWhiteSpace(v2) ? v2 : v1;

                                propMgr.Get4("Bounding Box Length", true, out v1, out v2);
                                flattenWidth = !string.IsNullOrWhiteSpace(v2) ? v2 : v1;

                                string[] bendKeys = new string[] {
                                    "Bend Allowance", "K-Factor", "K Factor", "K_Faktor", "KFactor",
                                    "Bend Allowance Standard", "Bend Deduction", "Dodatak za savijanje", "Bend_Allowance"
                                };

                                foreach (string bk in bendKeys)
                                {
                                    if (kFactorVal > 0) break;
                                    string b1 = "", b2 = "";
                                    propMgr.Get4(bk, true, out b1, out b2);
                                    string bEval = !string.IsNullOrWhiteSpace(b2) ? b2 : b1;
                                    if (!string.IsNullOrWhiteSpace(bEval))
                                    {
                                        string cleaned = bEval.ToLower().Replace("mm", "").Trim().Replace(',', '.');
                                        if (double.TryParse(cleaned, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double parsedB) && parsedB > 0)
                                        {
                                            kFactorVal = Math.Round(parsedB, 4);
                                            kFactorStr = kFactorVal.ToString("0.####", System.Globalization.CultureInfo.InvariantCulture);
                                            break;
                                        }
                                        else if (string.IsNullOrWhiteSpace(kFactorStr) && !cleaned.Equals("0") && !cleaned.Equals("0.00") && !cleaned.Equals("0,00"))
                                        {
                                            kFactorStr = bEval.Trim();
                                        }
                                    }
                                }

                                if (kFactorVal == 0 || thickness == 0)
                                {
                                    object pNamesObj = null, pTypesObj = null, pValOutObj = null, pResValOutObj = null, pPropLinkObj = null;
                                    propMgr.GetAll3(ref pNamesObj, ref pTypesObj, ref pValOutObj, ref pResValOutObj, ref pPropLinkObj);
                                    if (pNamesObj is Array propNames && pResValOutObj is Array propResVals)
                                    {
                                        for (int p = 0; p < propNames.Length; p++)
                                        {
                                            string pName = Convert.ToString(propNames.GetValue(p))?.Trim() ?? "";
                                            string pVal = Convert.ToString(propResVals.GetValue(p))?.Trim() ?? "";
                                            if (string.IsNullOrWhiteSpace(pVal) && pValOutObj is Array propVals)
                                            {
                                                pVal = Convert.ToString(propVals.GetValue(p))?.Trim() ?? "";
                                            }

                                            string pNameLower = pName.ToLower();

                                            if (thickness == 0 && (pNameLower.Contains("thick") || pNameLower.Contains("debljin")))
                                            {
                                                string thStr = pVal.ToLower().Replace("mm", "").Trim().Replace(',', '.');
                                                if (double.TryParse(thStr, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double parsedTh) && parsedTh > 0)
                                                {
                                                    thickness = parsedTh;
                                                }
                                            }

                                            if (kFactorVal == 0 && (pNameLower.Contains("k-factor") || pNameLower.Contains("k factor") || pNameLower.Contains("k_faktor") || pNameLower.Contains("kfactor") || pNameLower.Contains("bend allowance") || pNameLower.Contains("bendallowance") || pNameLower.Contains("bend deduction") || pNameLower.Contains("dodatak za savijanje")))
                                            {
                                                string kStr = pVal.ToLower().Replace("mm", "").Trim().Replace(',', '.');
                                                if (double.TryParse(kStr, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double parsedK) && parsedK > 0)
                                                {
                                                    kFactorVal = Math.Round(parsedK, 4);
                                                    kFactorStr = kFactorVal.ToString("0.####", System.Globalization.CultureInfo.InvariantCulture);
                                                }
                                                else if (string.IsNullOrWhiteSpace(kFactorStr) && !kStr.Equals("0") && !kStr.Equals("0.00") && !kStr.Equals("0,00"))
                                                {
                                                    kFactorStr = pVal.Trim();
                                                }
                                            }
                                        }
                                    }
                                }

                                string cutMat1 = "", cutMat2 = "";
                                propMgr.Get4("Material", true, out cutMat1, out cutMat2);
                                string cutMat = !string.IsNullOrWhiteSpace(cutMat2) ? cutMat2 : cutMat1;
                                if (!string.IsNullOrWhiteSpace(cutMat))
                                {
                                    materialEval = cutMat;
                                    materialForExcel = cutMat;
                                    string cutMatUpper = cutMat.ToUpper();
                                    if (IsStainlessText(cutMatUpper) || cutMatUpper.Contains("1.4301") || cutMatUpper.Contains("EN-AW") || cutMatUpper.Contains("ALU") || cutMatUpper.Contains("GALVANIZED STEEL"))
                                    {
                                        uvecajPerimetar = true;
                                    }
                                }

                                string cutOut1 = "", cutOut2 = "", cutIn1 = "", cutIn2 = "";
                                propMgr.Get4("Cutting Length-Outer", true, out cutOut1, out cutOut2);
                                propMgr.Get4("Cutting Length-Inner", true, out cutIn1, out cutIn2);

                                string cutOutEval = !string.IsNullOrWhiteSpace(cutOut2) ? cutOut2 : cutOut1;
                                string cutInEval = !string.IsNullOrWhiteSpace(cutIn2) ? cutIn2 : cutIn1;

                                if (double.TryParse(cutOutEval?.Replace(".", ","), out double valOutCut) &&
                                    double.TryParse(cutInEval?.Replace(".", ","), out double valInCut))
                                {
                                    flattenPerimeter = valOutCut + valInCut;
                                }
                            }
                        }
                        swSub = swSub.GetNextSubFeature() as Feature;
                    }
                }
                else if (checkThreads && (typeName == "WinHole" || typeName == "HoleWzd"))
                {
                    dynamic swHoleData = swFeat.GetDefinition();
                    if (swHoleData != null)
                    {
                        try
                        {
                            int holeType = Convert.ToInt32(swHoleData.Type);
                            if (holeType == 2 || holeType == 3 || holeType == 5)
                            {
                                object refLocs = swHoleData.ReferenceLocations;
                                if (refLocs is Array vPoints)
                                {
                                    tCount += (vPoints.Length > 0) ? (vPoints.Length / 3) : 1;
                                }
                                else
                                {
                                    tCount += 1;
                                }
                            }
                        }
                        catch { }

                        if (checkCsk && swFeat.Name != null && swFeat.Name.ToLower().Contains("csk"))
                        {
                            try
                            {
                                Array vPoints = null;
                                try { vPoints = swHoleData.ReferenceLocations as Array; } catch { }
                                if (vPoints == null || vPoints.Length == 0)
                                {
                                    try { vPoints = swHoleData.GetSketchPoints() as Array; } catch { }
                                }
                                cskCount += (vPoints != null && vPoints.Length > 0) ? vPoints.Length : 1;
                            }
                            catch { cskCount += 1; }
                        }
                    }
                }
                else if (checkThreads && typeName == "CosmeticThread")
                {
                    tCount++;
                }

                if (checkThreads && typeName != "SolidBodyFolder" && typeName != "SheetMetal")
                {
                    Feature swSubFeat = swFeat.GetFirstSubFeature() as Feature;
                    while (swSubFeat != null)
                    {
                        if ((swSubFeat.GetTypeName2() ?? "") == "CosmeticThread")
                        {
                            tCount++;
                        }
                        swSubFeat = swSubFeat.GetNextSubFeature() as Feature;
                    }
                }

                swFeat = swFeat.GetNextFeature() as Feature;
            }

            if (kFactorVal == 0 || thickness == 0)
            {
                List<string> confNames = new List<string>();
                try
                {
                    if (swPart.ConfigurationManager?.ActiveConfiguration != null)
                    {
                        string actName = swPart.ConfigurationManager.ActiveConfiguration.Name;
                        if (!string.IsNullOrWhiteSpace(actName)) confNames.Add(actName);
                    }
                }
                catch { }
                confNames.Add(""); // Global document properties

                foreach (string cName in confNames)
                {
                    if (kFactorVal > 0 && thickness > 0) break;
                    try
                    {
                        CustomPropertyManager partCustPropMgr = swPart.Extension.get_CustomPropertyManager(cName);
                        if (partCustPropMgr != null)
                        {
                            // 1. Direct Get4 check
                            if (kFactorVal == 0)
                            {
                                string[] kPropNames = new string[] {
                                    "K-Factor", "Bend Allowance", "K Factor", "K_Faktor", "KFactor",
                                    "Bend Allowance Standard", "Bend Deduction", "Dodatak za savijanje", "Bend_Allowance"
                                };
                                foreach (string kpn in kPropNames)
                                {
                                    string v1 = "", v2 = "";
                                    partCustPropMgr.Get4(kpn, true, out v1, out v2);
                                    string res = !string.IsNullOrWhiteSpace(v2) ? v2 : v1;
                                    if (!string.IsNullOrWhiteSpace(res))
                                    {
                                        string cleaned = res.ToLower().Replace("mm", "").Trim().Replace(',', '.');
                                        if (double.TryParse(cleaned, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double parsedK) && parsedK > 0)
                                        {
                                            kFactorVal = Math.Round(parsedK, 4);
                                            kFactorStr = kFactorVal.ToString("0.####", System.Globalization.CultureInfo.InvariantCulture);
                                            break;
                                        }
                                        else if (string.IsNullOrWhiteSpace(kFactorStr) && !cleaned.Equals("0") && !cleaned.Equals("0.00") && !cleaned.Equals("0,00"))
                                        {
                                            kFactorStr = res.Trim();
                                        }
                                    }
                                }
                            }

                            // 2. GetAll3 check
                            if (kFactorVal == 0 || thickness == 0)
                            {
                                object pNamesObj = null, pTypesObj = null, pValOutObj = null, pResValOutObj = null, pPropLinkObj = null;
                                partCustPropMgr.GetAll3(ref pNamesObj, ref pTypesObj, ref pValOutObj, ref pResValOutObj, ref pPropLinkObj);
                                if (pNamesObj is Array propNames && pResValOutObj is Array propResVals)
                                {
                                    for (int p = 0; p < propNames.Length; p++)
                                    {
                                        string pName = Convert.ToString(propNames.GetValue(p))?.Trim() ?? "";
                                        string pVal = Convert.ToString(propResVals.GetValue(p))?.Trim() ?? "";
                                        if (string.IsNullOrWhiteSpace(pVal) && pValOutObj is Array propVals)
                                        {
                                            pVal = Convert.ToString(propVals.GetValue(p))?.Trim() ?? "";
                                        }

                                        string pNameLower = pName.ToLower();

                                        if (thickness == 0 && (pNameLower.Contains("thick") || pNameLower.Contains("debljin")))
                                        {
                                            string thStr = pVal.ToLower().Replace("mm", "").Trim().Replace(',', '.');
                                            if (double.TryParse(thStr, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double parsedTh) && parsedTh > 0)
                                            {
                                                thickness = parsedTh;
                                            }
                                        }

                                        if (kFactorVal == 0 && (pNameLower.Contains("k-factor") || pNameLower.Contains("k factor") || pNameLower.Contains("k_faktor") || pNameLower.Contains("kfactor") || pNameLower.Contains("bend allowance") || pNameLower.Contains("bendallowance") || pNameLower.Contains("bend deduction") || pNameLower.Contains("dodatak za savijanje")))
                                        {
                                            string kStr = pVal.ToLower().Replace("mm", "").Trim().Replace(',', '.');
                                            if (double.TryParse(kStr, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double parsedK) && parsedK > 0)
                                            {
                                                kFactorVal = Math.Round(parsedK, 4);
                                                kFactorStr = kFactorVal.ToString("0.####", System.Globalization.CultureInfo.InvariantCulture);
                                            }
                                            else if (string.IsNullOrWhiteSpace(kFactorStr) && !kStr.Equals("0") && !kStr.Equals("0.00") && !kStr.Equals("0,00"))
                                            {
                                                kFactorStr = pVal.Trim();
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                    catch { }
                }
            }

            if (!isSheetMetal && (thickness > 0 || !string.IsNullOrWhiteSpace(bendEval) || !string.IsNullOrWhiteSpace(flattenLength) || kFactorVal > 0 || !string.IsNullOrWhiteSpace(kFactorStr)))
            {
                isSheetMetal = true;
            }
        }

        private void UpisiPerimetarPoDebljinama(int targetRow, double thickness, double perimeter, bool isSpecialMaterial)
        {
            if (perimeter <= 0) return;

            double finalPerimeter = isSpecialMaterial ? perimeter * 1.5 : perimeter;
            int thickRounded = (int)Math.Round(thickness);

            switch (thickRounded)
            {
                case 1: excelSheet.Cells[targetRow, 17].Value = finalPerimeter; break;
                case 2: excelSheet.Cells[targetRow, 23].Value = finalPerimeter; break;
                case 3: excelSheet.Cells[targetRow, 26].Value = finalPerimeter; break;
                case 4: excelSheet.Cells[targetRow, 29].Value = finalPerimeter; break;
                case 5: excelSheet.Cells[targetRow, 35].Value = finalPerimeter; break;
                case 8: excelSheet.Cells[targetRow, 38].Value = finalPerimeter; break;
                case 10: excelSheet.Cells[targetRow, 41].Value = finalPerimeter; break;
                case 12: excelSheet.Cells[targetRow, 44].Value = finalPerimeter; break;
                case 15: excelSheet.Cells[targetRow, 47].Value = finalPerimeter; break;
                case 20: excelSheet.Cells[targetRow, 50].Value = finalPerimeter; break;
                default:
                    if (Math.Abs(thickness - 1.5) < 0.1)
                    {
                        excelSheet.Cells[targetRow, 20].Value = finalPerimeter;
                    }
                    break;
            }
        }

        private void UpisiPodatkeGlavnogSklopa(ModelDoc2 swMainModel, int gpRow)
        {
            if (swMainModel == null) return;
            try
            {
                long mainChildCount = 1;
                double volM3 = 0, areaM2 = 0;

                if (swMainModel.GetType() == (int)swDocumentTypes_e.swDocASSEMBLY)
                {
                    AssemblyDoc mainAsm = swMainModel as AssemblyDoc;
                    if (mainAsm != null)
                    {
                        object vComps = mainAsm.GetComponents(true);
                        if (vComps is Array arrComps)
                        {
                            mainChildCount = arrComps.Length;
                        }
                    }
                }

                // Očitavanje tačnih Mass Properties iz SolidWorks-a (Volume u m3, Area u m2)
                try
                {
                    int status = 0;
                    Array massProps = null;
                    try
                    {
                        massProps = swMainModel.Extension.GetMassProperties2(1, out status, true) as Array;
                    }
                    catch { }

                    if (massProps == null || massProps.Length < 5)
                    {
                        try
                        {
                            int warning = 0;
                            massProps = swMainModel.Extension.GetMassProperties(1, ref warning) as Array;
                        }
                        catch { }
                    }

                    if (massProps != null && massProps.Length >= 5)
                    {
                        volM3 = Convert.ToDouble(massProps.GetValue(3));
                        areaM2 = Convert.ToDouble(massProps.GetValue(4));
                    }
                }
                catch { }

                // Rezervna opcija preko Bounding Box-a ako MassProperties vrati 0
                if (volM3 == 0 && swMainModel.GetType() == (int)swDocumentTypes_e.swDocASSEMBLY)
                {
                    AssemblyDoc mainAsm = swMainModel as AssemblyDoc;
                    if (mainAsm != null)
                    {
                        Array vBox = mainAsm.GetBox(0) as Array;
                        if (vBox != null && vBox.Length >= 6)
                        {
                            double bX = Math.Abs(Convert.ToDouble(vBox.GetValue(3)) - Convert.ToDouble(vBox.GetValue(0)));
                            double bY = Math.Abs(Convert.ToDouble(vBox.GetValue(4)) - Convert.ToDouble(vBox.GetValue(1)));
                            double bZ = Math.Abs(Convert.ToDouble(vBox.GetValue(5)) - Convert.ToDouble(vBox.GetValue(2)));

                            if (bX > 10.0 || bY > 10.0 || bZ > 10.0)
                            {
                                bX /= 1000.0;
                                bY /= 1000.0;
                                bZ /= 1000.0;
                            }

                            volM3 = bX * bY * bZ;
                            if (areaM2 == 0) areaM2 = 2.0 * (bX * bY + bX * bZ + bY * bZ);
                        }
                    }
                }

                excelSheet.Cells[gpRow, 120].Value = mainChildCount > 0 ? mainChildCount : 1;
                if (areaM2 > 0) excelSheet.Cells[gpRow, 121].Value = Math.Round(areaM2, 4);
                if (volM3 > 0) excelSheet.Cells[gpRow, 124].Value = Math.Round(volM3, 6);
            }
            catch { }
        }

        private double GetModelWeldLength(ModelDoc2 targetModel, bool isStainless, int targetRow, out bool hasWeld)
        {
            hasWeld = false;
            if (targetModel == null) return 0;
            ModelDocExtension swModelDocExt = targetModel.Extension;
            SelectionMgr swSelectionMgr = targetModel.SelectionManager as SelectionMgr;

            double suma_2mm = 0;
            double suma_5mm = 0;
            double suma_ostalo = 0;
            HashSet<string> uradjeniFolderi = new HashSet<string>();

            int k = 1;
            bool bStatus = swModelDocExt.SelectByID2($"Weld Bead{k}", "COSMETIC_WELD", 0, 0, 0, false, 0, null, 0);

            while (bStatus)
            {
                hasWeld = true;
                if (swSelectionMgr != null)
                {
                    Feature swFeat = swSelectionMgr.GetSelectedObject6(1, 0) as Feature;
                    if (swFeat != null)
                    {
                        try
                        {
                            dynamic swWeldData = swFeat.GetDefinition();
                            if (swWeldData != null)
                            {
                                dynamic swWeldFolder = swWeldData.GetWeldBeadFolder();
                                if (swWeldFolder != null)
                                {
                                    double totalLength = Convert.ToDouble(swWeldFolder.TotalLength);
                                    string fKey = "F_" + totalLength;

                                    if (!uradjeniFolderi.Contains(fKey))
                                    {
                                        uradjeniFolderi.Add(fKey);
                                        double dbl = Math.Round(Convert.ToDouble(swWeldData.BeadSize) * 1000.0, 0);

                                        if (dbl < 4) suma_2mm += totalLength;
                                        else if (dbl >= 4 && dbl <= 6) suma_5mm += totalLength;
                                        else suma_ostalo += totalLength;
                                    }
                                }
                            }
                        }
                        catch { }
                    }
                }

                k++;
                bStatus = swModelDocExt.SelectByID2($"Weld Bead{k}", "COSMETIC_WELD", 0, 0, 0, false, 0, null, 0);
            }

            try { targetModel.ClearSelection2(true); } catch { }

            double rezultat = (suma_2mm * 1.0) + (suma_5mm * 2.0) + (suma_ostalo * 3.0);
            if (rezultat > 0)
            {
                double rezRound = Math.Round(rezultat, 2);
                if (isStainless)
                {
                    excelSheet.Cells[targetRow, 87].Value2 = rezRound;
                    excelSheet.Cells[targetRow, 74].Value2 = null;
                }
                else
                {
                    excelSheet.Cells[targetRow, 74].Value2 = rezRound;
                    excelSheet.Cells[targetRow, 87].Value2 = null;
                }
            }
            else
            {
                excelSheet.Cells[targetRow, 87].Value2 = null;
                excelSheet.Cells[targetRow, 74].Value2 = null;
            }

            return rezultat;
        }

        private void ScanAssemblyChildren(ModelDoc2 swAsmModel, bool isStainless, int targetRow, bool hasWeld)
        {
            if (swAsmModel == null) return;
            Configuration activeConf = swAsmModel.ConfigurationManager?.ActiveConfiguration;
            if (activeConf == null) return;
            Component2 rootComp = activeConf.GetRootComponent3(true);
            if (rootComp == null) return;

            long vCount = 0;
            long mCount = 0;
            long pCount = 0;
            long pCelik = 0;
            long pInox = 0;
            long wCount = 0;

            object vKidsObj = rootComp.GetChildren();
            if (vKidsObj is Array kidsArray)
            {
                foreach (object kidObj in kidsArray)
                {
                    Component2 childComp = kidObj as Component2;
                    if (childComp == null || childComp.IsSuppressed()) continue;

                    string rawName = childComp.Name2 ?? "";
                    string cPath = "";
                    try { cPath = childComp.GetPathName() ?? ""; } catch { }

                    ModelDoc2 childModel = null;
                    try { childModel = childComp.GetModelDoc2() as ModelDoc2; } catch { }

                    string cTitle = "";
                    string gsNaziv = "";
                    string gsDodatni = "";
                    string gsIdent = "";
                    if (childModel != null)
                    {
                        try { cTitle = childModel.GetTitle() ?? ""; } catch { }
                        try
                        {
                            gsNaziv = GetCustomPropValue(childModel, childComp, "GS_Osnovni_naziv");
                            gsDodatni = GetCustomPropValue(childModel, childComp, "GS_Dodatni_opis");
                            gsIdent = GetCustomPropValue(childModel, childComp, "GS_Ident");
                        }
                        catch { }
                    }

                    string checkStr = $"{rawName} {cPath} {cTitle} {gsNaziv} {gsDodatni} {gsIdent}".ToLower();
                    string nClean = checkStr.Replace(" ", "").Replace("_", "").Replace("-", "");

                    if (IsPressFitNutOrFastener(checkStr))
                    {
                        pCount++;
                    }
                    else if (!nClean.Contains("din929") && !nClean.Contains("iso13918"))
                    {
                        if (IsAssemblyNut(checkStr))
                        {
                            mCount++;
                        }
                        else if (checkStr.Contains("vijak") && !IsPressFitNutOrFastener(checkStr))
                        {
                            vCount++;
                        }
                    }

                    if (!IsFastenerComponent(childComp, childModel) &&
                        !IsStandardOrFastenerText(checkStr, "") &&
                        !IsPressFitNutOrFastener(checkStr))
                    {
                        wCount++;
                    }

                    if (nClean.Contains("din929") || nClean.Contains("iso13918"))
                    {
                        if (nClean.Contains("a2k")) pCelik++;
                        else if (nClean.Contains("a2")) pInox++;
                        else pCelik++;
                    }
                }
            }

            if (vCount > 0) excelSheet.Cells[targetRow, 114].Value = vCount;
            if (pCount > 0) excelSheet.Cells[targetRow, 64].Value = pCount; // Uprešavanje je Kolona 64 (BL)
            if (mCount > 0)
            {
                excelSheet.Cells[targetRow, 115].Value = mCount; // Montaža matice je Kolona 115 (DK)
            }
            else
            {
                excelSheet.Cells[targetRow, 115].Value = null;
            }

            if (hasWeld)
            {
                long weldPrepCount = wCount > 1 ? (wCount - 1) : 1;
                if (isStainless)
                {
                    excelSheet.Cells[targetRow, 86].Value = weldPrepCount;
                    excelSheet.Cells[targetRow, 73].Value = null;
                }
                else
                {
                    excelSheet.Cells[targetRow, 73].Value = weldPrepCount;
                    excelSheet.Cells[targetRow, 86].Value = null;
                }
            }
            else
            {
                excelSheet.Cells[targetRow, 86].Value = null;
                excelSheet.Cells[targetRow, 73].Value = null;
            }

            if (pCelik > 0) excelSheet.Cells[targetRow, 70].Value = pCelik;
            if (pInox > 0) excelSheet.Cells[targetRow, 83].Value = pInox;
        }

        public void SanitizeNutAndFastenerColumns(dynamic sheet, int lastRow)
        {
            if (sheet == null) return;
            try
            {
                dynamic rawData = sheet.Range[$"A10:C{Math.Max(lastRow, 15)}"].Value2;
                if (rawData is object[,] dataArr)
                {
                    int maxR = dataArr.GetLength(0);
                    int emptyCount = 0;

                    for (int r = 1; r <= maxR; r++)
                    {
                        int actualRow = r + 9;
                        string pName = Convert.ToString(dataArr[r, 1])?.Trim() ?? "";
                        string pIdent = Convert.ToString(dataArr[r, 2])?.Trim() ?? "";
                        string pDesc = Convert.ToString(dataArr[r, 3])?.Trim() ?? "";

                        if (string.IsNullOrWhiteSpace(pName) && string.IsNullOrWhiteSpace(pIdent) && string.IsNullOrWhiteSpace(pDesc))
                        {
                            emptyCount++;
                            if (emptyCount >= 3) break;
                            continue;
                        }

                        emptyCount = 0;

                        // Sanitacija $PRP:"SW-File Name" u Koloni A
                        if (pName.Equals("$PRP:\"SW-File Name\"", StringComparison.OrdinalIgnoreCase))
                        {
                            string newColA = !string.IsNullOrWhiteSpace(pDesc) ? pDesc : (!string.IsNullOrWhiteSpace(pIdent) ? pIdent : "");
                            if (!string.IsNullOrWhiteSpace(newColA))
                            {
                                sheet.Cells[actualRow, 1].Value = newColA;
                                pName = newColA;
                            }
                        }

                        // Sanitacija polja Površinska zaštita (Kolona 136 / EJ) i Finish (Kolona 10 / J)
                        object pzVal = sheet.Cells[actualRow, 136].Value;
                        string pzStr = Convert.ToString(pzVal)?.Trim() ?? "";
                        if (pzStr == "-" || pzStr == "--")
                        {
                            pzStr = "";
                            sheet.Cells[actualRow, 136].Value = null;
                        }

                        object finishVal = sheet.Cells[actualRow, 10].Value;
                        string fStr = Convert.ToString(finishVal)?.Trim() ?? "";
                        if (fStr == "-" || fStr == "--")
                        {
                            fStr = "";
                            sheet.Cells[actualRow, 10].Value = null;
                        }

                        // Ako Površinska zaštita ne postoji ili je "-" ili "--", Finish u Koloni 10 ne smije postojati
                        if (string.IsNullOrWhiteSpace(pzStr))
                        {
                            sheet.Cells[actualRow, 10].Value = null;
                        }

                        string checkStr = $"{pIdent} {pName} {pDesc}".ToLower();

                        int qty = 1;
                        object qtyVal = sheet.Cells[actualRow, 4].Value;
                        if (qtyVal != null && int.TryParse(Convert.ToString(qtyVal), out int parsedQty) && parsedQty > 0)
                        {
                            qty = parsedQty;
                        }

                        if (IsPressFitNutOrFastener(checkStr))
                        {
                            sheet.Cells[actualRow, 64].Value = qty;
                            sheet.Cells[actualRow, 115].Value = null;
                            if (!checkStr.Contains("vijak"))
                            {
                                sheet.Cells[actualRow, 114].Value = null;
                            }
                        }
                        else if (IsAssemblyNut(checkStr))
                        {
                            sheet.Cells[actualRow, 115].Value = qty;
                            sheet.Cells[actualRow, 64].Value = null;
                        }
                        else
                        {
                            sheet.Cells[actualRow, 64].Value = null;
                            sheet.Cells[actualRow, 115].Value = null;
                        }
                    }
                }
            }
            catch { }
        }

        private static bool IsStandardOrFastenerText(string text, string matProp = "")
        {
            if (string.IsNullOrWhiteSpace(text)) return false;
            string t = text.ToLowerInvariant();

            // 1. DIN standard (DIN 933, DIN933, DIN-933, DIN_933, itd.)
            if (System.Text.RegularExpressions.Regex.IsMatch(t, @"\bdin\s*[-_.]?\s*\d+") ||
                t.Contains("din ") || t.Contains("din-") || t.Contains("din_") || t.Contains("din.") ||
                t.Contains("\\din\\") || t.Contains("/din/"))
            {
                return true;
            }

            // 2. ISO standard (ISO 4017, ISO4017, ISO-4017, ISO_4017, itd.)
            if (System.Text.RegularExpressions.Regex.IsMatch(t, @"\biso\s*[-_.]?\s*\d+") ||
                t.Contains("iso ") || t.Contains("iso-") || t.Contains("iso_") || t.Contains("iso.") ||
                t.Contains("\\iso\\") || t.Contains("/iso/"))
            {
                return true;
            }

            // 3. EN standard (EN ISO, EN 1661, EN 24017, itd.)
            if (t.Contains("en iso") || t.Contains("en-iso") || t.Contains("en_iso") || t.Contains("eniso") ||
                System.Text.RegularExpressions.Regex.IsMatch(t, @"\ben\s*[-_.]?\s*\d+"))
            {
                return true;
            }

            // 4. UNI, ANSI, ASME, BS
            if (System.Text.RegularExpressions.Regex.IsMatch(t, @"\b(uni|ansi|asme|bs)\s*[-_.]?\s*\d+"))
            {
                return true;
            }

            // 5. Provjera po pojedinačnim riječima (tokenima)
            string[] tokens = t.Split(new char[] { ' ', '_', '-', '.', '/', '\\', '(', ')', '[', ']', ':', ';', ',', '\"', '\'', '+', '=' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (var tok in tokens)
            {
                if (tok == "din" || tok == "iso" || tok == "uni" || tok == "ansi" || tok == "asme")
                {
                    return true;
                }
                if (tok.StartsWith("din") && tok.Length > 3 && char.IsDigit(tok[3])) return true;
                if (tok.StartsWith("iso") && tok.Length > 3 && char.IsDigit(tok[3])) return true;
            }

            // 6. Nazivi vijčane i standardne robe
            if (t.Contains("vijak") || t.Contains("vijci") || t.Contains("vijka") || t.Contains("vijkom") || t.Contains("vijcima") ||
                t.Contains("matica") || t.Contains("matice") || t.Contains("maticu") || t.Contains("maticom") || t.Contains("maticama") ||
                t.Contains("podloska") || t.Contains("podloška") || t.Contains("podloske") || t.Contains("podloške") ||
                t.Contains("podlosku") || t.Contains("podlošku") || t.Contains("podlo") ||
                t.Contains("navojna") || t.Contains("navojne") || t.Contains("navojni") || t.Contains("navojnu") ||
                t.Contains("zakovica") || t.Contains("zakovice") || t.Contains("zakovic") || t.Contains("zakovič") ||
                t.Contains("pop-zakovica") || t.Contains("pop zakovica") ||
                t.Contains("civija") || t.Contains("čivija") || t.Contains("zatik") || t.Contains("stift") || t.Contains("štift") ||
                t.Contains("seger") || t.Contains("zeger") || t.Contains("uskocnik") || t.Contains("uskočnik") ||
                t.Contains("opruga") || t.Contains("feder"))
            {
                return true;
            }

            // 7. Njemački i engleski nazivi
            if (t.Contains("schraube") || t.Contains("mutter") || t.Contains("scheibe") || t.Contains("federring") ||
                t.Contains("blindniet") || t.Contains("gewindestange") ||
                t.Contains("screw") || t.Contains("bolt") || t.Contains("nut") || t.Contains("washer") ||
                t.Contains("rivet") || t.Contains("fastener") || t.Contains("hardware") ||
                t.Contains("standoff") || t.Contains("spacer") || t.Contains("threaded rod") || t.Contains("stud"))
            {
                return true;
            }

            // 8. PEM / uprešavanje
            if (IsPressFitNutOrFastener(t) || IsAssemblyNut(t))
            {
                return true;
            }

            // 9. Putanje standardnih dijelova
            if (t.Contains("\\standard\\") || t.Contains("/standard/") ||
                t.Contains("\\standard parts\\") || t.Contains("/standard parts/") ||
                t.Contains("\\standardni\\") || t.Contains("/standardni/") ||
                t.Contains("\\normteile\\") || t.Contains("/normteile/") ||
                t.Contains("\\vijci\\") || t.Contains("/vijci/") ||
                t.Contains("\\matice\\") || t.Contains("/matice/") ||
                t.Contains("\\toolbox\\") || t.Contains("/toolbox/") ||
                t.Contains("\\fasteners\\") || t.Contains("/fasteners/"))
            {
                return true;
            }

            // 10. Klase čvrstoće vijaka (8.8, 10.9, 12.9, A2-70, A4-70, itd.)
            if (!string.IsNullOrWhiteSpace(matProp))
            {
                string mp = matProp.ToUpperInvariant();
                if (mp == "8.8" || mp == "10.9" || mp == "12.9" || mp == "4.8" || mp == "5.8" ||
                    mp.Contains("8.8") || mp.Contains("10.9") || mp.Contains("12.9") ||
                    mp.Contains("A2-70") || mp.Contains("A4-70") || mp.Contains("A2-50") || mp.Contains("A4-80") ||
                    mp.Contains("A2K"))
                {
                    return true;
                }
            }

            return false;
        }

        private bool IsFastenerComponent(Component2 comp, ModelDoc2 model)
        {
            if (comp == null && model == null) return false;


            string cName = comp?.Name2 ?? "";
            string path = "";
            try { path = comp?.GetPathName() ?? ""; } catch { }
            if (string.IsNullOrWhiteSpace(path) && model != null)
            {
                try { path = model.GetPathName() ?? ""; } catch { }
            }

            string title = "";
            try { title = model?.GetTitle() ?? ""; } catch { }

            string refConf = "";
            try { refConf = comp?.ReferencedConfiguration ?? ""; } catch { }
            if (string.IsNullOrWhiteSpace(refConf) && model != null)
            {
                try { refConf = model.ConfigurationManager?.ActiveConfiguration?.Name ?? ""; } catch { }
            }

            string gsNaziv = "";
            string gsDodatni = "";
            string desc = "";
            string gsIdent = "";
            string standardProp = "";
            string normProp = "";
            string matProp = "";

            if (model != null || comp != null)
            {
                try
                {
                    gsNaziv = GetCustomPropValue(model, comp, "GS_Osnovni_naziv");
                    gsDodatni = GetCustomPropValue(model, comp, "GS_Dodatni_opis");
                    desc = GetCustomPropValue(model, comp, "Description");
                    if (string.IsNullOrWhiteSpace(desc)) desc = GetCustomPropValue(model, comp, "NAZIV");
                    if (string.IsNullOrWhiteSpace(desc)) desc = GetCustomPropValue(model, comp, "Naziv");
                    if (string.IsNullOrWhiteSpace(desc)) desc = GetCustomPropValue(model, comp, "Opis");
                    gsIdent = GetCustomPropValue(model, comp, "GS_Ident");
                    if (string.IsNullOrWhiteSpace(gsIdent)) gsIdent = GetCustomPropValue(model, comp, "IDENT");

                    standardProp = GetCustomPropValue(model, comp, "Standard");
                    if (string.IsNullOrWhiteSpace(standardProp)) standardProp = GetCustomPropValue(model, comp, "STANDARD");

                    normProp = GetCustomPropValue(model, comp, "Norm");
                    if (string.IsNullOrWhiteSpace(normProp)) normProp = GetCustomPropValue(model, comp, "NORM");

                    matProp = GetCustomPropValue(model, comp, "Material");
                    if (string.IsNullOrWhiteSpace(matProp)) matProp = GetCustomPropValue(model, comp, "Materijal");
                }
                catch { }
            }

            string fullCheck = $"{cName} {path} {title} {refConf} {gsNaziv} {gsDodatni} {desc} {gsIdent} {standardProp} {normProp}".ToLowerInvariant();

            if (IsStandardOrFastenerText(fullCheck, matProp))
            {
                return true;
            }

            if (FastenerDbManager.IsFastener(cName, refConf, $"{title} {gsNaziv} {desc}"))
            {
                return true;
            }

            return false;
        }

        private static bool IsPressFitNutOrFastener(string name)
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

            string[] multiLetterPressPrefixes = new string[] {
                "bsos", "bso", "tsos", "tso", "clss", "cls", "fhs", "fha", "kss"
            };

            string[] tokens = n.Split(new char[] { ' ', '_', '-', '.', '<', '>', '/', '\\', '(', ')', ':', ';', '=', '+' }, StringSplitOptions.RemoveEmptyEntries);
            for (int i = 0; i < tokens.Length; i++)
            {
                string t = tokens[i];
                if (t == "pem" || t == "upres" || t == "upresna" || t == "uprešna" || t == "press" || t == "plm" || t == "upresnamatica" || t == "upresna_matica" || t == "upresavanje" || t == "uprešavanje") return true;

                if ((t == "s" || t == "ss" || t == "sp" || t == "so" || t == "ps" || t == "fh" || t == "cls") &&
                    i < tokens.Length - 1 && tokens[i + 1].StartsWith("m") && tokens[i + 1].Length >= 2 && char.IsDigit(tokens[i + 1][1]))
                {
                    return true;
                }

                foreach (var p in multiLetterPressPrefixes)
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

        private static bool IsAssemblyNut(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return false;

            if (IsPressFitNutOrFastener(name)) return false;

            string n = name.ToLower();

            if (n.Contains("din929") || n.Contains("iso13918") || n.Contains("din 929") || n.Contains("iso 13918")) return false;

            string[] tokens = n.Split(new char[] { ' ', '_', '-', '.', '<', '>', '/', '\\', '(', ')', ':', ';', '=' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (var t in tokens)
            {
                if (t == "kvt" || t == "huc" || t == "feks" || t == "riksg" || t == "fef" || t == "hut")
                {
                    return true;
                }
                if (t.StartsWith("kvt") || t.StartsWith("huc") || t.StartsWith("feks") || t.StartsWith("riksg") || t.StartsWith("fef") || t.StartsWith("hut"))
                {
                    return true;
                }
            }
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

        private static bool IsStainlessText(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return false;
            string t = text.ToUpperInvariant();

            if (t.Contains("INOX") ||
                t.Contains("ROSTFREI") ||
                t.Contains("STAINLESS") ||
                t.Contains("EDELSTAHL") ||
                t.Contains("NICHTROSTEND") ||
                t.Contains("NEHRĐAJ") ||
                t.Contains("NEHRDAJ") ||
                t.Contains("NERĐAJ") ||
                t.Contains("NERDJAJ"))
            {
                return true;
            }

            if (t.Contains("V2A") || t.Contains("V4A"))
            {
                return true;
            }

            if (t.Contains("AISI") || t.Contains("SUS"))
            {
                if (t.Contains("304") || t.Contains("316") || t.Contains("301") || t.Contains("303") ||
                    t.Contains("309") || t.Contains("310") || t.Contains("321") || t.Contains("430"))
                {
                    return true;
                }
            }
            if (t.Contains("304L") || t.Contains("316L") || t.Contains("316TI"))
            {
                return true;
            }

            if (t.Contains("Č.4580") || t.Contains("Č.4571") || t.Contains("Č.4570") || t.Contains("Č.4572") ||
                t.Contains("C.4580") || t.Contains("C.4571") || t.Contains("C.4570"))
            {
                return true;
            }

            if (t.Contains("X5CRNI") || t.Contains("X2CRNI") || t.Contains("X6CRNI") ||
                t.Contains("X10CRNI") || t.Contains("X2CRNIMO") || t.Contains("X5CRNIMO") ||
                t.Contains("X6CRNITI") || t.Contains("CRNI") || t.Contains("CR-NI"))
            {
                return true;
            }

            if (System.Text.RegularExpressions.Regex.IsMatch(t, @"\b1\.4[0-9]{3}\b") ||
                t.Contains("1.4301") || t.Contains("1.4404") || t.Contains("1.4571") ||
                t.Contains("1.4305") || t.Contains("1.4306") || t.Contains("1.4307") ||
                t.Contains("1.4401") || t.Contains("1.4016") || t.Contains("1.4541") ||
                t.Contains("1.4845") || t.Contains("1.4828") || t.Contains("1.4462") ||
                t.Contains("1.4539") || t.Contains("1.4310") || t.Contains("1.4000"))
            {
                return true;
            }

            return false;
        }

        private void RecordModelMaterial(ModelDoc2 model, string finalName, string rawName, string material, bool isStainless)
        {
            if (model == null) return;
            string path = "";
            try { path = model.GetPathName() ?? ""; } catch { }
            string title = "";
            try { title = model.GetTitle() ?? ""; } catch { }

            void AddKey(string key)
            {
                if (string.IsNullOrWhiteSpace(key)) return;
                dictModelMaterial[key] = material;
                dictModelIsStainless[key] = isStainless;

                string kLower = key.ToLowerInvariant();
                dictModelMaterial[kLower] = material;
                dictModelIsStainless[kLower] = isStainless;

                string kClean = Path.GetFileNameWithoutExtension(key);
                if (!string.IsNullOrWhiteSpace(kClean))
                {
                    dictModelMaterial[kClean] = material;
                    dictModelIsStainless[kClean] = isStainless;

                    dictModelMaterial[kClean.ToLowerInvariant()] = material;
                    dictModelIsStainless[kClean.ToLowerInvariant()] = isStainless;
                }
            }

            AddKey(path);
            AddKey(title);
            AddKey(finalName);
            AddKey(rawName);
        }

        private bool IsModelRecordedStainless(ModelDoc2 model, string compName = "", string compPath = "")
        {
            bool CheckKey(string k)
            {
                if (string.IsNullOrWhiteSpace(k)) return false;
                if (dictModelIsStainless.TryGetValue(k, out bool s) && s) return true;
                if (dictModelIsStainless.TryGetValue(k.ToLowerInvariant(), out bool s2) && s2) return true;
                string clean = Path.GetFileNameWithoutExtension(k);
                if (!string.IsNullOrWhiteSpace(clean))
                {
                    if (dictModelIsStainless.TryGetValue(clean, out bool s3) && s3) return true;
                    if (dictModelIsStainless.TryGetValue(clean.ToLowerInvariant(), out bool s4) && s4) return true;
                }
                return false;
            }

            if (!string.IsNullOrWhiteSpace(compPath) && CheckKey(compPath)) return true;
            if (!string.IsNullOrWhiteSpace(compName) && CheckKey(compName)) return true;

            if (model != null)
            {
                string path = "";
                try { path = model.GetPathName() ?? ""; } catch { }
                if (CheckKey(path)) return true;

                string title = "";
                try { title = model.GetTitle() ?? ""; } catch { }
                if (CheckKey(title)) return true;
            }

            return false;
        }

        private string GetModelRecordedMaterial(ModelDoc2 model, string compName = "", string compPath = "")
        {
            string CheckKey(string k)
            {
                if (string.IsNullOrWhiteSpace(k)) return "";
                if (dictModelMaterial.TryGetValue(k, out string m) && !string.IsNullOrWhiteSpace(m)) return m;
                if (dictModelMaterial.TryGetValue(k.ToLowerInvariant(), out string m2) && !string.IsNullOrWhiteSpace(m2)) return m2;
                string clean = Path.GetFileNameWithoutExtension(k);
                if (!string.IsNullOrWhiteSpace(clean))
                {
                    if (dictModelMaterial.TryGetValue(clean, out string m3) && !string.IsNullOrWhiteSpace(m3)) return m3;
                    if (dictModelMaterial.TryGetValue(clean.ToLowerInvariant(), out string m4) && !string.IsNullOrWhiteSpace(m4)) return m4;
                }
                return "";
            }

            if (!string.IsNullOrWhiteSpace(compPath))
            {
                string r = CheckKey(compPath);
                if (!string.IsNullOrWhiteSpace(r)) return r;
            }
            if (!string.IsNullOrWhiteSpace(compName))
            {
                string r = CheckKey(compName);
                if (!string.IsNullOrWhiteSpace(r)) return r;
            }

            if (model != null)
            {
                string path = "";
                try { path = model.GetPathName() ?? ""; } catch { }
                string r = CheckKey(path);
                if (!string.IsNullOrWhiteSpace(r)) return r;

                string title = "";
                try { title = model.GetTitle() ?? ""; } catch { }
                r = CheckKey(title);
                if (!string.IsNullOrWhiteSpace(r)) return r;
            }

            return "";
        }

        private bool IsComponentOrModelStainless(Component2 comp, ModelDoc2 model)
        {
            if (model == null && comp == null) return false;

            // STROGO: Ako je fastener/vijak/matica/DIN/ISO, nikad ne smije biti označen kao Inox pozicija za zavarivanje sklopa!
            if (IsFastenerComponent(comp, model)) return false;

            string compPath = "";
            try { compPath = comp?.GetPathName() ?? ""; } catch { }
            string compName = comp?.Name2 ?? "";

            // 1. Ako je već zabilježeno u memoriji
            if (IsModelRecordedStainless(model, compName, compPath)) return true;
            string recMat = GetModelRecordedMaterial(model, compName, compPath);
            if (IsStainlessText(recMat)) return true;

            // 2. SolidWorks materijal
            if (model is PartDoc partDoc)
            {
                string dummy = "";
                string pMat = "";
                try { pMat = partDoc.GetMaterialPropertyName2("", out dummy) ?? ""; } catch { }
                if (IsStainlessText(pMat)) return true;

                if (comp != null && !string.IsNullOrWhiteSpace(comp.ReferencedConfiguration))
                {
                    try { pMat = partDoc.GetMaterialPropertyName2(comp.ReferencedConfiguration, out dummy) ?? ""; } catch { }
                    if (IsStainlessText(pMat)) return true;
                }

                try
                {
                    string activeConf = model.ConfigurationManager?.ActiveConfiguration?.Name;
                    if (!string.IsNullOrWhiteSpace(activeConf))
                    {
                        pMat = partDoc.GetMaterialPropertyName2(activeConf, out dummy) ?? "";
                        if (IsStainlessText(pMat)) return true;
                    }
                }
                catch { }
            }

            // 3. Custom Properties
            string[] propNames = new string[] {
                "Material", "Materijal", "GS_Materijal", "Oznaka_Materijala", "User_Material",
                "GS_Osnovni_naziv", "GS_Dodatni_opis", "Description", "NAZIV", "Naziv"
            };

            foreach (var prop in propNames)
            {
                string val = GetCustomPropValue(model, comp, prop);
                if (IsStainlessText(val)) return true;
            }

            // 4. Naziv, putanja, title
            string path = compPath;
            if (string.IsNullOrWhiteSpace(path) && model != null)
            {
                try { path = model.GetPathName() ?? ""; } catch { }
            }
            string title = "";
            try { title = model?.GetTitle() ?? ""; } catch { }

            string checkText = $"{compName} {path} {title}";
            if (IsStainlessText(checkText)) return true;

            return false;
        }

        private bool DetermineIfAssemblyIsStainless(ModelDoc2 swAsmModel, out string detectedMaterial)
        {
            detectedMaterial = "";
            if (swAsmModel == null) return false;

            // 1. Provera properties/naziva samog sklopa
            string asmTitle = "";
            try { asmTitle = swAsmModel.GetTitle() ?? ""; } catch { }
            if (IsStainlessText(asmTitle))
            {
                detectedMaterial = "1.4301";
                return true;
            }
            string asmMat = GetCustomPropValue(swAsmModel, null, "Material");
            if (string.IsNullOrWhiteSpace(asmMat)) asmMat = GetCustomPropValue(swAsmModel, null, "Materijal");
            if (!string.IsNullOrWhiteSpace(asmMat) && IsStainlessText(asmMat))
            {
                detectedMaterial = asmMat;
                return true;
            }

            AssemblyDoc asmDoc = swAsmModel as AssemblyDoc;
            if (asmDoc == null) return false;

            // 2. Pravilo: Uzmi samo direktne komponente (TopLevelOnly = true) i provjeri prvu pravu poziciju!
            object vCompsObj = asmDoc.GetComponents(true);
            if (vCompsObj is Array compsArray)
            {
                foreach (object c in compsArray)
                {
                    Component2 comp = c as Component2;
                    if (comp == null) continue;
                    try { if (comp.IsSuppressed()) continue; } catch { }

                    ModelDoc2 subCompModel = null;
                    try { subCompModel = comp.GetModelDoc2() as ModelDoc2; } catch { }

                    // Zanemari standardnu i vijčanu robu (DIN, ISO, vijak, matica, PEM...)
                    if (IsFastenerComponent(comp, subCompModel))
                    {
                        continue;
                    }

                    // PRONAĐENA JE PRVA PRAVA POZICIJA!
                    // Pravilo korisnika: "ako ti sklop ima zavarivanje provjeris prvu poziciju,
                    // ako je ona inox onda mora biti zavarivanje inoxa, ako je celik onda moraju i ostale biti celik i mora se upisati u celik zavarivanje"
                    bool isFirstInox = false;

                    if (subCompModel is PartDoc pDoc)
                    {
                        if (IsComponentOrModelStainless(comp, subCompModel))
                        {
                            isFirstInox = true;
                        }
                        else
                        {
                            string dummy = "";
                            string pm = "";
                            try { pm = pDoc.GetMaterialPropertyName2("", out dummy) ?? ""; } catch { }
                            if (string.IsNullOrWhiteSpace(pm))
                            {
                                try
                                {
                                    string cConf = comp.ReferencedConfiguration;
                                    if (!string.IsNullOrWhiteSpace(cConf))
                                        pm = pDoc.GetMaterialPropertyName2(cConf, out dummy) ?? "";
                                }
                                catch { }
                            }
                            if (string.IsNullOrWhiteSpace(pm)) pm = GetCustomPropValue(subCompModel, comp, "Material");
                            if (string.IsNullOrWhiteSpace(pm)) pm = GetCustomPropValue(subCompModel, comp, "Materijal");
                            if (string.IsNullOrWhiteSpace(pm)) pm = GetCustomPropValue(subCompModel, comp, "GS_Materijal");

                            if (!string.IsNullOrWhiteSpace(pm) && IsStainlessText(pm))
                            {
                                isFirstInox = true;
                            }
                        }
                    }
                    else if (subCompModel is AssemblyDoc subAsm)
                    {
                        // Ako je prva pozicija podsklop, provjeri da li je zabeležen kao inox ili uzmi njegovu prvu pravu poziciju
                        string subPath = "";
                        try { subPath = comp.GetPathName() ?? ""; } catch { }
                        if (IsModelRecordedStainless(subCompModel, comp.Name2, subPath))
                        {
                            isFirstInox = true;
                        }
                        else
                        {
                            object vSubComps = subAsm.GetComponents(true);
                            if (vSubComps is Array subArray)
                            {
                                foreach (object sc in subArray)
                                {
                                    Component2 sComp = sc as Component2;
                                    if (sComp == null) continue;
                                    try { if (sComp.IsSuppressed()) continue; } catch { }

                                    ModelDoc2 sModel = null;
                                    try { sModel = sComp.GetModelDoc2() as ModelDoc2; } catch { }
                                    if (IsFastenerComponent(sComp, sModel)) continue;

                                    if (IsComponentOrModelStainless(sComp, sModel))
                                    {
                                        isFirstInox = true;
                                    }
                                    else if (sModel is PartDoc sPartDoc)
                                    {
                                        string dummy = "";
                                        string pm = "";
                                        try { pm = sPartDoc.GetMaterialPropertyName2("", out dummy) ?? ""; } catch { }
                                        if (string.IsNullOrWhiteSpace(pm)) pm = GetCustomPropValue(sModel, sComp, "Material");
                                        if (string.IsNullOrWhiteSpace(pm)) pm = GetCustomPropValue(sModel, sComp, "Materijal");
                                        if (!string.IsNullOrWhiteSpace(pm) && IsStainlessText(pm)) isFirstInox = true;
                                    }
                                    break; // Samo prva prava pozicija podsklopa!
                                }
                            }
                        }
                    }
                    else
                    {
                        if (IsComponentOrModelStainless(comp, subCompModel))
                        {
                            isFirstInox = true;
                        }
                    }

                    if (isFirstInox)
                    {
                        detectedMaterial = "1.4301";
                        return true;
                    }
                    else
                    {
                        detectedMaterial = "Čelik";
                        return false;
                    }
                }
            }

            return false;
        }

        private void UpisiZbirniSheetSumirano(int lastRow)
        {
            try
            {
                if (excelSheet == null || lastRow <= 10) return;

                List<PonudaPartInfo> sheetRowsData = new List<PonudaPartInfo>();

                for (int r = 10; r < lastRow; r++)
                {
                    try
                    {
                        string partName = Convert.ToString(excelSheet.Cells[r, 1].Value)?.Trim() ?? "";
                        string matStr = Convert.ToString(excelSheet.Cells[r, 134].Value)?.Trim() ?? "";

                        object qObj = excelSheet.Cells[r, 4].Value;
                        int qty = 1;
                        if (qObj != null)
                        {
                            if (qObj is double dQ) qty = (int)dQ;
                            else if (qObj is int iQ) qty = iQ;
                            else if (int.TryParse(Convert.ToString(qObj), out int parsedQ)) qty = parsedQ;
                        }

                        object thObj = excelSheet.Cells[r, 5].Value;
                        double thickness = 0.0;
                        if (thObj != null)
                        {
                            if (thObj is double dTh) thickness = dTh;
                            else if (double.TryParse(Convert.ToString(thObj).Replace(',', '.'), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double parsedTh)) thickness = parsedTh;
                        }

                        object mObj = excelSheet.Cells[r, 8].Value;
                        double massSingle = 0.0;
                        if (mObj != null)
                        {
                            if (mObj is double dM) massSingle = dM;
                            else if (double.TryParse(Convert.ToString(mObj).Replace(',', '.'), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double parsedM)) massSingle = parsedM;
                        }

                        if (massSingle == 0 && thickness > 0)
                        {
                            try
                            {
                                object fObj = excelSheet.Cells[r, 6].Value;
                                object gObj = excelSheet.Cells[r, 7].Value;
                                double l = 0, w = 0;
                                double.TryParse(Convert.ToString(fObj)?.Replace(',', '.'), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out l);
                                double.TryParse(Convert.ToString(gObj)?.Replace(',', '.'), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out w);

                                if (l > 0 && w > 0)
                                {
                                    double rho = 7.85e-6; // kg/mm3
                                    string mUpper = (matStr ?? "").ToUpper();
                                    if (mUpper.Contains("INOX") || mUpper.Contains("1.4301") || mUpper.Contains("1.4404")) rho = 7.9e-6;
                                    else if (mUpper.Contains("ALU") || mUpper.Contains("EN-AW")) rho = 2.7e-6;

                                    massSingle = Math.Round(thickness * l * w * rho, 3);
                                }
                            }
                            catch { }
                        }

                        if (thickness > 0)
                        {
                            sheetRowsData.Add(new PonudaPartInfo
                            {
                                PartName = partName,
                                Material = (!string.IsNullOrWhiteSpace(matStr) && matStr != "--" && matStr != "-") ? matStr : "Čelik",
                                Thickness = thickness,
                                MassSingleKg = massSingle,
                                CompKey = r.ToString()
                            });
                        }
                    }
                    catch { }
                }

                if (sheetRowsData.Count == 0) return;

                var grouped = sheetRowsData
                    .GroupBy(p => new {
                        Mat = CleanMaterialName(p.Material),
                        Thick = Math.Round(p.Thickness, 2)
                    })
                    .Select(g => {
                        int totalQty = 0;
                        double totalMass = 0.0;

                        foreach (var item in g)
                        {
                            int rIndex = int.Parse(item.CompKey);
                            int qty = 1;
                            try
                            {
                                string qValStr = Convert.ToString(excelSheet.Cells[rIndex, 4].Value);
                                if (int.TryParse(qValStr, out int parsedQty)) qty = parsedQty;
                            }
                            catch { }

                            totalQty += qty;
                            totalMass += (qty * item.MassSingleKg);
                        }

                        string erpIdent = SheetMetalDbManager.FindRawMaterialIdent(g.Key.Mat, g.Key.Thick);

                        return new
                        {
                            ErpIdent = erpIdent,
                            Material = g.Key.Mat,
                            Thickness = g.Key.Thick,
                            TotalQty = totalQty,
                            TotalMass = Math.Round(totalMass, 3)
                        };
                    })
                    .OrderBy(g => g.Material)
                    .ThenBy(g => g.Thickness)
                    .ToList();

                if (grouped.Count == 0) return;

                dynamic sumiranoSheet = null;
                try
                {
                    sumiranoSheet = excelWorkbook.Sheets["Sumirano"];
                    sumiranoSheet.Cells.Clear();
                }
                catch
                {
                    try
                    {
                        sumiranoSheet = excelWorkbook.Sheets.Add(After: excelWorkbook.Sheets[excelWorkbook.Sheets.Count]);
                        sumiranoSheet.Name = "Sumirano";
                    }
                    catch { }
                }

                if (sumiranoSheet == null) return;

                sumiranoSheet.Cells[1, 1].Value = "REKAPITULACIJA LIMOVA PO MATERIJALU I DEBLJINI (PONUDA " + ident + ")";
                sumiranoSheet.Cells[1, 1].Font.Bold = true;
                sumiranoSheet.Cells[1, 1].Font.Size = 13;

                sumiranoSheet.Cells[3, 1].Value = "R.BR";
                sumiranoSheet.Cells[3, 2].Value = "ERP IDENT SIROVINE";
                sumiranoSheet.Cells[3, 3].Value = "MATERIJAL";
                sumiranoSheet.Cells[3, 4].Value = "DEBLJINA (mm)";
                sumiranoSheet.Cells[3, 5].Value = "UKUPNO KOMADA";
                sumiranoSheet.Cells[3, 6].Value = "UKUPNA MASA (kg)";

                for (int c = 1; c <= 6; c++)
                {
                    sumiranoSheet.Cells[3, c].Font.Bold = true;
                    sumiranoSheet.Cells[3, c].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(55, 71, 79));
                    sumiranoSheet.Cells[3, c].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
                }

                int sRow = 4;
                int rBr = 1;
                int grandTotalQty = 0;
                double grandTotalMass = 0.0;

                foreach (var item in grouped)
                {
                    sumiranoSheet.Cells[sRow, 1].Value = rBr++;
                    sumiranoSheet.Cells[sRow, 2].Value = !string.IsNullOrWhiteSpace(item.ErpIdent) ? item.ErpIdent : "—";
                    sumiranoSheet.Cells[sRow, 3].Value = item.Material;
                    sumiranoSheet.Cells[sRow, 4].Value = item.Thickness.ToString("0.00");
                    sumiranoSheet.Cells[sRow, 5].Value = item.TotalQty;
                    sumiranoSheet.Cells[sRow, 6].Value = item.TotalMass;

                    grandTotalQty += item.TotalQty;
                    grandTotalMass += item.TotalMass;
                    sRow++;
                }

                sumiranoSheet.Cells[sRow, 1].Value = "UKUPNO";
                sumiranoSheet.Cells[sRow, 1].Font.Bold = true;
                sumiranoSheet.Cells[sRow, 5].Value = grandTotalQty;
                sumiranoSheet.Cells[sRow, 5].Font.Bold = true;
                sumiranoSheet.Cells[sRow, 6].Value = Math.Round(grandTotalMass, 3);
                sumiranoSheet.Cells[sRow, 6].Font.Bold = true;

                for (int c = 1; c <= 6; c++)
                {
                    try { sumiranoSheet.Columns[c].AutoFit(); } catch { }
                }

                try { excelSheet.Activate(); } catch { }
            }
            catch { }
        }

        private string CleanMaterialName(string rawMat)
        {
            if (string.IsNullOrWhiteSpace(rawMat)) return "Čelik";
            string m = rawMat.Trim();
            string mUpper = m.ToUpperInvariant();

            if (mUpper.Contains("INOX") || mUpper.Contains("1.4301") || mUpper.Contains("1.4404") || mUpper.Contains("ROSTFREI"))
            {
                return "Inox (" + m + ")";
            }
            if (mUpper.Contains("ALU") || mUpper.Contains("EN-AW"))
            {
                return "Aluminijum (" + m + ")";
            }
            return m;
        }

        private string GetCustomPropValue(ModelDoc2 model, Component2 comp, string propName)
        {
            if (model == null && comp == null) return "";

            string valOut = "", resValOut = "";

            List<string> confsToTry = new List<string>();
            if (comp != null && !string.IsNullOrWhiteSpace(comp.ReferencedConfiguration))
            {
                confsToTry.Add(comp.ReferencedConfiguration);
            }
            if (model != null && model.ConfigurationManager != null && model.ConfigurationManager.ActiveConfiguration != null)
            {
                string activeConf = model.ConfigurationManager.ActiveConfiguration.Name;
                if (!confsToTry.Contains(activeConf)) confsToTry.Add(activeConf);
            }

            if (model != null)
            {
                foreach (string confName in confsToTry)
                {
                    try
                    {
                        var confProp = model.Extension.get_CustomPropertyManager(confName);
                        if (confProp != null)
                        {
                            confProp.Get4(propName, true, out valOut, out resValOut);
                            string res = !string.IsNullOrWhiteSpace(resValOut) ? resValOut : valOut;
                            if (!string.IsNullOrWhiteSpace(res)) return res;
                        }
                    }
                    catch { }
                }

                try
                {
                    var globalProp = model.Extension.get_CustomPropertyManager("");
                    if (globalProp != null)
                    {
                        globalProp.Get4(propName, true, out valOut, out resValOut);
                        string res = !string.IsNullOrWhiteSpace(resValOut) ? resValOut : valOut;
                        if (!string.IsNullOrWhiteSpace(res)) return res;
                    }
                }
                catch { }
            }

            return "";
        }

        public static void PostaviValidacijuOnSheet(dynamic sheet, int lastRow = 300)
        {
            if (sheet == null) return;
            try
            {
                int endRow = Math.Max(300, lastRow);

                dynamic daRange = sheet.Range[$"DA10:DA{endRow}"].Validation;
                daRange.Delete();
                daRange.Add(Type: 3, AlertStyle: 1, Formula1: "=$ER$1000:$EU$1000");
                daRange.IgnoreBlank = true;
                daRange.InCellDropdown = true;

                dynamic cqRange = sheet.Range[$"CQ10:CQ{endRow}"].Validation;
                cqRange.Delete();
                cqRange.Add(Type: 3, AlertStyle: 1, Formula1: "=$ER$999:$ES$999");
                cqRange.IgnoreBlank = true;
                cqRange.InCellDropdown = true;
            }
            catch { }
        }

        private void PostaviValidaciju()
        {
            PostaviValidacijuOnSheet(excelSheet, rowIndex + 1);
        }
    }
}
