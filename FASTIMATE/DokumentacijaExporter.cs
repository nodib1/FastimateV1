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
    public class DokumentacijaExporter
    {
        private SldWorks swApp;
        private ModelDoc2 swModel;
        private string rootPath;
        private string peoFolder;
        private string masterName;

        private bool exportPdf;
        private bool exportDxf;
        private bool exportStep;

        private Dictionary<string, ComponentInfo> compDict;

        // Stanje SolidWorksa prije izvoza – sve što je bilo otvoreno ostaje kako jeste
        private readonly HashSet<string> docsOpenBefore = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private readonly HashSet<string> docsVisibleBefore = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private string activeDocPathBefore = "";

        [DllImport("user32.dll")]
        private static extern bool EnableWindow(IntPtr hWnd, bool bEnable);

        /// <summary>Pamti koji su dokumenti učitani, koji imaju otvoren prozor i koji je aktivan.</summary>
        private void SnapshotOpenDocuments()
        {
            docsOpenBefore.Clear();
            docsVisibleBefore.Clear();
            try { activeDocPathBefore = (swApp.ActiveDoc as ModelDoc2)?.GetPathName() ?? ""; } catch { activeDocPathBefore = ""; }

            try
            {
                if (swApp.GetDocuments() is object[] docs)
                {
                    foreach (object o in docs)
                    {
                        if (!(o is ModelDoc2 doc)) continue;
                        string p = "";
                        try { p = doc.GetPathName() ?? ""; } catch { }
                        if (string.IsNullOrWhiteSpace(p)) continue;
                        docsOpenBefore.Add(p);
                        try { if (doc.Visible) docsVisibleBefore.Add(p); } catch { }
                    }
                }
            }
            catch { }
        }

        /// <summary>
        /// Vraća već otvoren dokument ili ga otvara. openedHere = true znači da ga je izvoz otvorio
        /// i da ga mora zatvoriti čim završi s njim.
        /// </summary>
        private ModelDoc2 OpenForExport(string filePath, int docType, out bool openedHere)
        {
            openedHere = false;
            ModelDoc2 doc = null;
            try { doc = swApp.GetOpenDocumentByName(filePath) as ModelDoc2; } catch { }
            if (doc != null) return doc;

            int err = 0, warn = 0;
            doc = swApp.OpenDoc6(filePath, docType, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", ref err, ref warn);
            openedHere = doc != null;
            return doc;
        }

        /// <summary>
        /// Zatvara prozor/dokument nakon izvoza, ali samo ako ga korisnik nije imao otvorenog prije izvoza.
        /// Za dio koji je učitan kao komponenta sklopa zatvara se samo prozor – dio ostaje u sklopu.
        /// </summary>
        private void CloseIfNotOpenBefore(string filePath, bool openedHere)
        {
            if (string.IsNullOrWhiteSpace(filePath)) return;
            if (string.Equals(filePath, activeDocPathBefore, StringComparison.OrdinalIgnoreCase)) return;
            if (docsVisibleBefore.Contains(filePath)) return; // korisnik ga je imao otvorenog – ne diramo

            // Dokument koji je izvoz otvorio, ili novi prozor za dio koji je bio učitan samo kao komponenta
            if (openedHere || IsDocVisible(filePath))
            {
                try { swApp.CloseDoc(filePath); } catch { }
            }
        }

        private bool IsDocVisible(string filePath)
        {
            try
            {
                ModelDoc2 doc = swApp.GetOpenDocumentByName(filePath) as ModelDoc2;
                return doc != null && doc.Visible;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Na kraju: zatvara sve što je izvoz ostavio otvorenim (novi dokumenti i novi prozori),
        /// a dokumente koji su bili otvoreni prije izvoza ne dira. Vraća aktivni dokument.
        /// </summary>
        private void RestoreDocumentState()
        {
            try
            {
                if (swApp.GetDocuments() is object[] docs)
                {
                    List<string> toClose = new List<string>();
                    foreach (object o in docs)
                    {
                        if (!(o is ModelDoc2 doc)) continue;
                        string p = "";
                        try { p = doc.GetPathName() ?? ""; } catch { }
                        if (string.IsNullOrWhiteSpace(p) || docsVisibleBefore.Contains(p)) continue;
                        if (string.Equals(p, activeDocPathBefore, StringComparison.OrdinalIgnoreCase)) continue;

                        bool visibleNow = false;
                        try { visibleNow = doc.Visible; } catch { }

                        // Novi prozor (dio otvoren radi izvoza) ili dokument koji prije nije bio učitan
                        if (visibleNow || !docsOpenBefore.Contains(p)) toClose.Add(p);
                    }
                    foreach (string p in toClose)
                    {
                        try { swApp.CloseDoc(p); } catch { }
                    }
                }
            }
            catch { }

            if (!string.IsNullOrWhiteSpace(activeDocPathBefore))
            {
                try
                {
                    int err = 0;
                    swApp.ActivateDoc3(activeDocPathBefore, false, (int)swRebuildOnActivation_e.swDontRebuildActiveDoc, ref err);
                }
                catch { }
            }
        }

        private class ComponentInfo
        {
            public string Name { get; set; }
            public string FilePath { get; set; }
            public int Quantity { get; set; }
            public bool IsPart { get; set; }
            public bool IsSheetMetal { get; set; }
            public string Material { get; set; }
            public double ThicknessMm { get; set; }
            public string KontrolaIzvoz { get; set; }
            public ModelDoc2 LoadedModelDoc { get; set; }
        }

        private class DxfRecord
        {
            public string DxfFileName { get; set; }
            public string ThicknessStr { get; set; }
            public int Quantity { get; set; }
            public string Material { get; set; }
        }

        public static bool RunExport(SldWorks activeSwApp, ModelDoc2 activeSwModel, string path, bool pdf, bool dxf, bool step)
        {
            try
            {
                var exporter = new DokumentacijaExporter();
                return exporter.Execute(activeSwApp, activeSwModel, path, pdf, dxf, step);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Greška u pokretanju izvoza dokumentacije: {ex.Message}", "Greška", MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }
        }

        public bool Execute(SldWorks activeSwApp, ModelDoc2 activeSwModel, string path, bool pdf, bool dxf, bool step)
        {
            swApp = activeSwApp;
            swModel = activeSwModel;
            rootPath = path;
            exportPdf = pdf;
            exportDxf = dxf;
            exportStep = step;

            if (swApp == null || swModel == null)
            {
                MessageBox.Show("Nema aktivnog SolidWorks dokumenta.", "Greška", MessageBoxButton.OK, MessageBoxImage.Warning);
                return false;
            }

            SnapshotOpenDocuments();

            peoFolder = Path.Combine(rootPath, "Dokumentacija PEO");
            Directory.CreateDirectory(peoFolder);

            masterName = Path.GetFileNameWithoutExtension(swModel.GetPathName());
            if (string.IsNullOrWhiteSpace(masterName))
            {
                masterName = swModel.GetTitle();
            }

            IntPtr swHwnd = IntPtr.Zero;
            try
            {
                swApp.Visible = true;
                dynamic frame = swApp.IFrameObject();
                if (frame != null) swHwnd = (IntPtr)frame.GetHWnd();
            }
            catch { }

            if (swHwnd == IntPtr.Zero)
            {
                try
                {
                    var procs = System.Diagnostics.Process.GetProcessesByName("SLDWORKS");
                    if (procs.Length > 0) swHwnd = procs[0].MainWindowHandle;
                }
                catch { }
            }

            if (swHwnd != IntPtr.Zero)
            {
                try { EnableWindow(swHwnd, false); } catch { }
            }

            System.Windows.Forms.Form swLoadingBox = ProcesUTokuProzor.Prikazi();

            UserProgressBar swProgress = null;
            try
            {
                swApp.GetUserProgressBar(out swProgress);
                if (swProgress != null)
                {
                    swProgress.Start(0, 100, "Izvoz dokumentacije u toku...");
                    swProgress.UpdateTitle("FASTIMATE - Izvoz Dokumentacije");
                    swProgress.UpdateProgress(5);
                    System.Windows.Forms.Application.DoEvents();
                }
            }
            catch { }

            try
            {
                compDict = new Dictionary<string, ComponentInfo>(StringComparer.OrdinalIgnoreCase);

                Configuration swConf = swModel.ConfigurationManager.ActiveConfiguration;
                Component2 swRootComp = swConf.GetRootComponent3(true);

                // Evidentiramo i sam glavni sklop
                string rootPathName = swModel.GetPathName();
                if (!string.IsNullOrWhiteSpace(rootPathName))
                {
                    ComponentInfo rootInfo = ExtractComponentInfo(swModel, rootPathName, false);
                    rootInfo.Quantity = 1;
                    compDict[rootPathName] = rootInfo;
                }

                if (swRootComp != null)
                {
                    TraverseAssembly(swRootComp);
                }

                if (swProgress != null)
                {
                    try { swProgress.UpdateProgress(30); System.Windows.Forms.Application.DoEvents(); } catch { }
                }

                int totalExportedPdf = 0;
                int totalMissingPdf = 0;
                int totalExportedDxf = 0;
                int totalExportedStep = 0;

                // 1. PDF IZVOZ
                if (exportPdf)
                {
                    ExportPdfDrawings(out totalExportedPdf, out totalMissingPdf);
                }

                if (swProgress != null)
                {
                    try { swProgress.UpdateProgress(60); System.Windows.Forms.Application.DoEvents(); } catch { }
                }

                // 2. DXF IZVOZ (FLAT PATTERN)
                if (exportDxf)
                {
                    ExportDxfFiles(out totalExportedDxf);
                }

                if (swProgress != null)
                {
                    try { swProgress.UpdateProgress(85); System.Windows.Forms.Application.DoEvents(); } catch { }
                }

                // 3. STEP IZVOZ
                if (exportStep)
                {
                    ExportStepFiles(out totalExportedStep);
                }

                if (swProgress != null)
                {
                    try { swProgress.UpdateProgress(100); System.Windows.Forms.Application.DoEvents(); } catch { }
                }

                string poruka = $"Izvoz dokumentacije je uspješno završen!\n\nPEO Mapa: {peoFolder}";
                if (exportPdf)
                {
                    poruka += $"\n- PDF crteži: {totalExportedPdf} kreirano (Nedostajući crteži: {totalMissingPdf})";
                    string mergedFileName = $"{masterName}_ALL.pdf";
                    string mergedFilePath = Path.Combine(peoFolder, "PDF", mergedFileName);
                    if (File.Exists(mergedFilePath))
                    {
                        poruka += $"\n- Spojeni PDF (_ALL): {mergedFileName}";
                    }
                }
                if (exportDxf) poruka += $"\n- DXF fajlovi: {totalExportedDxf} kreirano";
                if (exportStep) poruka += $"\n- STEP modeli: {totalExportedStep} kreirano";

                ProcesUTokuProzor.Zatvori(swLoadingBox);
                MessageBox.Show(poruka, "Izvoz Završen", MessageBoxButton.OK, MessageBoxImage.Information);
                return true;
            }
            catch (Exception ex)
            {
                ProcesUTokuProzor.Zatvori(swLoadingBox);
                MessageBox.Show($"Greška pri izvozu dokumentacije: {ex.Message}", "Greška", MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }
            finally
            {
                ProcesUTokuProzor.Zatvori(swLoadingBox);

                if (swHwnd != IntPtr.Zero)
                {
                    try { EnableWindow(swHwnd, true); } catch { }
                }

                // Zatvara samo ono što je izvoz otvorio; dokumenti otvoreni prije izvoza ostaju, aktivni se vraća
                RestoreDocumentState();

                if (swProgress != null)
                {
                    try { swProgress.End(); } catch { }
                    swProgress = null;
                }
            }
        }

        private void TraverseAssembly(Component2 swComp)
        {
            if (swComp == null || swComp.IsSuppressed()) return;

            ModelDoc2 compModel = swComp.GetModelDoc2() as ModelDoc2;
            if (compModel != null)
            {
                string pathName = compModel.GetPathName();
                if (!string.IsNullOrWhiteSpace(pathName))
                {
                    if (compDict.ContainsKey(pathName))
                    {
                        compDict[pathName].Quantity++;
                    }
                    else
                    {
                        bool isPart = compModel.GetType() == (int)swDocumentTypes_e.swDocPART;
                        ComponentInfo info = ExtractComponentInfo(compModel, pathName, isPart);
                        info.Quantity = 1;
                        compDict[pathName] = info;
                    }
                }
            }

            object vChildren = swComp.GetChildren();
            if (vChildren is Array childrenArr)
            {
                foreach (object childObj in childrenArr)
                {
                    if (childObj is Component2 childComp)
                    {
                        TraverseAssembly(childComp);
                    }
                }
            }
        }

        private ComponentInfo ExtractComponentInfo(ModelDoc2 model, string filePath, bool isPart)
        {
            var info = new ComponentInfo
            {
                Name = Path.GetFileNameWithoutExtension(filePath),
                FilePath = filePath,
                IsPart = isPart,
                LoadedModelDoc = model,
                KontrolaIzvoz = GetCustomProp(model, "GS_Kontrola_Izvoz"),
                Material = "",
                ThicknessMm = 0.0,
                IsSheetMetal = false
            };

            if (string.IsNullOrWhiteSpace(info.KontrolaIzvoz))
            {
                info.KontrolaIzvoz = GetCustomProp(model, "kontrola izvoz");
            }
            if (string.IsNullOrWhiteSpace(info.KontrolaIzvoz))
            {
                info.KontrolaIzvoz = GetCustomProp(model, "Kontrola Izvoz");
            }

            if (string.IsNullOrWhiteSpace(info.KontrolaIzvoz))
            {
                try
                {
                    string activeConfName = model.ConfigurationManager?.ActiveConfiguration?.Name ?? "";
                    if (!string.IsNullOrWhiteSpace(activeConfName))
                    {
                        CustomPropertyManager confProp = model.Extension.get_CustomPropertyManager(activeConfName);
                        string valOut = "", resValOut = "";
                        confProp.Get4("GS_Kontrola_Izvoz", true, out valOut, out resValOut);
                        info.KontrolaIzvoz = !string.IsNullOrWhiteSpace(resValOut) ? resValOut.Trim() : valOut.Trim();
                    }
                }
                catch { }
            }

            if (isPart && model is PartDoc partDoc)
            {
                info.Material = partDoc.GetMaterialPropertyName2("", out _);
                if (string.IsNullOrWhiteSpace(info.Material))
                {
                    info.Material = GetCustomProp(model, "Material");
                }
                if (string.IsNullOrWhiteSpace(info.Material))
                {
                    info.Material = GetCustomProp(model, "Materijal");
                }

                Feature feat = model.FirstFeature() as Feature;
                while (feat != null)
                {
                    string typeName = feat.GetTypeName2();
                    if (typeName == "SheetMetal" || typeName == "SMBaseFlange" || typeName == "FlatPattern" || typeName == "ProcessBends")
                    {
                        info.IsSheetMetal = true;
                        dynamic featData = feat.GetDefinition();
                        if (featData != null)
                        {
                            try
                            {
                                double th = featData.Thickness;
                                if (th > 0) info.ThicknessMm = Math.Round(th * 1000.0, 2);
                            }
                            catch { }
                        }
                    }
                    feat = feat.GetNextFeature() as Feature;
                }
            }

            return info;
        }

        private string GetCustomProp(ModelDoc2 model, string propName)
        {
            if (model == null) return "";
            try
            {
                string activeConf = model.ConfigurationManager?.ActiveConfiguration?.Name ?? "";
                if (!string.IsNullOrWhiteSpace(activeConf))
                {
                    CustomPropertyManager confProp = model.Extension.get_CustomPropertyManager(activeConf);
                    if (confProp != null)
                    {
                        string v1 = "", v2 = "";
                        confProp.Get4(propName, true, out v1, out v2);
                        string res = !string.IsNullOrWhiteSpace(v2) ? v2.Trim() : v1.Trim();
                        if (!string.IsNullOrWhiteSpace(res)) return res;
                    }
                }

                CustomPropertyManager custProp = model.Extension.get_CustomPropertyManager("");
                if (custProp != null)
                {
                    string valOut = "", resValOut = "";
                    custProp.Get4(propName, true, out valOut, out resValOut);
                    return !string.IsNullOrWhiteSpace(resValOut) ? resValOut.Trim() : valOut.Trim();
                }
            }
            catch { }
            return "";
        }

        private string GetDrawingRevision(string partFilePath, ModelDoc2 partModelDoc)
        {
            string revVal = "";

            try
            {
                string slddrwPath = Path.ChangeExtension(partFilePath, ".SLDDRW");
                if (!File.Exists(slddrwPath))
                {
                    slddrwPath = Path.ChangeExtension(partFilePath, ".slddrw");
                }

                if (File.Exists(slddrwPath))
                {
                    ModelDoc2 drwDoc = OpenForExport(slddrwPath, (int)swDocumentTypes_e.swDocDRAWING, out bool openedHere);
                    if (drwDoc != null)
                    {
                        revVal = GetCustomProp(drwDoc, "Revision");
                        if (string.IsNullOrWhiteSpace(revVal)) revVal = GetCustomProp(drwDoc, "revision");
                        if (string.IsNullOrWhiteSpace(revVal)) revVal = GetCustomProp(drwDoc, "REVISION");
                        if (string.IsNullOrWhiteSpace(revVal)) revVal = GetCustomProp(drwDoc, "Rev");
                        if (string.IsNullOrWhiteSpace(revVal)) revVal = GetCustomProp(drwDoc, "GS_Revision");

                        CloseIfNotOpenBefore(drwDoc.GetPathName(), openedHere);
                    }
                }
            }
            catch { }

            if (string.IsNullOrWhiteSpace(revVal) && partModelDoc != null)
            {
                revVal = GetCustomProp(partModelDoc, "Revision");
                if (string.IsNullOrWhiteSpace(revVal)) revVal = GetCustomProp(partModelDoc, "revision");
                if (string.IsNullOrWhiteSpace(revVal)) revVal = GetCustomProp(partModelDoc, "REVISION");
                if (string.IsNullOrWhiteSpace(revVal)) revVal = GetCustomProp(partModelDoc, "Rev");
                if (string.IsNullOrWhiteSpace(revVal)) revVal = GetCustomProp(partModelDoc, "GS_Revision");
            }

            if (string.IsNullOrWhiteSpace(revVal))
            {
                return "-";
            }

            return revVal.Trim();
        }

        private void ExportPdfDrawings(out int exportedCount, out int missingCount)
        {
            exportedCount = 0;
            missingCount = 0;

            string pdfFolder = Path.Combine(peoFolder, "PDF");
            Directory.CreateDirectory(pdfFolder);

            string logPath = Path.Combine(pdfFolder, "Nedostajući_crteži.log");
            List<string> missingLines = new List<string>();

            foreach (var kvp in compDict.Values)
            {
                if (string.Equals(kvp.KontrolaIzvoz, "DA", StringComparison.OrdinalIgnoreCase))
                {
                    string slddrwPath = Path.ChangeExtension(kvp.FilePath, ".SLDDRW");
                    if (!File.Exists(slddrwPath))
                    {
                        slddrwPath = Path.ChangeExtension(kvp.FilePath, ".slddrw");
                    }

                    if (File.Exists(slddrwPath))
                    {
                        try
                        {
                            ModelDoc2 drwDoc = OpenForExport(slddrwPath, (int)swDocumentTypes_e.swDocDRAWING, out bool openedHere);
                            if (drwDoc != null)
                            {
                                string pdfPath = Path.Combine(pdfFolder, $"{kvp.Name}.pdf");
                                int saveErr = 0, saveWarn = 0;
                                try
                                {
                                    drwDoc.Extension.SaveAs(pdfPath, (int)swSaveAsVersion_e.swSaveAsCurrentVersion, (int)swSaveAsOptions_e.swSaveAsOptions_Silent, null, ref saveErr, ref saveWarn);
                                }
                                finally
                                {
                                    CloseIfNotOpenBefore(drwDoc.GetPathName(), openedHere);
                                }
                                exportedCount++;
                            }
                        }
                        catch
                        {
                            missingLines.Add($"{kvp.Name} - GS_Kontrola_Izvoz = DA");
                            missingCount++;
                        }
                    }
                    else
                    {
                        missingLines.Add($"{kvp.Name} - GS_Kontrola_Izvoz = DA");
                        missingCount++;
                    }
                }
            }

            if (missingLines.Count > 0)
            {
                try
                {
                    File.WriteAllLines(logPath, missingLines);
                }
                catch { }
            }

            // AUTOMATSKI MERGE SVIH PDF CRTEŽA U <masterName>_ALL.pdf
            try
            {
                string mergedFileName = $"{masterName}_ALL.pdf";
                string mergedFilePath = Path.Combine(pdfFolder, mergedFileName);

                var pdfFilesToMerge = Directory.GetFiles(pdfFolder, "*.pdf")
                    .Where(f => !Path.GetFileName(f).Equals(mergedFileName, StringComparison.OrdinalIgnoreCase))
                    .ToList();

                if (pdfFilesToMerge.Count > 0)
                {
                    string masterPdfPath = Path.Combine(pdfFolder, $"{masterName}.pdf");
                    pdfFilesToMerge.Sort((a, b) =>
                    {
                        if (string.Equals(a, masterPdfPath, StringComparison.OrdinalIgnoreCase)) return -1;
                        if (string.Equals(b, masterPdfPath, StringComparison.OrdinalIgnoreCase)) return 1;
                        return string.Compare(a, b, StringComparison.OrdinalIgnoreCase);
                    });

                    PdfMerger.MergePdfFiles(pdfFilesToMerge, mergedFilePath);
                }
            }
            catch { }
        }

        private void ExportDxfFiles(out int exportedCount)
        {
            exportedCount = 0;

            string dxfSteel = Path.Combine(peoFolder, "DXF");
            string dxfInox = Path.Combine(peoFolder, "DXF inoks lim");
            string dxfGalv = Path.Combine(peoFolder, "DXF pocinčani lim");
            string dxfOther = Path.Combine(peoFolder, "DXF nerazvrstano");

            var dxfRecords = new Dictionary<string, List<DxfRecord>>(StringComparer.OrdinalIgnoreCase)
            {
                { dxfSteel, new List<DxfRecord>() },
                { dxfInox, new List<DxfRecord>() },
                { dxfGalv, new List<DxfRecord>() },
                { dxfOther, new List<DxfRecord>() }
            };

            foreach (var kvp in compDict.Values)
            {
                if (kvp.IsPart && string.Equals(kvp.KontrolaIzvoz, "DA", StringComparison.OrdinalIgnoreCase))
                {
                    string targetFolder = DetermineDxfFolder(kvp.Material, dxfSteel, dxfInox, dxfGalv, dxfOther);
                    Directory.CreateDirectory(targetFolder);

                    string thStr = kvp.ThicknessMm > 0 ? $"{kvp.ThicknessMm}mm" : "";
                    if (string.IsNullOrWhiteSpace(thStr) && kvp.LoadedModelDoc != null)
                    {
                        thStr = GetCustomProp(kvp.LoadedModelDoc, "Debljina");
                        if (string.IsNullOrWhiteSpace(thStr)) thStr = GetCustomProp(kvp.LoadedModelDoc, "DEBLJINA");
                        if (string.IsNullOrWhiteSpace(thStr)) thStr = GetCustomProp(kvp.LoadedModelDoc, "Thickness");
                    }
                    if (string.IsNullOrWhiteSpace(thStr)) thStr = "0mm";
                    if (!thStr.EndsWith("mm", StringComparison.OrdinalIgnoreCase)) thStr += "mm";

                    string revStr = GetDrawingRevision(kvp.FilePath, kvp.LoadedModelDoc);
                    string dxfFileName = $"{kvp.Name}_{thStr}_Rev.{revStr}.dxf";
                    string dxfFilePath = Path.Combine(targetFolder, dxfFileName);

                    bool openedHere = false;
                    try
                    {
                        ModelDoc2 modelDoc = kvp.LoadedModelDoc;
                        if (modelDoc == null)
                        {
                            modelDoc = OpenForExport(kvp.FilePath, (int)swDocumentTypes_e.swDocPART, out openedHere);
                        }

                        if (modelDoc != null)
                        {
                            int actErr = 0;
                            try { swApp.ActivateDoc3(kvp.FilePath, false, (int)swRebuildOnActivation_e.swDontRebuildActiveDoc, ref actErr); } catch { }

                            if (modelDoc is PartDoc partDoc)
                            {
                                double[] dataAlignment = new double[12];
                                dataAlignment[0] = 0; dataAlignment[1] = 0; dataAlignment[2] = 0;
                                dataAlignment[3] = 1; dataAlignment[4] = 0; dataAlignment[5] = 0;
                                dataAlignment[6] = 0; dataAlignment[7] = 1; dataAlignment[8] = 0;
                                dataAlignment[9] = 0; dataAlignment[10] = 0; dataAlignment[11] = 1;

                                int exportOptions = 1; // 1 = Sheet Metal Flat Pattern Export to DXF/DWG
                                bool success = partDoc.ExportToDWG2(dxfFilePath, kvp.FilePath, exportOptions, true, dataAlignment, false, false, 0, null);

                                if (success || File.Exists(dxfFilePath))
                                {
                                    exportedCount++;
                                    dxfRecords[targetFolder].Add(new DxfRecord
                                    {
                                        DxfFileName = dxfFileName,
                                        ThicknessStr = thStr,
                                        Quantity = kvp.Quantity,
                                        Material = kvp.Material
                                    });
                                }
                            }
                        }
                    }
                    catch { }
                    finally
                    {
                        // Prozor dijela se zatvara odmah – dio ostaje učitan u sklopu
                        CloseIfNotOpenBefore(kvp.FilePath, openedHere);
                    }
                }
            }

            // Ponovo aktiviramo prozor glavnog sklopa
            try
            {
                int resetErr = 0;
                string mainPath = swModel.GetPathName();
                if (!string.IsNullOrWhiteSpace(mainPath))
                {
                    swApp.ActivateDoc3(mainPath, false, 0, ref resetErr);
                }
            }
            catch { }

            // Kreiranje Excel specifikacija u DXF folderima sa novom konvencijom naziva i izgleda
            CreateDxfExcelSpecification(dxfSteel, $"Specifikacija za laser - crni lim - {masterName}.xlsx", $"Specifikacija za laser - {masterName}", dxfRecords[dxfSteel]);
            CreateDxfExcelSpecification(dxfInox, $"Specifikacija za laser - inox lim - {masterName}.xlsx", $"Specifikacija za laser - {masterName}", dxfRecords[dxfInox]);
            CreateDxfExcelSpecification(dxfGalv, $"Specifikacija za laser - pocinčani lim - {masterName}.xlsx", $"Specifikacija za laser - {masterName}", dxfRecords[dxfGalv]);
            CreateDxfExcelSpecification(dxfOther, $"Specifikacija za laser - nerazvrstano - {masterName}.xlsx", $"Specifikacija za laser - {masterName}", dxfRecords[dxfOther]);
        }

        private string DetermineDxfFolder(string material, string steelFolder, string inoxFolder, string galvFolder, string otherFolder)
        {
            if (string.IsNullOrWhiteSpace(material)) return otherFolder;
            string mat = material.ToLowerInvariant().Trim();

            // 1. INOX / STAINLESS
            string[] inoxKeywords = new string[]
            {
                "inox", "inoks", "stainless", "1.4301", "1.4404", "1.4571", "1.4003",
                "1.4305", "1.4307", "1.4401", "1.4541", "1.4462", "304", "316", "316l", "304l", "321", "430"
            };
            foreach (var kw in inoxKeywords)
            {
                if (mat.Contains(kw)) return inoxFolder;
            }

            // 2. POCINČANI / GALVANIZED
            string[] galvKeywords = new string[]
            {
                "pocin", "galvaniz", "galv", "dx51d", "z275", "zinc", "zn", "hdg"
            };
            foreach (var kw in galvKeywords)
            {
                if (mat.Contains(kw)) return galvFolder;
            }

            // 3. ČELIK / CRNI LIM (C45, S235, S355, S231, Plain Carbon Steel, ST-37, C15, C35, DC01, etc.)
            string[] steelKeywords = new string[]
            {
                "čelik", "celik", "steel", "plain carbon steel", "plain carbon", "carbon steel", "mild steel", "low carbon", "high carbon", "cast carbon",
                "c45", "c15", "c35", "c10", "c22", "c55", "c60",
                "s235", "s275", "s355", "s231", "s420", "s460", "s690", "st-37", "st37",
                "st-52", "st52", "st35", "dx51", "1.0037", "1.0038", "1.0570", "1.0503",
                "structural", "carbon", "alloy", "iron", "željezo", "zeljezo", "dc01", "dc02", "dc03", "dc04", "dd11", "dd12", "dd13"
            };
            foreach (var kw in steelKeywords)
            {
                if (mat.Contains(kw)) return steelFolder;
            }

            return otherFolder;
        }

        private string DetermineLimSpecification(string material, string thicknessStr, string dxfFileName, string folderPath)
        {
            double thMm = 0.0;

            if (!string.IsNullOrWhiteSpace(thicknessStr))
            {
                string cleanTh = System.Text.RegularExpressions.Regex.Replace(thicknessStr, @"[^\d.,]", "").Replace(',', '.');
                double.TryParse(cleanTh, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out thMm);
            }

            if (thMm <= 0.0 && !string.IsNullOrWhiteSpace(dxfFileName))
            {
                var match = System.Text.RegularExpressions.Regex.Match(dxfFileName, @"_(\d+(?:[.,]\d+)?)\s*mm", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                if (match.Success)
                {
                    string cleanTh = match.Groups[1].Value.Replace(',', '.');
                    double.TryParse(cleanTh, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out thMm);
                }
            }

            string rawMat = (material ?? "").Trim();
            string matLower = rawMat.ToLowerInvariant();
            string folderName = Path.GetFileName(folderPath ?? "").ToLowerInvariant();

            // 1. INOX -> EN 10088-2-<vrsta_inox_materijala_iz_SW>
            if (folderName.Contains("inoks") || folderName.Contains("inox") ||
                matLower.Contains("inox") || matLower.Contains("inoks") || matLower.Contains("1.4301") || matLower.Contains("1.4404") ||
                matLower.Contains("stainless") || matLower.Contains("304") || matLower.Contains("316"))
            {
                if (rawMat.StartsWith("EN 10088-2", StringComparison.OrdinalIgnoreCase))
                {
                    return rawMat;
                }

                string swInoxMat = rawMat;
                if (string.IsNullOrWhiteSpace(swInoxMat) || matLower.Equals("inox") || matLower.Equals("inoks") || matLower.Equals("stainless"))
                {
                    swInoxMat = "1.4301";
                }

                swInoxMat = swInoxMat.TrimStart('-', ' ');
                return $"EN 10088-2-{swInoxMat}";
            }

            // 2. POCINČANI LIM
            if (folderName.Contains("pocin") || matLower.Contains("pocin") || matLower.Contains("galv") || matLower.Contains("dx51d"))
            {
                return "EN 10346 DX51D+Z275";
            }

            // 3. PLAIN CARBON STEEL (Samo ako je izričito "plain carbon steel") -> mapiraj prema debljini (DC01 ili DD11)
            if (matLower.Contains("plain carbon steel") || matLower.Contains("plain carbon"))
            {
                if (thMm > 0 && thMm < 4.0)
                {
                    return "EN 10130 DC01-A-m";
                }
                else
                {
                    return "EN 10111 DD11";
                }
            }

            // 4. AKO JE DODIJELJEN DRUGI SPECIFIČNI MATERIJAL U SW (npr. S235JR, S355, C45, DC01...) -> UPISUJE SE TAČNO KAKO JE U SW
            if (!string.IsNullOrWhiteSpace(rawMat))
            {
                return rawMat;
            }

            // 5. FALLBACK ZA PRAZAN / NEDETEKTOVAN MATERIJAL -> PREMA DEBLJINI
            if (thMm > 0 && thMm < 4.0)
            {
                return "EN 10130 DC01-A-m";
            }
            else
            {
                return "EN 10111 DD11";
            }
        }

        private void CreateDxfExcelSpecification(string folderPath, string fileName, string titleText, List<DxfRecord> records)
        {
            if (records == null || records.Count == 0) return;

            try
            {
                Directory.CreateDirectory(folderPath);
                string fullExcelPath = Path.Combine(folderPath, fileName);

                Type excelType = Type.GetTypeFromProgID("Excel.Application");
                if (excelType == null) return;

                dynamic excelApp = Activator.CreateInstance(excelType);
                excelApp.Visible = false;
                excelApp.DisplayAlerts = false;

                dynamic wb = excelApp.Workbooks.Add();
                dynamic sheet = wb.Sheets[1];

                // Red 1: Naslov "Specifikacija za laser - NAZIV_PROJEKTA" spojen u A1:D1
                dynamic titleRange = sheet.Range["A1:D1"];
                titleRange.Merge();
                titleRange.Value = titleText;
                titleRange.Font.Bold = true;
                titleRange.Font.Size = 14;
                titleRange.HorizontalAlignment = -4108; // xlCenter
                titleRange.VerticalAlignment = -4108;

                try
                {
                    titleRange.Borders.LineStyle = 1; // xlContinuous
                    titleRange.Borders.Weight = 2;    // xlThin
                }
                catch { }

                // Red 4: Zaglavlja exact po slici korisnika: R.b. | Naziv dxf fajla | Lim | Br.kom.u sklopu
                sheet.Cells[4, 1].Value = "R.b.";
                sheet.Cells[4, 2].Value = "Naziv dxf fajla";
                sheet.Cells[4, 3].Value = "Lim";
                sheet.Cells[4, 4].Value = "Br.kom.u sklopu";

                dynamic headerRange = sheet.Range["A4:D4"];
                headerRange.Font.Bold = true;
                headerRange.Font.Size = 11;
                headerRange.HorizontalAlignment = -4108; // Center header text
                try
                {
                    headerRange.Borders.LineStyle = 1;
                    headerRange.Borders.Weight = 2;
                }
                catch { }

                int row = 5;
                int rBr = 1;
                foreach (var rec in records)
                {
                    string cleanDxfName = rec.DxfFileName ?? "";
                    if (cleanDxfName.EndsWith(".dxf", StringComparison.OrdinalIgnoreCase))
                    {
                        cleanDxfName = cleanDxfName.Substring(0, cleanDxfName.Length - 4);
                    }

                    sheet.Cells[row, 1].Value = rBr++;
                    sheet.Cells[row, 2].Value = cleanDxfName;
                    sheet.Cells[row, 3].Value = DetermineLimSpecification(rec.Material, rec.ThicknessStr, rec.DxfFileName, folderPath);
                    sheet.Cells[row, 4].Value = rec.Quantity;

                    // Centriranje r.b., debljine i količine, naziv dijela ulijevo
                    try
                    {
                        sheet.Cells[row, 1].HorizontalAlignment = -4108;
                        sheet.Cells[row, 2].HorizontalAlignment = -4131; // xlLeft
                        sheet.Cells[row, 3].HorizontalAlignment = -4108;
                        sheet.Cells[row, 4].HorizontalAlignment = -4108;
                    }
                    catch { }

                    row++;
                }

                // Linije okvira za tabelu A4:D{row-1}
                try
                {
                    dynamic tableRange = sheet.Range[$"A4:D{row - 1}"];
                    tableRange.Borders.LineStyle = 1;
                    tableRange.Borders.Weight = 2;
                }
                catch { }

                sheet.Columns["A:D"].AutoFit();

                wb.SaveAs(fullExcelPath);
                wb.Close(false);
                sheet = null;
                wb = null;
                excelApp = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            catch { }
        }

        private void ExportStepFiles(out int exportedCount)
        {
            exportedCount = 0;
            string stepFolder = Path.Combine(peoFolder, "STEP");
            Directory.CreateDirectory(stepFolder);

            foreach (var kvp in compDict.Values)
            {
                string stepFileName = $"{kvp.Name}.step";
                string stepFilePath = Path.Combine(stepFolder, stepFileName);

                bool openedHere = false;
                try
                {
                    ModelDoc2 modelDoc = kvp.LoadedModelDoc;
                    if (modelDoc == null)
                    {
                        int docType = kvp.IsPart ? (int)swDocumentTypes_e.swDocPART : (int)swDocumentTypes_e.swDocASSEMBLY;
                        modelDoc = OpenForExport(kvp.FilePath, docType, out openedHere);
                    }

                    if (modelDoc != null)
                    {
                        int saveErr = 0, saveWarn = 0;
                        bool success = modelDoc.Extension.SaveAs(stepFilePath, (int)swSaveAsVersion_e.swSaveAsCurrentVersion, (int)swSaveAsOptions_e.swSaveAsOptions_Silent, null, ref saveErr, ref saveWarn);

                        if (success || File.Exists(stepFilePath))
                        {
                            exportedCount++;
                        }
                    }
                }
                catch { }
                finally
                {
                    // Zatvara se samo ako ga je izvoz otvorio (glavni sklop i ranije otvoreni dokumenti ostaju)
                    if (openedHere) CloseIfNotOpenBefore(kvp.FilePath, true);
                }
            }
        }
    }
}
