using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

namespace FASTIMATE
{
    public static class PdfMerger
    {
        public static bool MergePdfFiles(List<string> pdfFiles, string outputFile)
        {
            if (pdfFiles == null || pdfFiles.Count == 0 || string.IsNullOrWhiteSpace(outputFile))
                return false;

            var validFiles = pdfFiles.Where(f => File.Exists(f) && new FileInfo(f).Length > 0).ToList();
            if (validFiles.Count == 0) return false;

            if (validFiles.Count == 1)
            {
                try
                {
                    File.Copy(validFiles[0], outputFile, true);
                    return true;
                }
                catch
                {
                    return false;
                }
            }

            try
            {
                using (PdfDocument outputDocument = new PdfDocument())
                {
                    foreach (string file in validFiles)
                    {
                        try
                        {
                            using (PdfDocument inputDocument = PdfReader.Open(file, PdfDocumentOpenMode.Import))
                            {
                                int count = inputDocument.PageCount;
                                for (int idx = 0; idx < count; idx++)
                                {
                                    PdfPage page = inputDocument.Pages[idx];
                                    outputDocument.AddPage(page);
                                }
                            }
                        }
                        catch { }
                    }

                    if (outputDocument.PageCount > 0)
                    {
                        outputDocument.Save(outputFile);
                        return File.Exists(outputFile) && new FileInfo(outputFile).Length > 0;
                    }
                }
                return false;
            }
            catch
            {
                return false;
            }
        }
    }
}
