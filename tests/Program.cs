using LocalPDF;
using LocalPDF.Services;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using System;
using System.Collections.Generic;
using System.IO;

namespace LocalPDF.Tests
{
    internal static class Program
    {
        private static readonly List<string> Passed = new List<string>();
        private static string _root;

        private static int Main()
        {
            _root = Path.Combine(Path.GetTempPath(), "LocalPDF-tests-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(_root);
            try
            {
                Run("merge preserves exact page order", MergePreservesOrder);
                Run("merge supports Unicode and spaces", MergeSupportsUnicode);
                Run("damaged PDF is rejected", DamagedPdfRejected);
                Run("scanned-page heuristic distinguishes text", ScannedHeuristic);
                Run("stale app temp workspaces are cleaned", StaleCleanup);
                Run("helper JSON preserves Unicode paths", JsonRoundTrip);
                Run("raw helper result JSON is compatible", WordHelperJson);
                Run("file-size labels are stable", FileSizes);
                Console.WriteLine("PASS: " + Passed.Count + "/8 LocalPDF tests");
                foreach (string name in Passed) Console.WriteLine("  + " + name);
                return 0;
            }
            catch (Exception error)
            {
                Console.Error.WriteLine("FAIL: " + error);
                return 1;
            }
            finally { try { Directory.Delete(_root, true); } catch { } }
        }

        private static void Run(string name, Action test) { test(); Passed.Add(name); }

        private static void MergePreservesOrder()
        {
            var inputs = new List<string>();
            foreach (double width in new[] { 101d, 202d, 303d })
            {
                string path = Path.Combine(_root, "page-" + width + ".pdf");
                CreatePdf(path, width, false);
                inputs.Add(path);
            }
            string output = Path.Combine(_root, "ordered.pdf");
            PdfService.Merge(inputs, output);
            using (var merged = PdfReader.Open(output, PdfDocumentOpenMode.Import))
            {
                Equal(3, merged.PageCount, "page count");
                for (int i = 0; i < 3; i++) Near(new[] { 101d, 202d, 303d }[i], merged.Pages[i].Width.Point, "width/order");
            }
        }

        private static void MergeSupportsUnicode()
        {
            string first = Path.Combine(_root, "résumé 一 first.pdf");
            string second = Path.Combine(_root, "spaced second.pdf");
            string output = Path.Combine(_root, "résultat merged 文件.pdf");
            CreatePdf(first, 140, false);
            CreatePdf(second, 150, false);
            PdfService.Merge(new[] { first, second }, output);
            PdfService.ValidatePdf(output);
            True(File.Exists(output), "Unicode output was not created");
        }

        private static void DamagedPdfRejected()
        {
            string path = Path.Combine(_root, "damaged.pdf");
            File.WriteAllText(path, "%PDF-not-a-document");
            bool rejected = false;
            try { PdfService.ValidatePdf(path); } catch { rejected = true; }
            True(rejected, "damaged PDF was accepted");
        }

        private static void ScannedHeuristic()
        {
            string noText = Path.Combine(_root, "scan-like.pdf");
            string text = Path.Combine(_root, "text.pdf");
            CreatePdf(noText, 200, false);
            CreatePdf(text, 200, true);
            True(PdfService.AppearsScanned(noText), "page without text was not flagged");
            True(!PdfService.AppearsScanned(text), "text page was incorrectly flagged");
        }

        private static void StaleCleanup()
        {
            Directory.CreateDirectory(TempWorkspace.Root);
            string stale = Path.Combine(TempWorkspace.Root, "test-stale-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(stale);
            Directory.SetCreationTimeUtc(stale, DateTime.UtcNow.AddDays(-2));
            TempWorkspace.CleanupStale();
            True(!Directory.Exists(stale), "stale workspace remains");
        }

        private static void JsonRoundTrip()
        {
            string path = Path.Combine(_root, "request.json");
            var value = new HelperRequest { Operation = "merge", Inputs = new List<string> { @"C:\文件\résumé one.pdf" }, Output = @"C:\résultat.pdf" };
            JsonFiles.Write(path, value);
            var read = JsonFiles.Read<HelperRequest>(path);
            Equal(value.Inputs[0], read.Inputs[0], "JSON input");
            Equal(value.Output, read.Output, "JSON output");
        }

        private static void FileSizes()
        {
            Equal("999 B", FileSize.Format(999), "bytes");
            Equal("2.0 KB", FileSize.Format(2048), "kilobytes");
            Equal("2.0 MB", FileSize.Format(2 * 1024 * 1024), "megabytes");
        }

        private static void WordHelperJson()
        {
            string path = Path.Combine(_root, "word-result.json");
            File.WriteAllText(path, "{\"ok\":true,\"code\":\"ok\",\"detail\":\"\"}");
            var result = JsonFiles.Read<WordHelperResult>(path);
            True(result.Ok && result.Code == "ok", "raw helper JSON did not deserialize");
        }

        private static void CreatePdf(string path, double width, bool addText)
        {
            using (var document = new PdfDocument())
            {
                PdfPage page = document.AddPage();
                page.Width = XUnit.FromPoint(width);
                page.Height = XUnit.FromPoint(240);
                if (addText)
                {
                    using (var graphics = XGraphics.FromPdfPage(page))
                    {
                        var font = new XFont("Arial", 12);
                        graphics.DrawString("LocalPDF selectable text", font, XBrushes.Black, 12, 30);
                    }
                }
                document.Save(path);
            }
        }

        private static void True(bool condition, string message) { if (!condition) throw new Exception(message); }
        private static void Equal<T>(T expected, T actual, string label) { if (!EqualityComparer<T>.Default.Equals(expected, actual)) throw new Exception(label + ": expected " + expected + ", got " + actual); }
        private static void Near(double expected, double actual, string label) { if (Math.Abs(expected - actual) > 0.01) throw new Exception(label + ": expected " + expected + ", got " + actual); }
    }
}
