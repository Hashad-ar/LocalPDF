using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;

namespace LocalPDF.Services
{
    internal static class WordService
    {
        internal static void Convert(HelperRequest request)
        {
            if (request.Inputs == null || request.Inputs.Count != 1) throw new InvalidOperationException("Exactly one input file is required.");
            string input = System.IO.Path.GetFullPath(request.Inputs[0]);
            string output = System.IO.Path.GetFullPath(request.Output);
            if (!File.Exists(input)) throw new FileNotFoundException("The source file no longer exists.", input);

            var before = WordPids();
            dynamic word = null;
            dynamic documents = null;
            dynamic document = null;
            var owned = new List<Process>();
            using (var job = new WordJob())
            {
                try
                {
                    var wordType = Type.GetTypeFromProgID("Word.Application", throwOnError: false);
                    if (wordType == null) throw new InvalidOperationException("Microsoft Word is not installed. LocalPDF uses your local copy of Word for conversions.");
                    word = Activator.CreateInstance(wordType);
                    word.Visible = false;
                    word.DisplayAlerts = 0;
                    try { word.AutomationSecurity = 3; } catch { }

                    for (int attempt = 0; attempt < 50 && owned.Count == 0; attempt++)
                    {
                        owned = WordPids().Where(pid => !before.Contains(pid)).Select(Process.GetProcessById).ToList();
                        if (owned.Count == 0) Thread.Sleep(100);
                    }
                    if (owned.Count == 0) throw new InvalidOperationException("Word did not create an isolated process, so the conversion was stopped to protect documents already open in Word.");
                    foreach (var process in owned) job.Add(process);
                    var status = new HelperStatus { OwnedWordPids = owned.Select(p => p.Id).ToList(), Phase = "opening" };
                    JsonFiles.Write(request.StatusPath, status);

                    documents = word.Documents;
                    document = documents.Open(input, false, true, false);
                    status.Phase = "saving";
                    JsonFiles.Write(request.StatusPath, status);
                    if (request.Operation == "word-to-pdf")
                        document.ExportAsFixedFormat(output, 17);
                    else if (request.Operation == "pdf-to-word")
                        document.SaveAs2(output, 16, AddToRecentFiles: false);
                    else throw new InvalidOperationException("Unknown conversion request.");

                    status.Phase = "closing";
                    JsonFiles.Write(request.StatusPath, status);
                    document.Close(false);
                    document = null;
                    word.Quit(false);
                    word = null;
                }
                finally
                {
                    try { if (document != null) document.Close(false); } catch { }
                    try { if (word != null) word.Quit(false); } catch { }
                    Release((object)document);
                    Release((object)documents);
                    Release((object)word);
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }
            }
        }

        private static HashSet<int> WordPids()
        {
            return new HashSet<int>(Process.GetProcessesByName("WINWORD").Select(process => process.Id));
        }

        private static void Release(object value)
        {
            try
            {
                if (value != null && Marshal.IsComObject(value)) Marshal.FinalReleaseComObject(value);
            }
            catch { }
        }
    }
}
