using System;
using System.IO;
using System.Runtime.InteropServices;

namespace LocalPDF.Services
{
    internal static class HelperHost
    {
        internal static int Run(string[] args)
        {
            string requestPath = null;
            for (int i = 0; i < args.Length - 1; i++) if (args[i] == "--request") requestPath = args[i + 1];
            if (string.IsNullOrWhiteSpace(requestPath) || !File.Exists(requestPath)) return 64;
            HelperRequest request = null;
            try
            {
                request = JsonFiles.Read<HelperRequest>(requestPath);
                if (Array.Exists(args, value => value == "--merge-helper")) PdfService.Merge(request.Inputs, request.Output);
                else WordService.Convert(request);
                JsonFiles.Write(request.ResultPath, new HelperResult { Success = true });
                return 0;
            }
            catch (Exception error)
            {
                string diagnosticPath = Environment.GetEnvironmentVariable("LOCALPDF_DIAGNOSTIC_FILE");
                if (!string.IsNullOrWhiteSpace(diagnosticPath))
                    try { File.WriteAllText(diagnosticPath, error.ToString()); } catch { }
                if (request != null && !string.IsNullOrWhiteSpace(request.ResultPath))
                    try { JsonFiles.Write(request.ResultPath, new HelperResult { Success = false, Error = Friendly(error), ErrorCode = error.GetType().Name }); } catch { }
                else
                    try { File.WriteAllText(requestPath + ".error", error.ToString()); } catch { }
                return 1;
            }
        }

        private static string Friendly(Exception error)
        {
            if (error is UnauthorizedAccessException) return "Windows denied access to one of the selected paths. Choose a file and destination you can edit.";
            if (error is InvalidDataException) return error.Message;
            if (error is COMException) return "Microsoft Word could not convert this file. It may be damaged, password-protected, or use an unsupported format.";
            return string.IsNullOrWhiteSpace(error.Message) ? "The operation could not be completed." : error.Message;
        }
    }
}
