using System;
using System.Collections.Generic;
using System.IO;
using PdfSharp.Pdf;
using PdfSharp.Pdf.Content;
using PdfSharp.Pdf.Content.Objects;
using PdfSharp.Pdf.IO;

namespace LocalPDF.Services
{
    internal static class PdfService
    {
        internal static void ValidatePdf(string path)
        {
            if (!File.Exists(path)) throw new FileNotFoundException("The PDF file no longer exists.", path);
            using (var stream = File.OpenRead(path))
            {
                var signature = new byte[5];
                if (stream.Read(signature, 0, signature.Length) != signature.Length ||
                    System.Text.Encoding.ASCII.GetString(signature) != "%PDF-")
                    throw new InvalidDataException("This file is not a valid PDF.");
            }
            using (PdfReader.Open(path, PdfDocumentOpenMode.Import)) { }
        }

        internal static void Merge(IReadOnlyList<string> inputs, string output)
        {
            if (inputs == null || inputs.Count < 2) throw new InvalidOperationException("Choose at least two PDF files to merge.");
            using (var document = new PdfDocument())
            {
                foreach (var input in inputs)
                {
                    ValidatePdf(input);
                    using (var source = PdfReader.Open(input, PdfDocumentOpenMode.Import))
                        for (int page = 0; page < source.PageCount; page++) document.AddPage(source.Pages[page]);
                }
                document.Save(output);
            }
            ValidatePdf(output);
        }

        internal static bool AppearsScanned(string path)
        {
            ValidatePdf(path);
            using (var document = PdfReader.Open(path, PdfDocumentOpenMode.Import))
            {
                if (document.PageCount == 0) return false;
                for (int page = 0; page < document.PageCount; page++)
                {
                    CSequence sequence = ContentReader.ReadContent(document.Pages[page]);
                    if (ContainsTextOperator(sequence)) return false;
                }
                return true;
            }
        }

        private static bool ContainsTextOperator(CSequence sequence)
        {
            foreach (CObject item in sequence)
            {
                var operation = item as COperator;
                if (operation != null)
                {
                    string name = operation.OpCode.Name.ToString();
                    if (name == "Tj" || name == "TJ" || name == "QuoteSingle" || name == "QuoteDbl") return true;
                    if (ContainsTextOperator(operation.Operands)) return true;
                }
                var nested = item as CSequence;
                if (nested != null && ContainsTextOperator(nested)) return true;
            }
            return false;
        }
    }
}
