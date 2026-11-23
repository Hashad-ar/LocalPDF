using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.Serialization;

namespace LocalPDF
{
    public sealed class InputFile
    {
        public string Path { get; set; }
        public string Name => System.IO.Path.GetFileName(Path);
        public long Bytes { get; set; }
        public string SizeLabel => FileSize.Format(Bytes);
        public override string ToString() => Name;
    }

    public sealed class HelperRequest
    {
        public string Operation { get; set; }
        public List<string> Inputs { get; set; }
        public string Output { get; set; }
        public string StatusPath { get; set; }
        public string ResultPath { get; set; }
    }

    public sealed class HelperStatus
    {
        public List<int> OwnedWordPids { get; set; } = new List<int>();
        public string Phase { get; set; }
    }

    public sealed class HelperResult
    {
        public bool Success { get; set; }
        public string Error { get; set; }
        public string ErrorCode { get; set; }
    }

    [DataContract]
    public sealed class WordHelperResult
    {
        [DataMember(Name = "ok")]
        public bool Ok { get; set; }
        [DataMember(Name = "code")]
        public string Code { get; set; }
        [DataMember(Name = "detail")]
        public string Detail { get; set; }
    }

    public sealed class OperationResult
    {
        public string OutputPath { get; set; }
        public string Warning { get; set; }
    }

    internal static class FileSize
    {
        public static string Format(long bytes)
        {
            if (bytes < 1024) return bytes + " B";
            if (bytes < 1024 * 1024) return (bytes / 1024d).ToString("0.0") + " KB";
            return (bytes / 1024d / 1024d).ToString("0.0") + " MB";
        }
    }
}
