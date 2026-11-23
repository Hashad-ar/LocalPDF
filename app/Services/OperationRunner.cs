using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace LocalPDF.Services
{
    internal sealed class OperationRunner : IDisposable
    {
        private readonly object _sync = new object();
        private Process _helper;
        private TempWorkspace _workspace;
        private string _statusPath;

        internal event Action<int, string> Progress;

        internal async Task<OperationResult> RunAsync(string operation, IReadOnlyList<string> inputs, string destination, CancellationToken token)
        {
            ValidateInputs(operation, inputs);
            _workspace = new TempWorkspace();
            string extension = operation == "word-to-pdf" || operation == "merge" ? ".pdf" : ".docx";
            string staged = Path.Combine(_workspace.Path, "result" + extension);
            string requestPath = Path.Combine(_workspace.Path, "request.json");
            _statusPath = Path.Combine(_workspace.Path, "status.json");
            string resultPath = Path.Combine(_workspace.Path, "result.json");
            var request = new HelperRequest
            {
                Operation = operation,
                Inputs = inputs.ToList(),
                Output = staged,
                StatusPath = _statusPath,
                ResultPath = resultPath
            };
            JsonFiles.Write(requestPath, request);
            Progress?.Invoke(8, operation == "merge" ? "Preparing PDFs…" : "Starting Microsoft Word…");

            bool isMerge = operation == "merge";
            string executable = isMerge
                ? Process.GetCurrentProcess().MainModule.FileName
                : Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "LocalPDF-WordHelper.exe");
            if (!File.Exists(executable)) throw new FileNotFoundException("The private Microsoft Word helper is missing. Reinstall LocalPDF.", executable);
            string arguments = isMerge
                ? "--merge-helper --request " + Quote(requestPath)
                : "--word-helper " + Quote(operation) + " " + Quote(inputs[0]) + " " + Quote(staged) + " " + Quote(_statusPath) + " " + Quote(resultPath);
            var start = new ProcessStartInfo(executable, arguments)
            {
                UseShellExecute = false,
                CreateNoWindow = true,
                WindowStyle = ProcessWindowStyle.Hidden,
                WorkingDirectory = _workspace.Path
            };
            if (!isMerge)
                start.EnvironmentVariables["LOCALPDF_PREEXISTING_WORD_PIDS"] = string.Join(",", Process.GetProcessesByName("WINWORD").Select(item => item.Id));
            lock (_sync)
            {
                _helper = Process.Start(start);
                if (_helper == null) throw new InvalidOperationException("The private conversion helper could not be started.");
            }

            using (token.Register(Cancel))
            {
                Progress?.Invoke(35, operation == "merge" ? "Combining pages in order…" : "Converting locally…");
                await WaitForExitAsync(_helper).ConfigureAwait(false);
                token.ThrowIfCancellationRequested();
            }

            if (!File.Exists(resultPath)) throw new InvalidOperationException("The helper stopped before returning a result.");
            if (isMerge)
            {
                var result = JsonFiles.Read<HelperResult>(resultPath);
                if (!result.Success) throw new InvalidOperationException(result.Error);
            }
            else
            {
                var result = JsonFiles.Read<WordHelperResult>(resultPath);
                if (!result.Ok) throw new InvalidOperationException(FriendlyWordError(result));
            }
            if (!File.Exists(staged) || new FileInfo(staged).Length == 0) throw new InvalidDataException("The operation produced an empty file.");
            if (extension == ".pdf") PdfService.ValidatePdf(staged);

            Progress?.Invoke(88, "Finishing safely…");
            CommitOutput(staged, destination);
            Progress?.Invoke(100, "Done");
            return new OperationResult { OutputPath = destination };
        }

        internal void Cancel()
        {
            Process helper;
            lock (_sync) helper = _helper;
            try { if (helper != null && !helper.HasExited) helper.Kill(); } catch { }
            KillOwnedWordProcesses();
        }

        private void KillOwnedWordProcesses()
        {
            try
            {
                if (string.IsNullOrWhiteSpace(_statusPath) || !File.Exists(_statusPath)) return;
                string raw = File.ReadAllText(_statusPath).Trim();
                IEnumerable<int> pids;
                if (int.TryParse(raw, out int singlePid)) pids = new[] { singlePid };
                else pids = JsonFiles.Read<HelperStatus>(_statusPath).OwnedWordPids;
                foreach (int pid in pids)
                    try { using (var process = Process.GetProcessById(pid)) if (process.ProcessName.Equals("WINWORD", StringComparison.OrdinalIgnoreCase)) process.Kill(); } catch { }
            }
            catch { }
        }

        private static void ValidateInputs(string operation, IReadOnlyList<string> inputs)
        {
            if (inputs == null || inputs.Count == 0) throw new InvalidOperationException("Choose a file first.");
            if (operation == "merge" && inputs.Count < 2) throw new InvalidOperationException("Choose at least two PDF files to merge.");
            foreach (string input in inputs) if (!File.Exists(input)) throw new FileNotFoundException("A selected file no longer exists.", input);
        }

        private static void CommitOutput(string staged, string destination)
        {
            string directory = Path.GetDirectoryName(destination);
            if (string.IsNullOrWhiteSpace(directory) || !Directory.Exists(directory)) throw new DirectoryNotFoundException("The destination folder no longer exists.");
            string pending = Path.Combine(directory, "." + Path.GetFileName(destination) + ".localpdf-" + Guid.NewGuid().ToString("N") + ".tmp");
            try
            {
                File.Copy(staged, pending, true);
                if (File.Exists(destination)) File.Replace(pending, destination, null);
                else File.Move(pending, destination);
            }
            finally { try { if (File.Exists(pending)) File.Delete(pending); } catch { } }
        }

        private static Task WaitForExitAsync(Process process)
        {
            var completion = new TaskCompletionSource<bool>();
            process.EnableRaisingEvents = true;
            process.Exited += (sender, args) => completion.TrySetResult(true);
            if (process.HasExited) completion.TrySetResult(true);
            return completion.Task;
        }

        private static string Quote(string value) => "\"" + value.Replace("\"", "\\\"") + "\"";

        private static string FriendlyWordError(WordHelperResult result)
        {
            if (result.Code == "word_unavailable") return "Microsoft Word is not installed. LocalPDF uses your local copy of Word for conversions.";
            if (result.Code == "word_open_failed") return "Microsoft Word could not open this file. It may be damaged, password-protected, or unsupported.";
            return string.IsNullOrWhiteSpace(result.Detail) ? "Microsoft Word could not complete the conversion." : result.Detail;
        }

        public void Dispose()
        {
            Cancel();
            lock (_sync) { _helper?.Dispose(); _helper = null; }
            _workspace?.Dispose();
            _workspace = null;
        }
    }
}
