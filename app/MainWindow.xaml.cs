using Microsoft.Win32;
using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using LocalPDF.Services;

namespace LocalPDF
{
    public partial class MainWindow : Window, INotifyPropertyChanged
    {
        private readonly ObservableCollection<InputFile> _files = new ObservableCollection<InputFile>();
        private string _operation = "word-to-pdf";
        private OperationRunner _runner;
        private CancellationTokenSource _cancellation;
        private Point _dragStart;
        private string _completedPath;
        private bool _isBusy;

        public bool IsMerge => _operation == "merge";
        public event PropertyChangedEventHandler PropertyChanged;

        public MainWindow()
        {
            InitializeComponent();
            DataContext = this;
            FileList.ItemsSource = _files;
            SelectTool("word-to-pdf");
        }

        private void Window_SourceInitialized(object sender, EventArgs e) => NativeMethods.EnableDarkTitleBar(new WindowInteropHelper(this).Handle);

        private void Tool_Click(object sender, RoutedEventArgs e)
        {
            if (_isBusy) return;
            SelectTool((string)((Button)sender).Tag);
        }

        private void SelectTool(string operation)
        {
            _operation = operation;
            _files.Clear();
            _completedPath = null;
            ShowSelectionState();
            bool merge = IsMerge;
            if (operation == "word-to-pdf")
            {
                ToolTitle.Text = "Word to PDF";
                ToolDescription.Text = "Create a faithful PDF using Microsoft Word on this computer.";
                DropTitle.Text = "Drop a Word file here";
                HintText.Text = ".docx and .doc files • Microsoft Word required";
                RunButton.Content = "Convert to PDF";
            }
            else if (operation == "pdf-to-word")
            {
                ToolTitle.Text = "PDF to Word";
                ToolDescription.Text = "Turn a PDF into an editable Word document using Microsoft Word.";
                DropTitle.Text = "Drop a PDF here";
                HintText.Text = ".pdf files • Microsoft Word required • scanned PDFs may not be editable";
                RunButton.Content = "Convert to Word";
            }
            else
            {
                ToolTitle.Text = "Merge PDFs";
                ToolDescription.Text = "Combine complete PDF pages in the exact order shown below.";
                DropTitle.Text = "Drop two or more PDFs here";
                HintText.Text = "Drag items or use ↑ ↓ to set the page order";
                RunButton.Content = "Merge PDFs";
            }
            AddButton.Visibility = UpButton.Visibility = DownButton.Visibility = merge ? Visibility.Visible : Visibility.Collapsed;
            StyleToolButtons();
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(IsMerge)));
        }

        private void StyleToolButtons()
        {
            foreach (Button button in new[] { WordToPdfButton, PdfToWordButton, MergeButton })
            {
                bool selected = Equals(button.Tag, _operation);
                button.Background = selected ? (Brush)FindResource("AccentBrush") : (Brush)FindResource("ElevatedBrush");
                button.BorderBrush = selected ? (Brush)FindResource("AccentBrush") : (Brush)FindResource("BorderBrush");
            }
        }

        private void Browse_Click(object sender, RoutedEventArgs e)
        {
            if (_isBusy) return;
            var dialog = new OpenFileDialog
            {
                Multiselect = IsMerge,
                CheckFileExists = true,
                Filter = _operation == "word-to-pdf" ? "Word documents (*.docx;*.doc)|*.docx;*.doc" : "PDF files (*.pdf)|*.pdf"
            };
            if (dialog.ShowDialog(this) == true) AddPaths(dialog.FileNames);
        }

        private void Window_DragEnter(object sender, DragEventArgs e)
        {
            e.Effects = !_isBusy && e.Data.GetDataPresent(DataFormats.FileDrop) ? DragDropEffects.Copy : DragDropEffects.None;
            e.Handled = true;
        }

        private void Window_Drop(object sender, DragEventArgs e)
        {
            if (_isBusy || !e.Data.GetDataPresent(DataFormats.FileDrop)) return;
            AddPaths((string[])e.Data.GetData(DataFormats.FileDrop));
            e.Handled = true;
        }

        private void AddPaths(string[] paths)
        {
            string[] allowed = _operation == "word-to-pdf" ? new[] { ".doc", ".docx" } : new[] { ".pdf" };
            var valid = paths.Where(File.Exists).Where(path => allowed.Contains(Path.GetExtension(path), StringComparer.OrdinalIgnoreCase)).ToArray();
            if (valid.Length == 0)
            {
                MessageBox.Show(this, _operation == "word-to-pdf" ? "Choose a .docx or .doc file." : "Choose a PDF file.", "Unsupported file", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }
            if (!IsMerge) _files.Clear();
            foreach (string path in valid)
            {
                string fullPath = Path.GetFullPath(path);
                if (_files.All(item => !item.Path.Equals(fullPath, StringComparison.OrdinalIgnoreCase)))
                    _files.Add(new InputFile { Path = fullPath, Bytes = new FileInfo(fullPath).Length });
                if (!IsMerge) break;
            }
            ShowSelectionState();
        }

        private void ShowSelectionState()
        {
            DropZone.Visibility = _files.Count == 0 ? Visibility.Visible : Visibility.Collapsed;
            SelectionPanel.Visibility = _files.Count > 0 ? Visibility.Visible : Visibility.Collapsed;
            ProgressPanel.Visibility = DonePanel.Visibility = Visibility.Collapsed;
            RunButton.Visibility = HintText.Visibility = Visibility.Visible;
            RunButton.IsEnabled = IsMerge ? _files.Count >= 2 : _files.Count == 1;
        }

        private void Remove_Click(object sender, RoutedEventArgs e)
        {
            if (FileList.SelectedItem is InputFile selected) _files.Remove(selected);
            ShowSelectionState();
        }

        private void Clear_Click(object sender, RoutedEventArgs e) { _files.Clear(); ShowSelectionState(); }
        private void MoveUp_Click(object sender, RoutedEventArgs e) => MoveSelected(-1);
        private void MoveDown_Click(object sender, RoutedEventArgs e) => MoveSelected(1);

        private void MoveSelected(int delta)
        {
            int oldIndex = FileList.SelectedIndex;
            int newIndex = oldIndex + delta;
            if (oldIndex < 0 || newIndex < 0 || newIndex >= _files.Count) return;
            _files.Move(oldIndex, newIndex);
            FileList.SelectedIndex = newIndex;
        }

        private void FileList_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e) => _dragStart = e.GetPosition(null);

        private void FileList_PreviewMouseMove(object sender, MouseEventArgs e)
        {
            if (!IsMerge || e.LeftButton != MouseButtonState.Pressed) return;
            Point current = e.GetPosition(null);
            if (Math.Abs(current.X - _dragStart.X) < SystemParameters.MinimumHorizontalDragDistance && Math.Abs(current.Y - _dragStart.Y) < SystemParameters.MinimumVerticalDragDistance) return;
            if (FileList.SelectedItem is InputFile item) DragDrop.DoDragDrop(FileList, item, DragDropEffects.Move);
        }

        private void FileList_Drop(object sender, DragEventArgs e)
        {
            if (!IsMerge || !(e.Data.GetData(typeof(InputFile)) is InputFile source)) return;
            var element = FileList.InputHitTest(e.GetPosition(FileList)) as DependencyObject;
            while (element != null && !(element is ListBoxItem)) element = VisualTreeHelper.GetParent(element);
            if (!(element is ListBoxItem targetContainer) || !(targetContainer.DataContext is InputFile target)) return;
            int targetIndex = _files.IndexOf(target);
            int sourceIndex = _files.IndexOf(source);
            if (sourceIndex >= 0 && targetIndex >= 0 && sourceIndex != targetIndex) _files.Move(sourceIndex, targetIndex);
        }

        private async void Run_Click(object sender, RoutedEventArgs e)
        {
            if (_isBusy || !RunButton.IsEnabled) return;
            string destination = AskDestination();
            if (destination == null) return;

            if (_operation == "pdf-to-word")
            {
                try
                {
                    if (await Task.Run(() => PdfService.AppearsScanned(_files[0].Path)))
                    {
                        var proceed = MessageBox.Show(this, "This PDF appears to contain scanned pages without selectable text. Word may preserve the pages as images, but the text may not be editable. Continue?",
                            "Scanned PDF detected", MessageBoxButton.YesNo, MessageBoxImage.Warning);
                        if (proceed != MessageBoxResult.Yes) return;
                    }
                }
                catch (Exception error) { ShowError(error.Message); return; }
            }

            _isBusy = true;
            SetToolButtonsEnabled(false);
            SelectionPanel.Visibility = DropZone.Visibility = Visibility.Collapsed;
            ProgressPanel.Visibility = Visibility.Visible;
            DonePanel.Visibility = Visibility.Collapsed;
            RunButton.Visibility = HintText.Visibility = Visibility.Collapsed;
            ProgressBar.Value = 0;
            ProgressTitle.Text = _operation == "merge" ? "Merging PDFs locally…" : "Converting locally…";
            _cancellation = new CancellationTokenSource();
            _runner = new OperationRunner();
            _runner.Progress += (value, message) => Dispatcher.Invoke(() => { ProgressBar.Value = value; ProgressMessage.Text = message; });
            try
            {
                var result = await _runner.RunAsync(_operation, _files.Select(file => file.Path).ToArray(), destination, _cancellation.Token);
                _completedPath = result.OutputPath;
                ProgressPanel.Visibility = Visibility.Collapsed;
                DonePanel.Visibility = Visibility.Visible;
                DoneFileName.Text = Path.GetFileName(_completedPath);
            }
            catch (OperationCanceledException)
            {
                ShowSelectionState();
                MessageBox.Show(this, "The operation was cancelled. No output file was created or replaced.", "Cancelled", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception error) { ShowSelectionState(); ShowError(error.Message); }
            finally
            {
                _runner?.Dispose();
                _runner = null;
                _cancellation?.Dispose();
                _cancellation = null;
                _isBusy = false;
                SetToolButtonsEnabled(true);
            }
        }

        private string AskDestination()
        {
            string extension = _operation == "word-to-pdf" || _operation == "merge" ? ".pdf" : ".docx";
            string baseName = _operation == "merge" ? "Merged" : Path.GetFileNameWithoutExtension(_files[0].Path);
            var dialog = new SaveFileDialog
            {
                AddExtension = true,
                DefaultExt = extension,
                FileName = baseName + extension,
                Filter = extension == ".pdf" ? "PDF file (*.pdf)|*.pdf" : "Word document (*.docx)|*.docx",
                OverwritePrompt = true,
                InitialDirectory = Path.GetDirectoryName(_files[0].Path)
            };
            return dialog.ShowDialog(this) == true ? dialog.FileName : null;
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            CancelButton.IsEnabled = false;
            ProgressMessage.Text = "Cancelling and cleaning up…";
            _cancellation?.Cancel();
            _runner?.Cancel();
        }

        private void Reset_Click(object sender, RoutedEventArgs e) { _files.Clear(); _completedPath = null; ShowSelectionState(); }
        private void OpenFile_Click(object sender, RoutedEventArgs e) { if (File.Exists(_completedPath)) Process.Start(new ProcessStartInfo(_completedPath) { UseShellExecute = true }); }
        private void OpenFolder_Click(object sender, RoutedEventArgs e) { if (File.Exists(_completedPath)) Process.Start("explorer.exe", "/select,\"" + _completedPath + "\""); }

        private void ShowError(string message) => MessageBox.Show(this, message, "LocalPDF", MessageBoxButton.OK, MessageBoxImage.Error);
        private void SetToolButtonsEnabled(bool enabled) { WordToPdfButton.IsEnabled = PdfToWordButton.IsEnabled = MergeButton.IsEnabled = enabled; }

        private void Window_Closing(object sender, CancelEventArgs e)
        {
            if (!_isBusy) return;
            _cancellation?.Cancel();
            _runner?.Cancel();
        }
    }
}
