using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.Diagnostics;
using System.IO;
using System.Threading.Tasks;
using Microsoft.UI.Dispatching;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using PDFly.Infrastructure;
using PDFly.Models;
using PDFly.Services;
using Windows.Storage.Pickers;
using WinRT.Interop;

namespace PDFly.ViewModels;

/// <summary>ComboBox-friendly wrapper around <see cref="ExistingPdfPolicy"/>.</summary>
public sealed record PolicyChoice(ExistingPdfPolicy Policy, string Label)
{
    public override string ToString() => Label;
}

public sealed class MainViewModel : ObservableObject, IDisposable
{
    private static readonly string[] AcceptedExtensions = WordConverter.SupportedExtensions;

    private readonly PdflySettings _settings;
    private readonly ConversionWorker _worker;
    private readonly DispatcherQueue _dispatcher;
    private bool _wordWarningShown;
    private string? _lastOutputFolder;

    public ObservableCollection<ConversionItem> Items { get; } = new();

    public IReadOnlyList<PolicyChoice> PolicyChoices { get; } = new[]
    {
        new PolicyChoice(ExistingPdfPolicy.AddDateSuffix, "Add a date suffix"),
        new PolicyChoice(ExistingPdfPolicy.Overwrite,     "Overwrite it"),
        new PolicyChoice(ExistingPdfPolicy.Skip,          "Skip the file"),
    };

    public RelayCommand AddFilesCommand { get; }
    public RelayCommand AddFolderCommand { get; }
    public RelayCommand ClearCommand { get; }
    public RelayCommand OpenOutputCommand { get; }

    public MainViewModel(PdflySettings settings)
    {
        _settings = settings;
        // Acquire the dispatcher for the current (UI) thread. MainViewModel is always
        // constructed from the WinUI UI thread, so this is safe and avoids depending on
        // App.DispatcherQueue being set yet (it's assigned later in App.OnLaunched).
        _dispatcher = DispatcherQueue.GetForCurrentThread();

        _worker = new ConversionWorker { Policy = settings.ExistingPolicy };
        _worker.Progress += OnWorkerProgress;
        _worker.QueueDrained += OnWorkerQueueDrained;

        _selectedPolicy = PolicyChoices.FirstOrDefault(p => p.Policy == settings.ExistingPolicy) ?? PolicyChoices[0];
        _includeSubfolders = settings.IncludeSubfolders;
        _openFolderWhenDone = settings.OpenFolderWhenDone;

        Items.CollectionChanged += OnItemsCollectionChanged;

        AddFilesCommand   = new RelayCommand(() => _ = AddFilesViaDialogAsync());
        AddFolderCommand  = new RelayCommand(() => _ = AddFolderViaDialogAsync());
        ClearCommand      = new RelayCommand(ClearList, () => Items.Any(i => i.CanRemove));
        OpenOutputCommand = new RelayCommand(OpenLastOutputFolder, () => _lastOutputFolder is not null);
    }

    // ---------------------------------------------------------------- options

    private PolicyChoice _selectedPolicy;
    public PolicyChoice SelectedPolicy
    {
        get => _selectedPolicy;
        set
        {
            if (value is null || !SetProperty(ref _selectedPolicy, value)) return;
            _worker.Policy = value.Policy;
            _settings.ExistingPolicy = value.Policy;
            _settings.Save();
        }
    }

    private bool _includeSubfolders;
    public bool IncludeSubfolders
    {
        get => _includeSubfolders;
        set
        {
            if (!SetProperty(ref _includeSubfolders, value)) return;
            _settings.IncludeSubfolders = value;
            _settings.Save();
        }
    }

    private bool _openFolderWhenDone;
    public bool OpenFolderWhenDone
    {
        get => _openFolderWhenDone;
        set
        {
            if (!SetProperty(ref _openFolderWhenDone, value)) return;
            _settings.OpenFolderWhenDone = value;
            _settings.Save();
        }
    }

    // ------------------------------------------------------------- list state

    private bool _isDragOver;
    public bool IsDragOver { get => _isDragOver; set => SetProperty(ref _isDragOver, value); }

    public bool IsEmpty => Items.Count == 0;
    public bool HasActiveItems => Items.Any(i => i.IsActive);

    private int _convertedCount, _skippedCount, _failedCount;
    public int ConvertedCount { get => _convertedCount; private set { if (SetProperty(ref _convertedCount, value)) OnPropertyChanged(nameof(SummaryText)); } }
    public int SkippedCount   { get => _skippedCount;   private set { if (SetProperty(ref _skippedCount, value))   OnPropertyChanged(nameof(SummaryText)); } }
    public int FailedCount    { get => _failedCount;    private set { if (SetProperty(ref _failedCount, value))    OnPropertyChanged(nameof(SummaryText)); } }

    public string SummaryText
    {
        get
        {
            if (IsEmpty) return "Drop Word or PDF documents to convert them.";
            int active = Items.Count(i => i.IsActive);
            if (active > 0) return $"Converting…   {Items.Count - active} of {Items.Count} done";
            var parts = new List<string>(3);
            if (_convertedCount > 0) parts.Add($"{_convertedCount} converted");
            if (_skippedCount > 0)   parts.Add($"{_skippedCount} skipped");
            if (_failedCount > 0)    parts.Add($"{_failedCount} failed");
            return parts.Count == 0 ? "Done." : "Done   ·   " + string.Join("   ·   ", parts);
        }
    }

    // -------------------------------------------------------------- ingestion

    /// <summary>Add files / folders (folders are expanded to Word documents), dedupe
    /// against what's already queued, kick the worker.</summary>
    public void AddPaths(IEnumerable<string> paths)
    {
        var seen = new HashSet<string>(Items.Select(i => i.SourcePath), StringComparer.OrdinalIgnoreCase);
        var toAdd = new List<string>();

        foreach (var raw in paths)
        {
            if (string.IsNullOrWhiteSpace(raw)) continue;
            string full;
            try { full = Path.GetFullPath(raw.Trim().Trim('"')); }
            catch { continue; }

            try
            {
                if (Directory.Exists(full))
                {
                    foreach (var f in EnumerateDocs(full, _includeSubfolders))
                        if (seen.Add(f)) toAdd.Add(f);
                }
                else if (File.Exists(full) && WordConverter.IsSupported(full) && !IsWordTempFile(full))
                {
                    if (seen.Add(full)) toAdd.Add(full);
                }
            }
            catch { /* unreadable path — skip */ }
        }

        if (toAdd.Count == 0) return;

        toAdd.Sort((a, b) => string.Compare(a, b, StringComparison.OrdinalIgnoreCase));
        foreach (var f in toAdd)
        {
            var item = new ConversionItem(f);
            Items.Add(item);
            _worker.Enqueue(item);
        }
        ClearCommand.RaiseCanExecuteChanged();
        OpenOutputCommand.RaiseCanExecuteChanged();
    }

    private static IEnumerable<string> EnumerateDocs(string folder, bool recurse)
    {
        var options = new EnumerationOptions
        {
            RecurseSubdirectories = recurse,
            IgnoreInaccessible = true,
            AttributesToSkip = FileAttributes.Hidden | FileAttributes.System,
        };
        foreach (var f in Directory.EnumerateFiles(folder, "*", options))
            if (WordConverter.IsSupported(f) && !IsWordTempFile(f))
                yield return f;
    }

    private static bool IsWordTempFile(string path)
        => Path.GetFileName(path).StartsWith("~$", StringComparison.Ordinal);

    private async Task AddFilesViaDialogAsync()
    {
        var picker = new FileOpenPicker { SuggestedStartLocation = PickerLocationId.ComputerFolder };
        foreach (var ext in AcceptedExtensions) picker.FileTypeFilter.Add(ext);
        InitializeWithWindow.Initialize(picker, App.WindowHandle);
        var files = await picker.PickMultipleFilesAsync();
        if (files is { Count: > 0 })
            AddPaths(files.Select(f => f.Path));
    }

    private async Task AddFolderViaDialogAsync()
    {
        var picker = new FolderPicker { SuggestedStartLocation = PickerLocationId.ComputerFolder };
        picker.FileTypeFilter.Add("*");
        InitializeWithWindow.Initialize(picker, App.WindowHandle);
        var folder = await picker.PickSingleFolderAsync();
        if (folder is not null)
            AddPaths(new[] { folder.Path });
    }

    // ------------------------------------------------------------- list edits

    private void ClearList()
    {
        for (int i = Items.Count - 1; i >= 0; i--)
            if (Items[i].CanRemove) Items.RemoveAt(i);

        if (Items.Count == 0) _lastOutputFolder = null;
        ClearCommand.RaiseCanExecuteChanged();
        OpenOutputCommand.RaiseCanExecuteChanged();
    }

    public void RemoveItem(ConversionItem item)
    {
        if (!item.CanRemove) return;
        Items.Remove(item);
        ClearCommand.RaiseCanExecuteChanged();
    }

    public void RequeueItem(ConversionItem item)
    {
        if (!item.CanRemove) return;
        if (!File.Exists(item.SourcePath))
        {
            item.Detail = "the file no longer exists";
            item.OutputPath = null;
            item.Status = ConversionStatus.Failed;
            RecomputeCounts();
            return;
        }
        item.OutputPath = null;
        item.Detail = null;
        item.Status = ConversionStatus.Pending;
        _worker.Enqueue(item);
        RecomputeCounts();
        OnPropertyChanged(nameof(SummaryText));
    }

    public void OpenItem(ConversionItem item)
    {
        if (!string.IsNullOrEmpty(item.OutputPath) && File.Exists(item.OutputPath))
            ShellOpen(item.OutputPath);
        else if (File.Exists(item.SourcePath))
            ShellOpen(item.SourcePath);
    }

    public void RevealItem(ConversionItem item)
    {
        var target = !string.IsNullOrEmpty(item.OutputPath) && File.Exists(item.OutputPath)
            ? item.OutputPath
            : item.SourcePath;
        RevealInExplorer(target);
    }

    private void OpenLastOutputFolder()
    {
        if (_lastOutputFolder is { } f && Directory.Exists(f))
            ShellOpen(f);
    }

    // -------------------------------------------------------- worker callbacks

    private void OnWorkerProgress(object? sender, ConversionProgressEventArgs e)
    {
        if (!_dispatcher.HasThreadAccess)
        {
            _dispatcher.TryEnqueue(() => ApplyProgress(e.Item, e.Status, e.OutputPath, e.Detail));
            return;
        }
        ApplyProgress(e.Item, e.Status, e.OutputPath, e.Detail);
    }

    private void ApplyProgress(ConversionItem item, ConversionStatus status, string? outputPath, string? detail)
    {
        if (detail is not null) item.Detail = detail;
        if (outputPath is not null) item.OutputPath = outputPath;
        item.Status = status;

        if (!string.IsNullOrEmpty(outputPath))
        {
            var folder = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(folder))
            {
                if (status == ConversionStatus.Done) _lastOutputFolder = folder;
                else _lastOutputFolder ??= folder;
            }
        }

        if (status == ConversionStatus.Failed && item.Detail is { } d &&
            d.Contains("Word isn't installed", StringComparison.OrdinalIgnoreCase) && !_wordWarningShown)
        {
            _wordWarningShown = true;
            _ = ShowWarningAsync("Microsoft Word required", d);
        }

        RecomputeCounts();
        OnPropertyChanged(nameof(SummaryText));
        OnPropertyChanged(nameof(HasActiveItems));
        ClearCommand.RaiseCanExecuteChanged();
        OpenOutputCommand.RaiseCanExecuteChanged();
    }

    private void OnWorkerQueueDrained(object? sender, EventArgs e)
    {
        if (!_dispatcher.HasThreadAccess)
        {
            _dispatcher.TryEnqueue(OnQueueDrainedUi);
            return;
        }
        OnQueueDrainedUi();
    }

    private void OnQueueDrainedUi()
    {
        if (HasActiveItems) return;
        OnPropertyChanged(nameof(SummaryText));
        if (_openFolderWhenDone && _lastOutputFolder is { } f && Directory.Exists(f) && Items.Any(i => i.IsDone))
            ShellOpen(f);
    }

    private static async Task ShowWarningAsync(string title, string message)
    {
        try
        {
            var content = App.Window?.Content;
            if (content is null) return;
            var dialog = new ContentDialog
            {
                Title = title,
                Content = message,
                CloseButtonText = "Close",
                XamlRoot = content.XamlRoot,
            };
            await dialog.ShowAsync();
        }
        catch { /* dialog construction can fail during shutdown — best effort */ }
    }

    // ------------------------------------------------------------------ helpers

    private void OnItemsCollectionChanged(object? sender, NotifyCollectionChangedEventArgs e)
    {
        RecomputeCounts();
        OnPropertyChanged(nameof(IsEmpty));
        OnPropertyChanged(nameof(HasActiveItems));
        OnPropertyChanged(nameof(SummaryText));
        ClearCommand.RaiseCanExecuteChanged();
        OpenOutputCommand.RaiseCanExecuteChanged();
    }

    private void RecomputeCounts()
    {
        int done = 0, skip = 0, fail = 0;
        foreach (var i in Items)
            switch (i.Status)
            {
                case ConversionStatus.Done: done++; break;
                case ConversionStatus.Skipped: skip++; break;
                case ConversionStatus.Failed: fail++; break;
            }
        ConvertedCount = done;
        SkippedCount = skip;
        FailedCount = fail;
    }

    private static void ShellOpen(string path)
    {
        try { Process.Start(new ProcessStartInfo(path) { UseShellExecute = true }); }
        catch { }
    }

    private static void RevealInExplorer(string path)
    {
        try
        {
            if (File.Exists(path))
                Process.Start(new ProcessStartInfo("explorer.exe", $"/select,\"{path}\"") { UseShellExecute = true });
            else if (Directory.Exists(path))
                Process.Start(new ProcessStartInfo("explorer.exe", $"\"{path}\"") { UseShellExecute = true });
        }
        catch { }
    }

    public void Dispose()
    {
        _worker.Progress -= OnWorkerProgress;
        _worker.QueueDrained -= OnWorkerQueueDrained;
        _worker.Dispose();
    }
}
