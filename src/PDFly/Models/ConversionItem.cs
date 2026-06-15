using System.IO;
using PDFly.Infrastructure;

namespace PDFly.Models;

public enum ConversionStatus
{
    Pending,
    Converting,
    Done,
    Skipped,
    Failed,
}

/// <summary>What to do when a PDF with the target name already exists — mirrors the
/// "skip / overwrite / date suffix" prompt from the original batch script.</summary>
public enum ExistingPdfPolicy
{
    AddDateSuffix,
    Overwrite,
    Skip,
}

/// <summary>One Word document queued for conversion, plus its live status.</summary>
public sealed class ConversionItem : ObservableObject
{
    public ConversionItem(string sourcePath)
    {
        SourcePath = sourcePath;
        FileName = Path.GetFileName(sourcePath);
        SourceFolder = Path.GetDirectoryName(sourcePath) ?? string.Empty;
    }

    public string SourcePath { get; }
    public string FileName { get; }
    public string SourceFolder { get; }

    private ConversionStatus _status = ConversionStatus.Pending;
    public ConversionStatus Status
    {
        get => _status;
        set
        {
            if (!SetProperty(ref _status, value)) return;
            OnPropertyChanged(nameof(StatusText));
            OnPropertyChanged(nameof(IsDone));
            OnPropertyChanged(nameof(IsFailed));
            OnPropertyChanged(nameof(IsActive));
            OnPropertyChanged(nameof(CanRemove));
        }
    }

    private string? _outputPath;
    /// <summary>The PDF produced (Done) or the pre-existing PDF that was kept (Skipped).</summary>
    public string? OutputPath
    {
        get => _outputPath;
        set
        {
            if (!SetProperty(ref _outputPath, value)) return;
            OnPropertyChanged(nameof(OutputFileName));
            OnPropertyChanged(nameof(StatusText));
        }
    }

    public string? OutputFileName => string.IsNullOrEmpty(_outputPath) ? null : Path.GetFileName(_outputPath);

    private string? _detail;
    /// <summary>Free-text note: the error message (Failed) or the reason (Skipped).</summary>
    public string? Detail
    {
        get => _detail;
        set { if (SetProperty(ref _detail, value)) OnPropertyChanged(nameof(StatusText)); }
    }

    public bool IsDone => _status == ConversionStatus.Done;
    public bool IsFailed => _status == ConversionStatus.Failed;
    public bool IsActive => _status is ConversionStatus.Pending or ConversionStatus.Converting;
    public bool CanRemove => _status != ConversionStatus.Converting;

    public string StatusText => _status switch
    {
        ConversionStatus.Pending => "Waiting…",
        ConversionStatus.Converting => "Converting…",
        ConversionStatus.Done => OutputFileName is { } f ? $"Saved  {f}" : "Done",
        ConversionStatus.Skipped => string.IsNullOrEmpty(_detail) ? "Skipped" : _detail,
        ConversionStatus.Failed => string.IsNullOrEmpty(_detail) ? "Failed" : $"Failed — {_detail}",
        _ => string.Empty,
    };
}
