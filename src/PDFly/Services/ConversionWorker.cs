using System.Collections.Concurrent;
using System.IO;
using PDFly.Models;

namespace PDFly.Services;

public sealed class ConversionProgressEventArgs(
    ConversionItem item, ConversionStatus status, string? outputPath, string? detail) : EventArgs
{
    public ConversionItem Item { get; } = item;
    public ConversionStatus Status { get; } = status;
    public string? OutputPath { get; } = outputPath;
    public string? Detail { get; } = detail;
}

/// <summary>
/// Owns a background STA thread that drains a queue of <see cref="ConversionItem"/>s,
/// converting each via a single reused <see cref="WordConverter"/>. <see cref="Progress"/>
/// and <see cref="QueueDrained"/> are raised on the worker thread — subscribers must
/// marshal to the UI thread themselves.
/// </summary>
public sealed class ConversionWorker : IDisposable
{
    private static readonly TimeSpan IdleQuitDelay = TimeSpan.FromSeconds(20);

    private readonly BlockingCollection<ConversionItem> _queue = new();
    private readonly Thread _thread;
    private WordConverter? _converter;
    private volatile bool _disposed;

    /// <summary>Read fresh for every file, so changing the dropdown affects items still queued.</summary>
    public ExistingPdfPolicy Policy { get; set; } = ExistingPdfPolicy.AddDateSuffix;

    public event EventHandler<ConversionProgressEventArgs>? Progress;
    public event EventHandler? QueueDrained;

    public ConversionWorker()
    {
        _thread = new Thread(Run)
        {
            IsBackground = true,
            Name = "PDFly.WordWorker",
        };
        _thread.SetApartmentState(ApartmentState.STA);
        _thread.Start();
    }

    public void Enqueue(ConversionItem item)
    {
        if (_disposed) return;
        try { _queue.Add(item); }
        catch (InvalidOperationException) { /* CompleteAdding called — shutting down */ }
    }

    private void Run()
    {
        try
        {
            while (!_queue.IsCompleted)
            {
                if (!_queue.TryTake(out var item, IdleQuitDelay))
                {
                    // Idle for a while — release Word so WINWORD.EXE doesn't linger.
                    QuitWord();
                    try { item = _queue.Take(); }                  // block until real work arrives
                    catch (InvalidOperationException) { break; }   // CompleteAdding called
                }
                if (item is null) continue;

                ConvertOne(item);

                if (_queue.Count == 0)
                {
                    try { QueueDrained?.Invoke(this, EventArgs.Empty); } catch { }
                }
            }
        }
        catch (Exception)
        {
            // The worker loop must never bring down the app.
        }
        finally
        {
            QuitWord();
        }
    }

    private void ConvertOne(ConversionItem item)
    {
        Raise(item, ConversionStatus.Converting, null, null);
        try
        {
            _converter ??= new WordConverter();
            var result = _converter.Convert(item.SourcePath, Policy);
            if (result.Outcome == ConversionOutcome.Skipped)
                Raise(item, ConversionStatus.Skipped, result.OutputPath,
                    $"{Path.GetFileName(result.OutputPath)} already exists");
            else
                Raise(item, ConversionStatus.Done, result.OutputPath, null);
        }
        catch (WordUnavailableException ex)
        {
            // Word is the engine — if it's missing, nothing else in the queue will convert.
            Raise(item, ConversionStatus.Failed, null, ex.Message);
            DrainAndFail(ex.Message);
        }
        catch (Exception ex)
        {
            Raise(item, ConversionStatus.Failed, null, ex.Message);
        }
    }

    private void DrainAndFail(string reason)
    {
        while (_queue.TryTake(out var pending))
            Raise(pending, ConversionStatus.Failed, null, reason);
    }

    private void QuitWord()
    {
        var c = _converter;
        _converter = null;
        c?.Dispose();
    }

    private void Raise(ConversionItem item, ConversionStatus status, string? outputPath, string? detail)
    {
        try { Progress?.Invoke(this, new ConversionProgressEventArgs(item, status, outputPath, detail)); }
        catch { /* a misbehaving subscriber must not stall the queue */ }
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
        try { _queue.CompleteAdding(); } catch { }
        // The worker is a background thread; give it a short window to quit Word cleanly,
        // then let process teardown take care of the rest.
        try { _thread.Join(TimeSpan.FromSeconds(5)); } catch { }
        try { _queue.Dispose(); } catch { }
    }
}
