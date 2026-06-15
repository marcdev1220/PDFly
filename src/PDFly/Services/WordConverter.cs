using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using PDFly.Models;

namespace PDFly.Services;

/// <summary>Thrown when Microsoft Word can't be found or launched via COM.</summary>
public sealed class WordUnavailableException(string message) : Exception(message);

public enum ConversionOutcome { Converted, Skipped }

public readonly record struct ConversionResult(ConversionOutcome Outcome, string OutputPath);

/// <summary>
/// Drives a single, reused Word.Application instance via late-bound COM. A <c>.pdf</c>
/// source is reflowed back to <c>.docx</c>; anything else (.doc/.docx/.docm/.rtf/.odt)
/// is exported to PDF — the managed equivalent of the original batch script's VBScript.
/// Not thread-safe: create and use from one (STA) thread only.
/// </summary>
public sealed class WordConverter : IDisposable
{
    private const int WdFormatPdf = 17;          // WdSaveFormat.wdFormatPDF
    private const int WdFormatDocx = 12;         // WdSaveFormat.wdFormatXMLDocument (.docx)
    private const int WdDoNotSaveChanges = 0;    // WdSaveOptions.wdDoNotSaveChanges
    private const int WdAlertsNone = 0;          // WdAlertLevel.wdAlertsNone
    private const int MsoAutomationSecurityForceDisable = 3;

    /// <summary>File extensions PDFly will pick up. <c>.pdf</c> → .docx; the rest → .pdf.</summary>
    public static readonly string[] SupportedExtensions =
        { ".doc", ".docx", ".docm", ".rtf", ".odt", ".pdf" };

    private dynamic? _word;

    public static bool IsSupported(string path)
        => SupportedExtensions.Contains(Path.GetExtension(path), StringComparer.OrdinalIgnoreCase);

    public static bool IsPdf(string path)
        => string.Equals(Path.GetExtension(path), ".pdf", StringComparison.OrdinalIgnoreCase);

    /// <summary>
    /// Converts one document next to the source — PDF → .docx, otherwise → .pdf. Throws
    /// <see cref="WordUnavailableException"/> when Word isn't installed, or an
    /// <see cref="InvalidOperationException"/> with a friendly message when a particular
    /// file can't be opened or saved.
    /// </summary>
    public ConversionResult Convert(string sourcePath, ExistingPdfPolicy policy)
    {
        if (!File.Exists(sourcePath))
            throw new InvalidOperationException("the file no longer exists");

        bool toDocx = IsPdf(sourcePath);
        var (targetPath, skip) = ResolveTargetPath(sourcePath, policy, toDocx ? ".docx" : ".pdf");
        if (skip)
            return new ConversionResult(ConversionOutcome.Skipped, targetPath);

        EnsureWord();

        dynamic word = _word!;
        dynamic? documents = null;
        dynamic? doc = null;
        try
        {
            documents = word.Documents;
            // FileName, ConfirmConversions = false (suppresses the PDF-reflow prompt too),
            // ReadOnly = true, AddToRecentFiles = false
            doc = documents.Open(sourcePath, false, true, false);

            if (File.Exists(targetPath))
            {
                // Overwrite path (or a race) — clear it first so SaveAs2 never trips
                // over a stale/locked file in some Word builds.
                try { File.Delete(targetPath); }
                catch (Exception ex)
                {
                    throw new InvalidOperationException(
                        $"couldn't replace the existing file ({Tidy(ex.Message)})", ex);
                }
            }

            doc.SaveAs2(targetPath, toDocx ? WdFormatDocx : WdFormatPdf);
            return new ConversionResult(ConversionOutcome.Converted, targetPath);
        }
        catch (InvalidOperationException) { throw; }
        catch (Exception ex)
        {
            if (WordSeemsDead()) DisposeWord();   // so the next file relaunches Word
            throw new InvalidOperationException(Tidy(ex.Message), ex);
        }
        finally
        {
            if (doc is not null)
            {
                try { doc.Close(WdDoNotSaveChanges); } catch { }
                Release(doc);
            }
            if (documents is not null) Release(documents);
        }
    }

    private void EnsureWord()
    {
        if (_word is not null) return;

        var type = Type.GetTypeFromProgID("Word.Application", throwOnError: false);
        if (type is null)
            throw new WordUnavailableException(
                "Microsoft Word isn't installed. PDFly converts documents by automating Word " +
                "(the same way the batch script does), so Word has to be installed.");

        dynamic word;
        try
        {
            word = Activator.CreateInstance(type)
                   ?? throw new WordUnavailableException("Microsoft Word couldn't be started.");
        }
        catch (WordUnavailableException) { throw; }
        catch (Exception ex)
        {
            throw new WordUnavailableException(
                "Couldn't start Microsoft Word — make sure it's installed and not stuck on a " +
                $"dialog or activation prompt. ({Tidy(ex.Message)})");
        }

        try
        {
            word.Visible = false;
            word.DisplayAlerts = WdAlertsNone;
            try { word.AutomationSecurity = MsoAutomationSecurityForceDisable; } catch { }
        }
        catch { /* non-fatal: keep going with defaults */ }

        _word = word;
    }

    private bool WordSeemsDead()
    {
        if (_word is null) return true;
        try { _ = _word.Build; return false; }   // cheap string property; throws if the RPC link is gone
        catch { return true; }
    }

    private void DisposeWord()
    {
        if (_word is null) return;
        var word = _word;
        _word = null;
        try { word.Quit(WdDoNotSaveChanges); } catch { }
        Release(word);
        // Word's RCWs are notoriously sticky; nudge the finalizer so WINWORD.EXE exits.
        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();
        GC.WaitForPendingFinalizers();
    }

    private static (string targetPath, bool skip) ResolveTargetPath(
        string sourcePath, ExistingPdfPolicy policy, string outputExtension)
    {
        var dir = Path.GetDirectoryName(sourcePath) ?? string.Empty;
        var baseName = Path.GetFileNameWithoutExtension(sourcePath);
        var plain = Path.Combine(dir, baseName + outputExtension);

        if (!File.Exists(plain))
            return (plain, false);

        switch (policy)
        {
            case ExistingPdfPolicy.Skip:
                return (plain, true);

            case ExistingPdfPolicy.Overwrite:
                return (plain, false);

            case ExistingPdfPolicy.AddDateSuffix:
            default:
                var stamp = DateTime.Now.ToString("yyyyMMdd", CultureInfo.InvariantCulture);
                var candidate = Path.Combine(dir, $"{baseName}_{stamp}{outputExtension}");
                int n = 1;
                while (File.Exists(candidate))
                    candidate = Path.Combine(dir, $"{baseName}_{stamp}_{n++}{outputExtension}");
                return (candidate, false);
        }
    }

    private static string Tidy(string? message)
    {
        var m = message?.Trim();
        if (string.IsNullOrEmpty(m)) return "Word couldn't convert this file";
        return m.Replace("\r", " ").Replace("\n", " ");
    }

    private static void Release(object? comObject)
    {
        if (comObject is null) return;
        try { if (Marshal.IsComObject(comObject)) Marshal.ReleaseComObject(comObject); }
        catch { }
    }

    public void Dispose() => DisposeWord();
}
