using System.IO;
using System.Text.Json;
using System.Text.Json.Serialization;
using PDFly.Models;

namespace PDFly.Services;

/// <summary>Tiny persisted preferences blob in %AppData%\PDFly\settings.json.</summary>
public sealed class PdflySettings
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true,
        Converters = { new JsonStringEnumConverter() },
    };

    private static string FilePath
    {
        get
        {
            var dir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "PDFly");
            Directory.CreateDirectory(dir);
            return Path.Combine(dir, "settings.json");
        }
    }

    public ExistingPdfPolicy ExistingPolicy { get; set; } = ExistingPdfPolicy.AddDateSuffix;
    public bool IncludeSubfolders { get; set; }
    public bool OpenFolderWhenDone { get; set; }

    public static PdflySettings Load()
    {
        try
        {
            var path = FilePath;
            if (!File.Exists(path)) return new PdflySettings();
            var json = File.ReadAllText(path);
            if (string.IsNullOrWhiteSpace(json)) return new PdflySettings();
            return JsonSerializer.Deserialize<PdflySettings>(json, JsonOptions) ?? new PdflySettings();
        }
        catch { return new PdflySettings(); }
    }

    public void Save()
    {
        try
        {
            var json = JsonSerializer.Serialize(this, JsonOptions);
            var path = FilePath;
            var tmp = path + ".tmp";
            File.WriteAllText(tmp, json);
            File.Move(tmp, path, overwrite: true);
        }
        catch { /* best effort */ }
    }
}
