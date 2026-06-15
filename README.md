# PDFly

[![Support on Ko-fi](https://img.shields.io/badge/Ko--fi-Support%20me-FF5E5B?logo=ko-fi&logoColor=white)](https://ko-fi.com/marcdev)

A small Windows desktop app for **batch‑converting between Word and PDF** — drag in files or folders, and PDFly exports `.doc/.docx/.docm/.rtf/.odt` to PDF, and `.pdf` back to `.docx`. Built with **WinUI 3 / Windows App SDK** for a modern Fluent look (Mica backdrop, the new `TitleBar`, Segoe Fluent icons).

Conversion uses **Microsoft Word** under the hood (same engine the old batch‑file recipe used — Word's COM automation + `SaveAs2`), so Word needs to be installed.

## Download

Grab the latest `PDFly-<version>.zip` from the [**Releases**](../../releases) page, unzip it anywhere, and run **`PDFly.exe`**. It's self‑contained — no .NET install required.

> On first launch Windows SmartScreen may show a "Windows protected your PC" prompt (the app is unsigned) — click **More info → Run anyway**.

## What it does

- **Drop or pick** files / folders (or both). Direction is decided per file by the extension: `.pdf` → `.docx`, everything else → `.pdf`.
- **Converts straight away** as items are added — sequentially, reusing one hidden Word instance which shuts itself down after ~20 s of idle.
- **"If the output file already exists"** — choose *Add a date suffix* (`name_yyyymmdd.ext`, never destroys anything — default), *Overwrite it*, or *Skip the file*. The preference is remembered.
- **Include files in subfolders** toggle — when set, a dropped/picked folder pulls in documents from its whole subtree (inaccessible folders are skipped, hidden files are kept, system files are skipped).
- **Open the folder when finished** toggle — pops Explorer at the output once a batch finishes.
- Output is written next to the source document.
- Per‑row right‑click: **Open output**, **Show in folder**, **Convert again**, **Remove from list**. Double‑click a row opens the output (or the source if it hasn't run yet).
- Live status per row — *Waiting → Converting → Saved <file>* / *Skipped — already exists* / *Failed — <reason>* with a coloured glyph.

| Direction | What Word does |
|---|---|
| **Word → PDF** | Opens the document, exports via `SaveAs2(…, wdFormatPDF)`. |
| **PDF → DOCX** | Opens the PDF using Word's built‑in **PDF reflow**, saves as `.docx` via `SaveAs2(…, wdFormatXMLDocument)`. Clean on text PDFs; rough on heavily‑laid‑out ones (a scanned PDF comes back as images — Word doesn't OCR). |

Preferences live in `%AppData%\PDFly\settings.json`. Crash details, if anything goes wrong, are appended to `%AppData%\PDFly\crash.log`.

## Requirements

- **Windows 10 1809+** (Windows 11 recommended for Mica).
- **Microsoft Word installed** — PDFly drives Word via COM automation. Word **2013 or newer** is needed for the PDF → DOCX direction; Word → PDF works on older Word too.
- To build from source: **.NET 10 SDK** + the Windows 10 SDK (`10.0.26100`).

## Build, run & test

```powershell
dotnet build src/PDFly/PDFly.csproj
dotnet run   --project src/PDFly/PDFly.csproj
```

For a releasable self‑contained build, use the shared `build.bat` in the parent folder (`C:\Programs\.Projects\C#Projects\build.bat`) → pick **PDFly** → **3** (publish). The result is a folder at `src\PDFly\dist\Release-win-x64\` (PDFly.exe + the bundled Windows App SDK runtime); zip it for distribution.

## License

[MIT](LICENSE) (c) marcdev1220

## Support

If it's useful, you can [buy me a coffee on Ko‑fi](https://ko-fi.com/marcdev). Thanks!
