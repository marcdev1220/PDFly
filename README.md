# PDFly

[![Support on Ko-fi](https://img.shields.io/badge/Ko--fi-Support%20me-FF5E5B?logo=ko-fi&logoColor=white)](https://ko-fi.com/marcdev)

A small Windows desktop app for **batch‑converting between Word and PDF** — drag files or folders in, get exports back. Built with WinUI 3 / Windows App SDK for a modern Fluent look (Mica backdrop, the new `TitleBar`, Segoe Fluent icons).

## Download

Grab the latest **`PDFly-<version>.zip`** from the [**Releases**](../../releases) page, unzip it anywhere, and run **`PDFly.exe`**. It's self‑contained — no .NET install required.

> First launch may show a SmartScreen "Windows protected your PC" prompt (the app is unsigned) — click **More info → Run anyway**.

## What it does

Drop files or folders onto the window (or use **Add files… / Add folder…**). Direction is decided per file by the extension:

| Direction | Engine |
|---|---|
| **Word → PDF** (`.doc / .docx / .docm / .rtf / .odt` → `.pdf`) | Word's `SaveAs2(…, wdFormatPDF)` via COM. |
| **PDF → DOCX** (`.pdf` → `.docx`) | Word's built‑in PDF reflow + `SaveAs2(…, wdFormatXMLDocument)`. Clean on text PDFs; rough on heavily‑laid‑out ones (a scanned PDF comes back as images — Word doesn't OCR). |

- **Converts straight away** as items are added — sequentially, reusing one hidden Word instance that shuts itself down after ~20 s idle.
- **"If the output file already exists"** — *Add a date suffix* (default, never destroys), *Overwrite it*, or *Skip the file*. The preference is remembered.
- **Include files in subfolders** toggle for folder drops.
- **Open the folder when finished** toggle — pops Explorer at the output when a batch ends.
- Output is written **next to the source document**.
- Per‑row right‑click: **Open output**, **Show in folder**, **Convert again**, **Remove from list**. Double‑click opens the produced file.
- Live status per row — *Waiting → Converting → Saved \<file\>* / *Skipped — already exists* / *Failed — \<reason\>* with a coloured glyph.

Preferences live in `%AppData%\PDFly\settings.json`. Crash details, if anything throws, are appended to `%AppData%\PDFly\crash.log`.
## The app
<video src="https://github.com/user-attachments/assets/6287f139-311e-4af5-8640-b974d693fd7c" width="300" controls preload></video>

<video src="https://github.com/user-attachments/assets/6fb7d0e8-378b-4010-919d-c81fa28e11bd" width="300" controls preload></video>

## Requirements

- **Windows 10 1809+** (Windows 11 recommended for the Mica backdrop).
- **Microsoft Word installed** — PDFly drives Word via COM automation. Word **2013 or newer** is needed for the PDF → DOCX direction; Word → PDF works on older Word too.
- Nothing else to install — the bundle ships the Windows App SDK runtime + .NET 10 with the app.

## Support

If it's useful, you can [buy me a coffee on Ko‑fi](https://ko-fi.com/marcdev). Thanks!

## License

[MIT](LICENSE) (c) marcdev1220
