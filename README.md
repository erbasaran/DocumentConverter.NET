# DocumentConverter

DocumentConverter is a powerful document converter library that supports HTML to PDF, XLSX, and DOCX conversion, as well as PDF, XLS, XLSX, DOC, and DOCX to HTML conversion. It provides a simple and efficient API for transforming documents between popular formats while preserving content and structure.

[![NuGet Version](https://img.shields.io/nuget/v/DocumentConverter.svg?style=flat-square)](https://www.nuget.org/packages/DocumentConverter)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg?style=flat-square)](LICENSE)

---

## Key Features

- **🔄 Bidirectional Conversion**: 
  - **Document ➔ HTML**: Convert legacy/modern Word, Excel, and PDF files into structured HTML.
  - **HTML ➔ Document**: Convert HTML content back into fully formatted `.docx`, `.xlsx`, or `.pdf` files.
- **🐧 100% Cross-Platform & Linux Ready**: Completely decoupled from native OS dependencies like GDI+ or `System.Drawing.Common` for metadata lookup. Run it anywhere without Unix-compatibility workarounds.
- **🖼️ Advanced Image Processing**:
  - Resolves **Base64 Data URIs**, **local file paths**, and **remote HTTP/HTTPS image URLs** (using a thread-safe, cancellation-supported downloader).
  - Preserves cell-relative image alignments in PDF and Excel (anchors images exactly inside their source table cells and auto-scales row heights to prevent overlapping content).
- **⚡ Performance & Memory Optimized**:
  - Implements **Registry Pattern** for extensibility.
  - **Cell Style Caching** for NPOI Excel sheets to prevent Excel style limit exhaustion.
  - Efficient **backtracking StringBuilder text-wrapping** for PDF generation.

---

## Installation

Install the package via NuGet:

```bash
dotnet add package DocumentConverter
```

---

## Usage Examples

### 1. Convert Office Documents to HTML

Initialize the `DocumentConverterService` and perform conversions via file paths or streams:

```csharp
using DocumentConverter;
using DocumentConverter.Models;

var service = new DocumentConverterService();

// Option A: Convert from a file path
Result<string> pathResult = service.ConvertToHtml("report.docx");
if (pathResult.IsSuccess)
{
    string html = pathResult.Value;
    File.WriteAllText("report.html", html);
}

// Option B: Convert from a stream (ideal for web uploads/downloads)
using (var stream = File.OpenRead("data.xlsx"))
{
    Result<string> streamResult = service.ConvertToHtml(stream, ".xlsx");
    if (streamResult.IsSuccess)
    {
        string htmlTable = streamResult.Value;
    }
}
```

### 2. Convert HTML back to Office Documents (Bytes or File)

You can convert HTML back to Word, Excel, or PDF documents. The library will write the file directly or return the raw bytes:

```csharp
using DocumentConverter;
using DocumentConverter.Models;

var service = new DocumentConverterService();
string htmlContent = "<h1>Document Title</h1><p>This is a paragraph.</p>";

// Option A: Convert HTML and save directly to a file
Result<bool> fileResult = service.ConvertFromHtml(htmlContent, ".pdf", "output.pdf");
if (fileResult.IsSuccess)
{
    Console.WriteLine("PDF file generated successfully!");
}

// Option B: Convert HTML and retrieve the raw bytes
Result<byte[]> bytesResult = service.ConvertFromHtml(htmlContent, ".docx");
if (bytesResult.IsSuccess)
{
    byte[] docxBytes = bytesResult.Value;
    File.WriteAllBytes("output.docx", docxBytes);
}
```

---

## Supported Formats Matrix

| Format | Document Type | Document ➔ HTML | HTML ➔ Document | Engine / Parser |
| :---: | :---: | :---: | :---: | :---: |
| **`.doc`** | Legacy Word (97-2003) |  | ❌ | NPOI / HWPF |
| **`.docx`** | Modern Word (OpenXML) |  |  | NPOI / XWPF |
| **`.xls`** | Legacy Excel (97-2003) |  | ❌ | NPOI / HSSF |
| **`.xlsx`** | Modern Excel (OpenXML) |  |  | NPOI / XSSF |
| **`.pdf`** | PDF Document |  |  | UglyToad.PdfPig / PdfSharpCore |

---

## Contributing & Extensibility

Adding a new document format converter is simple. The service implements a **Registry Pattern**. You can register custom converters in `DocumentConverterService` by matching the `IDocumentConverter` (for Document ➔ HTML) and `IHtmlToDocumentConverter` (for HTML ➔ Document) interfaces.

---

## License

This project is licensed under the **MIT License**.
