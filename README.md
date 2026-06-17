# DocumentConverter

An open-source, lightweight, and cross-platform .NET Standard 2.0 library designed to convert office documents (**Word, Excel, and PDF**) into clean, semantic HTML formats. 

It is ideal for displaying document previews, indexing text content, or integrating document viewers into web applications.

---

## Installation

Add the library to your solution or project references. It depends on standard open-source libraries:
* **NPOI** & **ScratchPad.NPOI.HWPF** (for Word and Excel formats)
* **UglyToad.PdfPig** (for native PDF parsing)
* **System.Drawing.Common** (for Windows image manipulation)

---

## Usage Examples

Here are some standard ways to convert your documents in C#:

### 1. Converting a File Path to HTML
The simplest way is to pass the absolute or relative file path to the converter service:

```csharp
using DocumentConverter;
using DocumentConverter.Models;

// Initialize the converter service
var service = new DocumentConverterService();

// Convert a Word document path
Result<string> result = service.ConvertToHtml("C:\\Documents\\report.docx", includeHeaderFooter: false);

if (result.IsSuccess)
{
    string htmlContent = result.Value;
    // Save or display the HTML output
    System.IO.File.WriteAllText("C:\\Documents\\report.html", htmlContent);
    Console.WriteLine("Document converted successfully!");
}
else
{
    Console.WriteLine($"Conversion failed: {result.ErrorMessage}");
}
```

### 2. Converting a File Stream
If you are receiving files via a web upload or database stream, you can convert them directly by passing the stream along with its file extension:

```csharp
using System.IO;
using DocumentConverter;
using DocumentConverter.Models;

var service = new DocumentConverterService();

using (FileStream stream = File.OpenRead("C:\\Documents\\spreadsheet.xls"))
{
    // Pass the stream, file extension (dot prefix is optional), and options
    Result<string> result = service.ConvertToHtml(stream, ".xls");

    if (result.IsSuccess)
    {
        string htmlTable = result.Value;
        // Proceed with your business logic (e.g., render in a web view)
    }
    else
    {
        Console.WriteLine($"Error: {result.ErrorMessage}");
    }
}
```

---

## Supported Formats

| Extension | Document Type | Engine |
| :--- | :--- | :--- |
| `.doc` | Legacy Word Document (97-2003) | NPOI / HWPF |
| `.docx` | Modern Word Document (OpenXML) | NPOI / XWPF |
| `.xls` | Legacy Excel Spreadsheet (97-2003) | NPOI / HSSF |
| `.xlsx` | Modern Excel Spreadsheet (OpenXML) | NPOI / XSSF |
| `.pdf` | PDF Document | UglyToad.PdfPig |

---

## License

This project is open-source and released under the MIT License.
