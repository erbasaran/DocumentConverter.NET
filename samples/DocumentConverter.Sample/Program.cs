using System;
using System.IO;
using System.Text;
using DocumentConverter;
using DocumentConverter.Models;

Console.OutputEncoding = Encoding.UTF8;

Console.ForegroundColor = ConsoleColor.Cyan;
Console.WriteLine("=================================================================");
Console.WriteLine("                DocumentConverter.NET Console Sample            ");
Console.WriteLine("=================================================================");
Console.ResetColor();
Console.WriteLine();

// 1. Initialize the converter service
Console.WriteLine("Initializing DocumentConverterService...");
var service = new DocumentConverterService();
Console.ForegroundColor = ConsoleColor.Green;
Console.WriteLine("✔ Service successfully initialized!");
Console.ResetColor();
Console.WriteLine();

// 2. Define sample HTML content
string htmlContent = @"<!DOCTYPE html>
<html>
<head>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; color: #333; }
        h1 { color: #2C3E50; }
        h2 { color: #16A085; border-bottom: 1px solid #ddd; padding-bottom: 5px; }
        p { line-height: 1.6; }
        table { border-collapse: collapse; width: 100%; margin-top: 15px; }
        th, td { border: 1px solid #bdc3c7; padding: 10px; text-align: left; }
        th { background-color: #ecf0f1; color: #2c3e50; }
        .highlight { background-color: #f1c40f; font-weight: bold; }
    </style>
</head>
<body>
    <h1>DocumentConverter.NET Demo</h1>
    <p>This document was converted programmatically using <strong>DocumentConverter.NET</strong>.</p>
    
    <h2>Library Capabilities</h2>
    <ul>
        <li>Convert <b>HTML</b> input into Word (<b>.docx</b>), Excel (<b>.xlsx</b>), and PDF (<b>.pdf</b>) files.</li>
        <li>Convert <b>Word, Excel, and PDF</b> documents back into clean, semantic <b>HTML</b>.</li>
        <li>Flexible API supporting both file paths and stream operations.</li>
    </ul>

    <h2>Office / Employee Details Table</h2>
    <table>
        <thead>
            <tr>
                <th>ID</th>
                <th>Employee Name</th>
                <th>Department</th>
                <th>Role</th>
            </tr>
        </thead>
        <tbody>
            <tr>
                <td>1</td>
                <td class=""highlight"">Ramazan Erbasaran</td>
                <td>Software Engineering</td>
                <td>Lead Developer</td>
            </tr>
            <tr>
                <td>2</td>
                <td>Alice Green</td>
                <td>UI/UX Design</td>
                <td>Designer</td>
            </tr>
            <tr>
                <td>3</td>
                <td>Bob Miller</td>
                <td>Product Management</td>
                <td>Product Owner</td>
            </tr>
        </tbody>
    </table>

    <br />
    <p>Sample ended. For more information, visit the repository: <a href=""https://github.com/erbasaran/DocumentConverter.NET"">GitHub Repository</a></p>
</body>
</html>";

string outputDir = Path.Combine(AppContext.BaseDirectory, "sample_outputs");
if (!Directory.Exists(outputDir))
{
    Directory.CreateDirectory(outputDir);
}
Console.WriteLine($"Output files will be saved in: {outputDir}");
Console.WriteLine();

// --- STAGE 1: HTML to Documents ---
Console.ForegroundColor = ConsoleColor.Yellow;
Console.WriteLine("--- STAGE 1: Converting HTML to Office Documents ---");
Console.ResetColor();

// A. HTML to DOCX
string docxPath = Path.Combine(outputDir, "sample.docx");
Console.Write("Converting HTML to DOCX... ");
var docxResult = service.ConvertFromHtml(htmlContent, ".docx", docxPath);
PrintResult(docxResult, docxPath);

// B. HTML to XLSX
string xlsxPath = Path.Combine(outputDir, "sample.xlsx");
Console.Write("Converting HTML to XLSX... ");
var xlsxResult = service.ConvertFromHtml(htmlContent, ".xlsx", xlsxPath);
PrintResult(xlsxResult, xlsxPath);

// C. HTML to PDF
string pdfPath = Path.Combine(outputDir, "sample.pdf");
Console.Write("Converting HTML to PDF... ");
var pdfResult = service.ConvertFromHtml(htmlContent, ".pdf", pdfPath);
PrintResult(pdfResult, pdfPath);

Console.WriteLine();

// --- STAGE 2: Documents to HTML (Roundtrip) ---
Console.ForegroundColor = ConsoleColor.Yellow;
Console.WriteLine("--- STAGE 2: Converting Documents back to HTML (Roundtrip) ---");
Console.ResetColor();

// A. DOCX to HTML
if (File.Exists(docxPath))
{
    Console.Write("Converting DOCX back to HTML... ");
    var docxToHtmlResult = service.ConvertToHtml(docxPath, includeHeaderFooter: false, includeHtmlWrapper: true);
    PrintHtmlConversionResult(docxToHtmlResult);
}

// B. XLSX to HTML
if (File.Exists(xlsxPath))
{
    Console.Write("Converting XLSX back to HTML... ");
    var xlsxToHtmlResult = service.ConvertToHtml(xlsxPath, includeHeaderFooter: false, includeHtmlWrapper: true);
    PrintHtmlConversionResult(xlsxToHtmlResult);
}

// C. PDF to HTML
if (File.Exists(pdfPath))
{
    Console.Write("Converting PDF back to HTML... ");
    var pdfToHtmlResult = service.ConvertToHtml(pdfPath, includeHeaderFooter: false, includeHtmlWrapper: true);
    PrintHtmlConversionResult(pdfToHtmlResult);
}

Console.WriteLine();
Console.ForegroundColor = ConsoleColor.Cyan;
Console.WriteLine("=================================================================");
Console.WriteLine("             All conversion samples completed successfully!     ");
Console.WriteLine("=================================================================");
Console.ResetColor();

void PrintResult<T>(Result<T> result, string path)
{
    if (result.IsSuccess)
    {
        Console.ForegroundColor = ConsoleColor.Green;
        Console.WriteLine("✔ SUCCESS");
        Console.ResetColor();
        Console.WriteLine($"  File saved: {Path.GetFileName(path)} ({new FileInfo(path).Length} bytes)");
    }
    else
    {
        Console.ForegroundColor = ConsoleColor.Red;
        Console.WriteLine("✖ FAILED");
        Console.ResetColor();
        Console.WriteLine($"  Error: {result.ErrorMessage}");
    }
}

void PrintHtmlConversionResult(Result<string> result)
{
    if (result.IsSuccess)
    {
        Console.ForegroundColor = ConsoleColor.Green;
        Console.WriteLine("✔ SUCCESS");
        Console.ResetColor();
        Console.WriteLine($"  HTML output generated successfully (Length: {result.Value.Length} characters).");
        
        // Print a short preview
        string preview = result.Value.Length > 200 ? result.Value.Substring(0, 200) + "..." : result.Value;
        Console.ForegroundColor = ConsoleColor.DarkGray;
        Console.WriteLine("  Preview:");
        Console.WriteLine("  -------------------------------------------------------------");
        string[] lines = preview.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
        foreach (var line in lines)
        {
            Console.WriteLine($"  | {line.Trim()}");
        }
        Console.WriteLine("  -------------------------------------------------------------");
        Console.ResetColor();
    }
    else
    {
        Console.ForegroundColor = ConsoleColor.Red;
        Console.WriteLine("✖ FAILED");
        Console.ResetColor();
        Console.WriteLine($"  Error: {result.ErrorMessage}");
    }
}
