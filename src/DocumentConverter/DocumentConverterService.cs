namespace DocumentConverter
{
	using System;
	using System.IO;
	using DocumentConverter.Abstractions;
	using DocumentConverter.Converters;
	using DocumentConverter.Models;

	/// <summary>
	/// Entry point service for converting office documents (.doc, .docx, .xls, .xlsx, .pdf) to HTML.
	/// </summary>
	public class DocumentConverterService
	{
		private static readonly System.Collections.Generic.Dictionary<string, Func<bool, bool, IDocumentConverter>> DocumentConverters = 
			new System.Collections.Generic.Dictionary<string, Func<bool, bool, IDocumentConverter>>(StringComparer.OrdinalIgnoreCase)
			{
				{ ".doc", (includeHeaderFooter, includeHtmlWrapper) => new DocToHtmlConverter(includeHeaderFooter, includeHtmlWrapper) },
				{ ".docx", (includeHeaderFooter, includeHtmlWrapper) => new DocxToHtmlConverter(includeHeaderFooter, includeHtmlWrapper) },
				{ ".xls", (includeHeaderFooter, includeHtmlWrapper) => new ExcelToHtmlConverter(includeHtmlWrapper) },
				{ ".xlsx", (includeHeaderFooter, includeHtmlWrapper) => new ExcelToHtmlConverter(includeHtmlWrapper) },
				{ ".pdf", (includeHeaderFooter, includeHtmlWrapper) => new PdfToHtmlConverter(includeHtmlWrapper) }
			};

		private static readonly System.Collections.Generic.Dictionary<string, Func<IHtmlToDocumentConverter>> HtmlConverters = 
			new System.Collections.Generic.Dictionary<string, Func<IHtmlToDocumentConverter>>(StringComparer.OrdinalIgnoreCase)
			{
				{ ".doc", () => new HtmlToDocxConverter() },
				{ ".docx", () => new HtmlToDocxConverter() },
				{ ".xls", () => new HtmlToExcelConverter() },
				{ ".xlsx", () => new HtmlToExcelConverter() },
				{ ".pdf", () => new HtmlToPdfConverter() }
			};

		static DocumentConverterService()
		{
			try
			{
				System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
			}
			catch { }
		}
		/// <summary>
		/// Converts an office document file to HTML string.
		/// </summary>
		/// <param name="filePath">The absolute or relative path to the document file.</param>
		/// <param name="includeHeaderFooter">Determines whether to include header and footer sections for Word documents.</param>
		/// <param name="includeHtmlWrapper">Determines whether to include HTML wrapper tags (DOCTYPE, html, head, style, body) in the output.</param>
		/// <returns>A <see cref="Result{String}"/> containing the HTML content or failure message.</returns>
		public Result<string> ConvertToHtml(string filePath, bool includeHeaderFooter = false, bool includeHtmlWrapper = false)
		{
			if (string.IsNullOrEmpty(filePath))
			{
				return Result<string>.Failure("File path cannot be null or empty.");
			}

			try
			{
				if (!File.Exists(filePath))
				{
					return Result<string>.Failure($"File not found: {filePath}");
				}

				string extension = Path.GetExtension(filePath);
				using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read))
				{
					return ConvertToHtml(fs, extension, includeHeaderFooter, includeHtmlWrapper);
				}
			}
			catch (Exception ex)
			{
				return Result<string>.Failure($"File conversion failed: {ex.Message}");
			}
		}

		/// <summary>
		/// Converts an office document stream to HTML string.
		/// </summary>
		/// <param name="stream">The document stream.</param>
		/// <param name="fileExtension">The file extension (e.g. ".doc", ".docx", ".xls", ".xlsx", ".pdf") indicating the file type.</param>
		/// <param name="includeHeaderFooter">Determines whether to include header and footer sections for Word documents.</param>
		/// <param name="includeHtmlWrapper">Determines whether to include HTML wrapper tags (DOCTYPE, html, head, style, body) in the output.</param>
		/// <returns>A <see cref="Result{String}"/> containing the HTML content or failure message.</returns>
		public Result<string> ConvertToHtml(Stream stream, string fileExtension, bool includeHeaderFooter = false, bool includeHtmlWrapper = false)
		{
			if (stream == null)
			{
				return Result<string>.Failure("Stream cannot be null.");
			}
			if (string.IsNullOrEmpty(fileExtension))
			{
				return Result<string>.Failure("File extension must be provided to determine the format.");
			}

			try
			{
				string ext = fileExtension.Trim().ToLowerInvariant();
				if (!ext.StartsWith("."))
				{
					ext = "." + ext;
				}

				if (!DocumentConverters.TryGetValue(ext, out var factory))
				{
					return Result<string>.Failure($"Unsupported file extension: {fileExtension}");
				}

				IDocumentConverter converter = factory(includeHeaderFooter, includeHtmlWrapper);
				return converter.Convert(stream);
			}
			catch (Exception ex)
			{
				return Result<string>.Failure($"Stream conversion failed: {ex.Message}");
			}
		}

		/// <summary>
		/// Converts an HTML string back to an office document (.docx, .xlsx, .pdf) as a byte array.
		/// </summary>
		/// <param name="htmlContent">The HTML content string.</param>
		/// <param name="targetExtension">The target file extension (e.g. ".docx", ".xlsx", ".pdf").</param>
		/// <returns>A <see cref="Result{Byte[]}"/> containing the generated file bytes.</returns>
		public Result<byte[]> ConvertFromHtml(string htmlContent, string targetExtension)
		{
			if (string.IsNullOrEmpty(htmlContent))
			{
				return Result<byte[]>.Failure("HTML content cannot be null or empty.");
			}
			if (string.IsNullOrEmpty(targetExtension))
			{
				return Result<byte[]>.Failure("Target extension must be provided.");
			}

			try
			{
				string ext = targetExtension.Trim().ToLowerInvariant();
				if (!ext.StartsWith("."))
				{
					ext = "." + ext;
				}

				if (!HtmlConverters.TryGetValue(ext, out var factory))
				{
					return Result<byte[]>.Failure($"Unsupported target extension for HTML conversion: {targetExtension}");
				}

				IHtmlToDocumentConverter converter = factory();
				return converter.Convert(htmlContent);
			}
			catch (Exception ex)
			{
				return Result<byte[]>.Failure($"HTML conversion failed: {ex.Message}");
			}
		}

		/// <summary>
		/// Converts an HTML string back to an office document and writes the bytes to a file path.
		/// </summary>
		/// <param name="htmlContent">The HTML content string.</param>
		/// <param name="targetExtension">The target file extension (e.g. ".docx", ".xlsx", ".pdf").</param>
		/// <param name="outputFilePath">The output file path to write the converted document bytes to.</param>
		/// <returns>A <see cref="Result{Boolean}"/> containing success status.</returns>
		public Result<bool> ConvertFromHtml(string htmlContent, string targetExtension, string outputFilePath)
		{
			if (string.IsNullOrEmpty(outputFilePath))
			{
				return Result<bool>.Failure("Output file path cannot be null or empty.");
			}

			var bytesResult = ConvertFromHtml(htmlContent, targetExtension);
			if (!bytesResult.IsSuccess)
			{
				return Result<bool>.Failure(bytesResult.ErrorMessage);
			}

			try
			{
				string dir = Path.GetDirectoryName(outputFilePath);
				if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
				{
					Directory.CreateDirectory(dir);
				}

				File.WriteAllBytes(outputFilePath, bytesResult.Value);
				return Result<bool>.Success(true);
			}
			catch (Exception ex)
			{
				return Result<bool>.Failure($"Failed to write output file: {ex.Message}");
			}
		}
	}
}
