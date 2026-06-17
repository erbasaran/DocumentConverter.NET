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

				IDocumentConverter converter;
				switch (ext)
				{
					case ".doc":
						converter = new DocToHtmlConverter(includeHeaderFooter, includeHtmlWrapper);
						break;
					case ".docx":
						converter = new DocxToHtmlConverter(includeHeaderFooter, includeHtmlWrapper);
						break;
					case ".xls":
					case ".xlsx":
						converter = new ExcelToHtmlConverter(includeHtmlWrapper);
						break;
					case ".pdf":
						converter = new PdfToHtmlConverter(includeHtmlWrapper);
						break;
					default:
						return Result<string>.Failure($"Unsupported file extension: {fileExtension}");
				}

				return converter.Convert(stream);
			}
			catch (Exception ex)
			{
				return Result<string>.Failure($"Stream conversion failed: {ex.Message}");
			}
		}
	}
}
