namespace DocumentConverter.Converters
{
	using System;
	using System.Collections.Generic;
	using System.IO;
	using System.Linq;
	using System.Text;
	using DocumentConverter.Abstractions;
	using DocumentConverter.Models;
	using UglyToad.PdfPig;
	using UglyToad.PdfPig.Content;

	/// <summary>
	/// Converts PDF documents to HTML format using UglyToad.PdfPig.
	/// </summary>
	public class PdfToHtmlConverter : IDocumentConverter
	{
		private readonly bool _includeHtmlWrapper;

		public PdfToHtmlConverter(bool includeHtmlWrapper = false)
		{
			_includeHtmlWrapper = includeHtmlWrapper;
		}

		private class PageElement
		{
			public double Bottom { get; set; }
			public double Height { get; set; }
			public ElementType Type { get; set; }
			public string Text { get; set; }
			public byte[] ImageData { get; set; }
			public string ImageMime { get; set; }
		}

		private enum ElementType
		{
			Text,
			Image
		}

		public Result<string> Convert(Stream stream)
		{
			try
			{
				StringBuilder htmlBuilder = new StringBuilder();

				if (_includeHtmlWrapper)
				{
					htmlBuilder.AppendLine("<!DOCTYPE html>");
					htmlBuilder.AppendLine("<html>");
					htmlBuilder.AppendLine("<head>");
					htmlBuilder.AppendLine("<meta charset=\"utf-8\" />");
					htmlBuilder.AppendLine("<title>Converted PDF Document</title>");
					htmlBuilder.AppendLine("<style>");
					htmlBuilder.AppendLine("  body { font-family: -apple-system, BlinkMacSystemFont, \"Segoe UI\", Roboto, Helvetica, Arial, sans-serif; color: #334155; padding: 20px; max-width: 800px; margin: 0 auto; line-height: 1.6; }");
					htmlBuilder.AppendLine("  p { margin-bottom: 16px; text-align: justify; text-justify: inter-word; }");
					htmlBuilder.AppendLine("  img { max-width: 100%; height: auto; display: block; margin: 24px auto; border-radius: 4px; }");
					htmlBuilder.AppendLine("  .page-divider { border-top: 1px dashed #cbd5e1; margin: 40px 0; position: relative; text-align: center; }");
					htmlBuilder.AppendLine("  .page-number { background-color: #ffffff; color: #64748b; font-size: 0.8rem; padding: 0 8px; position: absolute; top: -10px; left: 50%; transform: translateX(-50%); }");
					htmlBuilder.AppendLine("</style>");
					htmlBuilder.AppendLine("</head>");
					htmlBuilder.AppendLine("<body>");
				}

				using (PdfDocument document = PdfDocument.Open(stream))
				{
					int pageCount = document.NumberOfPages;
					for (int pNum = 1; pNum <= pageCount; pNum++)
					{
						var page = document.GetPage(pNum);
						if (page == null) continue;

						if (pNum > 1)
						{
							htmlBuilder.AppendLine($"  <div class=\"page-divider\"><span class=\"page-number\">Page {pNum}</span></div>");
						}

						// Collect all page elements (text lines and images)
						List<PageElement> elements = new List<PageElement>();

						// 1. Process text words into lines
						var words = page.GetWords();
						if (words != null && words.Any())
						{
							var lines = new List<List<Word>>();
							foreach (var word in words.OrderByDescending(w => w.BoundingBox.Bottom).ThenBy(w => w.BoundingBox.Left))
							{
								bool placed = false;
								foreach (var line in lines)
								{
									double avgBottom = line.Average(w => w.BoundingBox.Bottom);
									double avgHeight = line.Average(w => w.BoundingBox.Height);
									double tolerance = avgHeight * 0.3;
									if (Math.Abs(word.BoundingBox.Bottom - avgBottom) < Math.Max(tolerance, 3.0))
									{
										line.Add(word);
										placed = true;
										break;
									}
								}
								if (!placed)
								{
									lines.Add(new List<Word> { word });
								}
							}

							foreach (var line in lines)
							{
								var sortedLine = line.OrderBy(w => w.BoundingBox.Left).ToList();
								string lineText = string.Join(" ", sortedLine.Select(w => w.Text));
								if (!string.IsNullOrWhiteSpace(lineText))
								{
									elements.Add(new PageElement
									{
										Bottom = sortedLine.Average(w => w.BoundingBox.Bottom),
										Height = sortedLine.Average(w => w.BoundingBox.Height),
										Type = ElementType.Text,
										Text = lineText
									});
								}
							}
						}

						// 2. Process images
						var images = page.GetImages();
						if (images != null)
						{
							foreach (var img in images)
							{
								byte[] imgBytes = null;
								string mime = "image/png";

								if (img.TryGetPng(out byte[] pngBytes))
								{
									imgBytes = pngBytes;
									mime = "image/png";
								}
								else
								{
									try
									{
										var raw = img.RawBytes.ToArray();
										if (raw.Length >= 3 && raw[0] == 0xFF && raw[1] == 0xD8 && raw[2] == 0xFF)
										{
											imgBytes = raw;
											mime = "image/jpeg";
										}
										else
										{
											imgBytes = raw;
											mime = "image/png";
										}
									}
									catch
									{
										continue;
									}
								}

								if (imgBytes != null && imgBytes.Length > 0)
								{
									elements.Add(new PageElement
									{
										Bottom = img.Bounds.Bottom,
										Height = img.Bounds.Height,
										Type = ElementType.Image,
										ImageData = imgBytes,
										ImageMime = mime
									});
								}
							}
						}

						// Sort all elements visually descending by their Bottom coordinate
						var sortedElements = elements.OrderByDescending(e => e.Bottom).ToList();

						List<string> paragraphBuffer = new List<string>();
						double lastLineBottom = -1;

						foreach (var elem in sortedElements)
						{
							if (elem.Type == ElementType.Text)
							{
								if (lastLineBottom != -1)
								{
									double distance = lastLineBottom - elem.Bottom - elem.Height;
									if (distance > elem.Height * 1.8) // Paragraph boundary
									{
										if (paragraphBuffer.Count > 0)
										{
											htmlBuilder.AppendLine($"  <p>{SafeHtmlEncode(string.Join(" ", paragraphBuffer))}</p>");
											paragraphBuffer.Clear();
										}
									}
								}

								paragraphBuffer.Add(elem.Text);
								lastLineBottom = elem.Bottom;
							}
							else if (elem.Type == ElementType.Image)
							{
								// Flush text buffer before image
								if (paragraphBuffer.Count > 0)
								{
									htmlBuilder.AppendLine($"  <p>{SafeHtmlEncode(string.Join(" ", paragraphBuffer))}</p>");
									paragraphBuffer.Clear();
								}
								lastLineBottom = -1; // Reset line alignment after image

								string base64 = System.Convert.ToBase64String(elem.ImageData);
								htmlBuilder.AppendLine($"  <img src=\"data:{elem.ImageMime};base64,{base64}\" />");
							}
						}

						// Flush remaining text
						if (paragraphBuffer.Count > 0)
						{
							htmlBuilder.AppendLine($"  <p>{SafeHtmlEncode(string.Join(" ", paragraphBuffer))}</p>");
						}
					}
				}

				if (_includeHtmlWrapper)
				{
					htmlBuilder.AppendLine("</body>");
					htmlBuilder.AppendLine("</html>");
				}

				return Result<string>.Success(htmlBuilder.ToString());
			}
			catch (Exception ex)
			{
				return Result<string>.Failure($"PDF conversion failed: {ex.Message}");
			}
		}

		private static string SafeHtmlEncode(string text)
		{
			if (string.IsNullOrEmpty(text)) return text;
			return text.Replace("&", "&amp;")
					   .Replace("<", "&lt;")
					   .Replace(">", "&gt;")
					   .Replace("\"", "&quot;")
					   .Replace("'", "&#39;");
		}
	}
}
