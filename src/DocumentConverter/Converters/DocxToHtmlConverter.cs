namespace DocumentConverter.Converters
{
	using System;
	using System.Collections.Generic;
	using System.IO;
	using System.Text;
	using DocumentConverter.Abstractions;
	using DocumentConverter.Models;
	using NPOI.XWPF.UserModel;

	/// <summary>
	/// Converts modern Word processing (.docx) documents to HTML format.
	/// </summary>
	public class DocxToHtmlConverter : IDocumentConverter
	{
		private readonly bool _includeHeaderFooter;
		private readonly bool _includeHtmlWrapper;

		public DocxToHtmlConverter(bool includeHeaderFooter = false, bool includeHtmlWrapper = false)
		{
			_includeHeaderFooter = includeHeaderFooter;
			_includeHtmlWrapper = includeHtmlWrapper;
		}

		/// <summary>
		/// Converts the given DOCX stream to an HTML string.
		/// </summary>
		/// <param name="stream">The DOCX document stream.</param>
		/// <returns>A successful result with the HTML string, or a failed result with an error message.</returns>
		public Result<string> Convert(Stream stream)
		{
			try
			{
				XWPFDocument doc = new XWPFDocument(stream);
				StringBuilder htmlBuilder = new StringBuilder();

				if (_includeHtmlWrapper)
				{
					htmlBuilder.AppendLine("<!DOCTYPE html>");
					htmlBuilder.AppendLine("<html>");
					htmlBuilder.AppendLine("<head>");
					htmlBuilder.AppendLine("<meta charset=\"utf-8\" />");
					htmlBuilder.AppendLine("<title>Converted Word Document</title>");
					htmlBuilder.AppendLine("<style>");
					htmlBuilder.AppendLine("  body { font-family: -apple-system, BlinkMacSystemFont, \"Segoe UI\", Roboto, Helvetica, Arial, sans-serif; color: #334155; line-height: 1.6; padding: 20px; max-width: 800px; margin: 0 auto; }");
					htmlBuilder.AppendLine("  p { margin-bottom: 16px; text-align: justify; }");
					htmlBuilder.AppendLine("  td p, th p { margin: 0 0 8px 0; }");
					htmlBuilder.AppendLine("  td p:last-child, th p:last-child { margin: 0; }");
					htmlBuilder.AppendLine("  h1, h2, h3, h4, h5, h6 { color: #0f172a; margin-top: 24px; margin-bottom: 12px; font-weight: 600; }");
					htmlBuilder.AppendLine("  table { border-collapse: collapse; width: 100%; margin: 20px 0; }");
					htmlBuilder.AppendLine("  th, td { border: 1px solid #cbd5e1; padding: 10px; text-align: left; }");
					htmlBuilder.AppendLine("  th { background-color: #f1f5f9; font-weight: 600; }");
					htmlBuilder.AppendLine("  img { max-width: 100%; height: auto; display: block; margin: 24px auto; border-radius: 4px; }");
					htmlBuilder.AppendLine("</style>");
					htmlBuilder.AppendLine("</head>");
					htmlBuilder.AppendLine("<body>");
				}

				if (_includeHeaderFooter && doc.HeaderList != null)
				{
					foreach (var header in doc.HeaderList)
					{
						if (header.BodyElements != null)
						{
							foreach (IBodyElement element in header.BodyElements)
							{
								if (element is XWPFParagraph paragraph)
								{
									RenderParagraph(paragraph, htmlBuilder);
								}
								else if (element is XWPFTable table)
								{
									RenderTable(table, htmlBuilder);
								}
							}
						}
					}
				}

				foreach (IBodyElement element in doc.BodyElements)
				{
					if (element is XWPFParagraph paragraph)
					{
						RenderParagraph(paragraph, htmlBuilder);
					}
					else if (element is XWPFTable table)
					{
						RenderTable(table, htmlBuilder);
					}
				}

				if (_includeHeaderFooter && doc.FooterList != null)
				{
					foreach (var footer in doc.FooterList)
					{
						if (footer.BodyElements != null)
						{
							foreach (IBodyElement element in footer.BodyElements)
							{
								if (element is XWPFParagraph paragraph)
								{
									RenderParagraph(paragraph, htmlBuilder);
								}
								else if (element is XWPFTable table)
								{
									RenderTable(table, htmlBuilder);
								}
							}
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
				return Result<string>.Failure($"DOCX conversion failed: {ex.Message}");
			}
		}

		private class ProcessedRun
		{
			public string Text { get; set; }
			public byte[] PictureBytes { get; set; }
			public string PictureMime { get; set; }
			public bool IsBold { get; set; }
			public bool IsItalic { get; set; }
			public bool IsUnderline { get; set; }
			public bool IsStrikeThrough { get; set; }
			public double FontSize { get; set; }
			public string FontFamily { get; set; }
			public string Color { get; set; }
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

		private void RenderParagraph(XWPFParagraph paragraph, StringBuilder sb)
		{
			string style = paragraph.Style;
			string tag = "p";

			if (!string.IsNullOrEmpty(style) && style.StartsWith("Heading", StringComparison.OrdinalIgnoreCase))
			{
				char num = style[style.Length - 1];
				if (char.IsDigit(num))
				{
					tag = $"h{num}";
				}
				else
				{
					tag = "h1";
				}
			}

			// Paragraph alignment mapping
			string alignStyle = "";
			switch (paragraph.Alignment)
			{
				case ParagraphAlignment.CENTER:
					alignStyle = "text-align: center;";
					break;
				case ParagraphAlignment.RIGHT:
					alignStyle = "text-align: right;";
					break;
				case ParagraphAlignment.DISTRIBUTE:
					alignStyle = "text-align: justify;";
					break;
			}

			string styleAttr = !string.IsNullOrEmpty(alignStyle) ? $" style=\"{alignStyle}\"" : "";

			List<ProcessedRun> processedRuns = new List<ProcessedRun>();
			foreach (XWPFRun run in paragraph.Runs)
			{
				var pictures = run.GetEmbeddedPictures();
				if (pictures != null && pictures.Count > 0)
				{
					foreach (var pic in pictures)
					{
						var picData = pic.GetPictureData();
						if (picData != null)
						{
							byte[] processedBytes = DocumentConverter.Helpers.ImageHelper.ProcessImage(picData.Data, picData.SuggestFileExtension(), out string mime);
							processedRuns.Add(new ProcessedRun
							{
								PictureBytes = processedBytes,
								PictureMime = mime
							});
						}
					}
				}
				else
				{
					string text = run.Text;
					if (!string.IsNullOrEmpty(text))
					{
						StringBuilder clean = new StringBuilder(text.Length);
						for (int i = 0; i < text.Length; i++)
						{
							char c = text[i];
							if (c == '\x0b')
							{
								clean.Append("\n");
								continue;
							}
							if (char.IsControl(c) && c != '\t' && c != '\n' && c != '\r')
							{
								continue;
							}
							clean.Append(c);
						}

						string cleaned = clean.ToString();
						if (!string.IsNullOrEmpty(cleaned))
						{
							processedRuns.Add(new ProcessedRun
							{
								Text = cleaned,
								IsBold = run.IsBold,
								IsItalic = run.IsItalic,
								IsUnderline = run.Underline != UnderlinePatterns.None,
								IsStrikeThrough = run.IsStrikeThrough,
								Color = run.GetColor(),
								FontSize = run.FontSize,
								FontFamily = run.FontFamily
							});
						}
					}
				}
			}

			// Merge adjacent text runs with same styles
			List<ProcessedRun> mergedRuns = new List<ProcessedRun>();
			foreach (var run in processedRuns)
			{
				if (mergedRuns.Count > 0 && run.Text != null && mergedRuns[mergedRuns.Count - 1].Text != null)
				{
					var prev = mergedRuns[mergedRuns.Count - 1];
					if (prev.IsBold == run.IsBold &&
						prev.IsItalic == run.IsItalic &&
						prev.IsUnderline == run.IsUnderline &&
						prev.IsStrikeThrough == run.IsStrikeThrough &&
						prev.FontSize == run.FontSize &&
						prev.FontFamily == run.FontFamily &&
						prev.Color == run.Color)
					{
						prev.Text += run.Text;
						continue;
					}
				}
				mergedRuns.Add(run);
			}

			if (mergedRuns.Count == 0 && tag == "p")
			{
				return;
			}

			sb.Append($"<{tag}{styleAttr}>");

			foreach (var run in mergedRuns)
			{
				if (run.PictureBytes != null)
				{
					string base64 = System.Convert.ToBase64String(run.PictureBytes);
					sb.Append($"<img src=\"data:{run.PictureMime};base64,{base64}\" />");
				}
				else
				{
					List<string> inlineStyles = new List<string>();
					if (!string.IsNullOrEmpty(run.Color)) inlineStyles.Add($"color: #{run.Color};");
					if (run.FontSize > 0) inlineStyles.Add($"font-size: {run.FontSize}pt;");
					if (!string.IsNullOrEmpty(run.FontFamily)) inlineStyles.Add($"font-family: '{run.FontFamily}', sans-serif;");

					string spanOpen = inlineStyles.Count > 0 ? $"<span style=\"{string.Join(" ", inlineStyles)}\">" : "";
					string spanClose = inlineStyles.Count > 0 ? "</span>" : "";

					string boldOpen = run.IsBold ? "<strong>" : "";
					string boldClose = run.IsBold ? "</strong>" : "";

					string italicOpen = run.IsItalic ? "<em>" : "";
					string italicClose = run.IsItalic ? "</em>" : "";

					string underlineOpen = run.IsUnderline ? "<u>" : "";
					string underlineClose = run.IsUnderline ? "</u>" : "";

					string strikeOpen = run.IsStrikeThrough ? "<del>" : "";
					string strikeClose = run.IsStrikeThrough ? "</del>" : "";

					sb.Append(spanOpen).Append(boldOpen).Append(italicOpen).Append(underlineOpen).Append(strikeOpen);

					string encodedText = SafeHtmlEncode(run.Text);
					encodedText = encodedText.Replace("\r\n", "<br/>").Replace("\n", "<br/>").Replace("\r", "<br/>");
					sb.Append(encodedText);

					sb.Append(strikeClose).Append(underlineClose).Append(italicClose).Append(boldClose).Append(spanClose);
				}
			}

			sb.AppendLine($"</{tag}>");
		}

		private void RenderTable(XWPFTable table, StringBuilder sb)
		{
			sb.AppendLine("<table>");
			foreach (XWPFTableRow row in table.Rows)
			{
				sb.AppendLine("  <tr>");
				foreach (XWPFTableCell cell in row.GetTableCells())
				{
					string color = cell.GetColor();
					string bgStyle = !string.IsNullOrEmpty(color) ? $"background-color: #{color};" : "";
					sb.Append($"    <td style=\"{bgStyle}\">");

					foreach (IBodyElement elem in cell.BodyElements)
					{
						if (elem is XWPFParagraph p)
						{
							RenderParagraph(p, sb);
						}
						else if (elem is XWPFTable tbl)
						{
							RenderTable(tbl, sb);
						}
					}

					sb.AppendLine("</td>");
				}
				sb.AppendLine("  </tr>");
			}
			sb.AppendLine("</table>");
		}
	}
}
