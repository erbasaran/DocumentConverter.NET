namespace DocumentConverter.Converters
{
	using System;
	using System.Collections.Generic;
	using System.IO;
	using System.IO.Compression;
	using System.Text;
	using DocumentConverter.Abstractions;
	using DocumentConverter.Models;
	using NPOI.HWPF;
	using NPOI.HWPF.Model;
	using NPOI.HWPF.UserModel;

	/// <summary>
	/// Converts legacy Word processing (.doc) documents to HTML format.
	/// </summary>
	public class DocToHtmlConverter : IDocumentConverter
	{
		private class CachedPicture
		{
			public byte[] Bytes { get; set; }
			public string MimeType { get; set; }
		}

		private int _pictureIndex = 0;
		private List<CachedPicture> _cachedPictures = new List<CachedPicture>();
		private readonly bool _includeHeaderFooter;
		private readonly bool _includeHtmlWrapper;

		public DocToHtmlConverter(bool includeHeaderFooter = false, bool includeHtmlWrapper = false)
		{
			_includeHeaderFooter = includeHeaderFooter;
			_includeHtmlWrapper = includeHtmlWrapper;
		}

		/// <summary>
		/// Converts the given DOC stream to an HTML string.
		/// </summary>
		/// <param name="stream">The DOC document stream.</param>
		/// <returns>A successful result with the HTML string, or a failed result with an error message.</returns>
		public Result<string> Convert(Stream stream)
		{
			try
			{
				_pictureIndex = 0;
				HWPFDocument doc = new HWPFDocument(stream);
				PicturesTable picturesTable = doc.GetPicturesTable();
				_cachedPictures = new List<CachedPicture>();
				try
				{
					var allPics = picturesTable.GetAllPictures();
					if (allPics != null)
					{
						foreach (var pic in allPics)
						{
							try
							{
								byte[] bytes = GetPictureBytes(pic, out string mime);
								if (bytes != null)
								{
									_cachedPictures.Add(new CachedPicture { Bytes = bytes, MimeType = mime });
								}
							}
							catch
							{
								// Ignore single picture load failure
							}
						}
					}
				}
				catch
				{
					// Ignore pictures table failure
				}
				Range range = doc.GetOverallRange();
				Range headerRange = doc.GetHeaderStoryRange();
				int headerStart = headerRange != null ? headerRange.StartOffset : -1;
				int headerEnd = headerRange != null ? headerRange.EndOffset : -1;
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

				for (int i = 0; i < range.NumParagraphs; i++)
				{
					Paragraph p = range.GetParagraph(i);
					if (!_includeHeaderFooter && headerStart != -1 && p.StartOffset >= headerStart && p.EndOffset <= headerEnd)
					{
						continue;
					}
					if (p.IsInTable())
					{
						try
						{
							Table table = range.GetTable(p);
							RenderTable(table, picturesTable, htmlBuilder);

							// Skip the paragraphs within this table
							if (table.NumParagraphs > 0)
							{
								i += table.NumParagraphs - 1;
							}
						}
						catch
						{
							// If table parsing fails, render it as a normal paragraph
							if (!IsParagraphEmpty(p, picturesTable))
							{
								RenderParagraph(p, picturesTable, htmlBuilder);
							}
						}
					}
					else
					{
						if (!IsParagraphEmpty(p, picturesTable))
						{
							RenderParagraph(p, picturesTable, htmlBuilder);
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
				return Result<string>.Failure($"DOC conversion failed: {ex.Message}");
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
			public int FontSize { get; set; }
			public string FontFamily { get; set; }
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

		private void RenderParagraph(Paragraph paragraph, PicturesTable picturesTable, StringBuilder sb)
		{
			string tag = "p";

			// Paragraph alignment mapping
			string alignStyle = "";
			switch (paragraph.GetJustification())
			{
				case 1:
					alignStyle = "text-align: center;";
					break;
				case 2:
					alignStyle = "text-align: right;";
					break;
				case 3:
					alignStyle = "text-align: justify;";
					break;
			}

			string styleAttr = !string.IsNullOrEmpty(alignStyle) ? $" style=\"{alignStyle}\"" : "";

			List<ProcessedRun> processedRuns = new List<ProcessedRun>();
			for (int j = 0; j < paragraph.NumCharacterRuns; j++)
			{
				CharacterRun run = paragraph.GetCharacterRun(j);

				// Check for picture
				if (picturesTable != null && (picturesTable.HasPicture(run) || picturesTable.HasEscherPicture(run)))
				{
					bool isPicRun = picturesTable.HasPicture(run);
					if (!isPicRun && picturesTable.HasEscherPicture(run))
					{
						try
						{
							isPicRun = picturesTable.ExtractPicture(run, true) != null;
						}
						catch
						{
							// Ignore
						}
					}

					if (isPicRun)
					{
						if (_pictureIndex >= 0 && _pictureIndex < _cachedPictures.Count)
						{
							var cachedPic = _cachedPictures[_pictureIndex];
							processedRuns.Add(new ProcessedRun
							{
								PictureBytes = cachedPic.Bytes,
								PictureMime = cachedPic.MimeType
							});
						}
						_pictureIndex++;
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
								IsBold = run.IsBold(),
								IsItalic = run.IsItalic(),
								IsUnderline = run.GetUnderlineCode() > 0,
								IsStrikeThrough = run.IsStrikeThrough(),
								FontSize = run.GetFontSize() / 2, // HWPF FontSize is in half-points
								FontFamily = run.GetFontName()
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
						prev.FontFamily == run.FontFamily)
					{
						prev.Text += run.Text;
						continue;
					}
				}
				mergedRuns.Add(run);
			}

			if (mergedRuns.Count == 0)
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

		private byte[] GetPictureBytes(Picture pic, out string mimeType)
		{
			byte[] rawBytes = null;
			string extension = pic.SuggestFileExtension();

			try
			{
				// Try standard NPOI extraction first
				rawBytes = pic.GetContent();
			}
			catch
			{
				// Fallback: manually parse and decompress raw zlib stream from raw content
				try
				{
					byte[] raw = pic.GetRawContent();

					// Detect zlib header 78 (with checksum val % 31 == 0)
					int zlibOffset = -1;
					for (int i = 0; i < raw.Length - 1; i++)
					{
						if (raw[i] == 0x78)
						{
							int val = (raw[i] << 8) | raw[i + 1];
							if (val % 31 == 0)
							{
								zlibOffset = i;
								break;
							}
						}
					}

					if (zlibOffset != -1)
					{
						using (var msInput = new MemoryStream(raw, zlibOffset + 2, raw.Length - (zlibOffset + 2)))
						using (var deflate = new DeflateStream(msInput, CompressionMode.Decompress))
						using (var msOutput = new MemoryStream())
						{
							deflate.CopyTo(msOutput);
							rawBytes = msOutput.ToArray();
						}
					}
				}
				catch
				{
					// Ignore and return raw if decompression fails
				}

				if (rawBytes == null)
				{
					rawBytes = pic.GetRawContent();
				}
			}

			// Process the raw bytes through our centralized ImageHelper
			return DocumentConverter.Helpers.ImageHelper.ProcessImage(rawBytes, extension, out mimeType);
		}

		private void RenderTable(Table table, PicturesTable picturesTable, StringBuilder sb)
		{
			if (table.NumRows == 0) return;

			// 1. Collect all unique X-coordinates (left and right edges)
			List<int> xList = new List<int>();
			for (int r = 0; r < table.NumRows; r++)
			{
				TableRow row = table.GetRow(r);
				for (int c = 0; c < row.NumCells(); c++)
				{
					TableCell cell = row.GetCell(c);
					int startX = cell.GetLeftEdge();
					int endX = startX + cell.GetWidth();
					if (!xList.Contains(startX)) xList.Add(startX);
					if (!xList.Contains(endX)) xList.Add(endX);
				}
			}
			xList.Sort();

			int numCols = xList.Count - 1;
			if (numCols <= 0) return;

			// Grid to store cell references and tracking of rendered slots
			TableCell[,] grid = new TableCell[table.NumRows, numCols];
			bool[,] covered = new bool[table.NumRows, numCols];

			// 2. Populate the grid
			for (int r = 0; r < table.NumRows; r++)
			{
				TableRow row = table.GetRow(r);
				for (int c = 0; c < row.NumCells(); c++)
				{
					TableCell cell = row.GetCell(c);
					if (cell.IsMerged() && !cell.IsFirstMerged())
					{
						continue; // Skip horizontally merged cells (they are represented by the first cell)
					}

					int startX = cell.GetLeftEdge();
					int endX = startX + cell.GetWidth();

					if (cell.IsMerged() && cell.IsFirstMerged())
					{
						// Expand endX to include all subsequent merged cells in this row
						for (int nextC = c + 1; nextC < row.NumCells(); nextC++)
						{
							TableCell nextCell = row.GetCell(nextC);
							if (nextCell.IsMerged() && !nextCell.IsFirstMerged())
							{
								endX = nextCell.GetLeftEdge() + nextCell.GetWidth();
							}
							else
							{
								break;
							}
						}
					}

					int startCol = xList.IndexOf(startX);
					int endCol = xList.IndexOf(endX);

					if (startCol != -1 && endCol != -1)
					{
						for (int col = startCol; col < endCol; col++)
						{
							grid[r, col] = cell;
						}
					}
				}
			}

			sb.AppendLine("<table>");

			// 3. Render cells using the grid
			for (int r = 0; r < table.NumRows; r++)
			{
				sb.AppendLine("  <tr>");
				for (int col = 0; col < numCols; col++)
				{
					if (covered[r, col])
					{
						continue; // Already covered by rowspan/colspan
					}

					TableCell cell = grid[r, col];
					if (cell == null)
					{
						sb.Append("    <td></td>");
						continue;
					}

					// Find colspan for this cell
					int startCol = col;
					int endCol = col;
					while (endCol < numCols && grid[r, endCol] == cell)
					{
						endCol++;
					}
					int colSpan = endCol - startCol;

					// Find rowspan for this cell
					int rowSpan = 1;
					if (cell.IsVerticallyMerged() && cell.IsFirstVerticallyMerged())
					{
						for (int nextR = r + 1; nextR < table.NumRows; nextR++)
						{
							TableCell nextCell = grid[nextR, startCol];
							if (nextCell != null && nextCell.IsVerticallyMerged() && !nextCell.IsFirstVerticallyMerged())
							{
								// Verify it spans the same columns
								bool sameSpan = true;
								for (int c = startCol; c < endCol; c++)
								{
									if (grid[nextR, c] != nextCell)
									{
										sameSpan = false;
										break;
									}
								}

								if (sameSpan)
								{
									rowSpan++;
								}
								else
								{
									break;
								}
							}
							else
							{
								break;
							}
						}
					}

					// Mark slots as covered
					for (int dr = 0; dr < rowSpan; dr++)
					{
						for (int dc = 0; dc < colSpan; dc++)
						{
							covered[r + dr, startCol + dc] = true;
						}
					}

					// Render cell
					string spans = "";
					if (colSpan > 1) spans += $" colspan=\"{colSpan}\"";
					if (rowSpan > 1) spans += $" rowspan=\"{rowSpan}\"";

					sb.Append($"    <td{spans}>");

					for (int p = 0; p < cell.NumParagraphs; p++)
					{
						Paragraph para = cell.GetParagraph(p);
						if (!IsParagraphEmpty(para, picturesTable))
						{
							RenderParagraph(para, picturesTable, sb);
						}
					}

					sb.AppendLine("</td>");

					// Move column pointer forward
					col = endCol - 1;
				}
				sb.AppendLine("  </tr>");
			}
			sb.AppendLine("</table>");
		}

		private bool IsParagraphEmpty(Paragraph p, PicturesTable picturesTable)
		{
			for (int j = 0; j < p.NumCharacterRuns; j++)
			{
				CharacterRun run = p.GetCharacterRun(j);
				if (picturesTable != null)
				{
					if (picturesTable.HasPicture(run))
					{
						return false;
					}
					if (picturesTable.HasEscherPicture(run))
					{
						try
						{
							var pic = picturesTable.ExtractPicture(run, true);
							if (pic != null)
							{
								return false;
							}
						}
						catch
						{
							// Ignore and check other conditions
						}
					}
				}
			}

			string text = p.Text;
			if (string.IsNullOrEmpty(text))
			{
				return true;
			}

			for (int i = 0; i < text.Length; i++)
			{
				char c = text[i];
				if (char.IsWhiteSpace(c) || c == '\xa0')
				{
					continue;
				}
				if (char.IsControl(c))
				{
					continue;
				}
				return false; // Found a visible character
			}

			return true;
		}

		private string GetMimeTypeFromExtension(string extension)
		{
			if (string.IsNullOrEmpty(extension)) return "image/png";
			extension = extension.TrimStart('.').ToLowerInvariant();
			switch (extension)
			{
				case "jpg":
				case "jpeg":
					return "image/jpeg";
				case "png":
					return "image/png";
				case "gif":
					return "image/gif";
				case "bmp":
					return "image/bmp";
				case "tiff":
				case "tif":
					return "image/tiff";
				case "svg":
					return "image/svg+xml";
				default:
					return $"image/{extension}";
			}
		}
	}
}
