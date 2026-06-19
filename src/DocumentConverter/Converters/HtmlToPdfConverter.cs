namespace DocumentConverter.Converters
{
	using System;
	using System.Collections.Generic;
	using System.IO;
	using System.Linq;
	using System.Text;
	using System.Text.RegularExpressions;
	using DocumentConverter.Abstractions;
	using DocumentConverter.Models;
	using HtmlAgilityPack;
	using PdfSharpCore.Drawing;
	using PdfSharpCore.Drawing.Layout;
	using PdfSharpCore.Pdf;

	/// <summary>
	/// Converts HTML content back to a PDF document using PdfSharpCore.
	/// </summary>
	public class HtmlToPdfConverter : IHtmlToDocumentConverter
	{
		private const double MarginLeft = 50;
		private const double MarginRight = 50;
		private const double MarginTop = 50;
		private const double MarginBottom = 50;

		public Result<byte[]> Convert(string html)
		{
			if (string.IsNullOrEmpty(html))
			{
				return Result<byte[]>.Failure("HTML content cannot be null or empty.");
			}

			try
			{
				var doc = new PdfDocument();
				var htmlDoc = new HtmlDocument();
				htmlDoc.LoadHtml(html);

				// Initialize first page
				PdfPage page = doc.AddPage();
				XGraphics gfx = XGraphics.FromPdfPage(page);
				double currentY = MarginTop;

				var bodyNode = htmlDoc.DocumentNode.SelectSingleNode("//body") ?? htmlDoc.DocumentNode;

				// Process nodes flow layout
				ProcessNodes(bodyNode.ChildNodes, doc, ref page, ref gfx, ref currentY);

				using (var ms = new MemoryStream())
				{
					doc.Save(ms);
					return Result<byte[]>.Success(ms.ToArray());
				}
			}
			catch (Exception ex)
			{
				return Result<byte[]>.Failure($"HTML to PDF conversion failed: {ex.Message}");
			}
		}

		private void ProcessNodes(HtmlNodeCollection nodes, PdfDocument doc, ref PdfPage page, ref XGraphics gfx, ref double currentY)
		{
			if (nodes == null) return;

			double printableWidth = page.Width - MarginLeft - MarginRight;

			foreach (var node in nodes)
			{
				if (node.NodeType == HtmlNodeType.Text)
				{
					string text = node.InnerText?.Trim();
					if (string.IsNullOrEmpty(text)) continue;

					// Draw text paragraph
					var font = new XFont("Arial", 10, XFontStyle.Regular);
					DrawParagraph(text, font, doc, ref page, ref gfx, ref currentY, printableWidth);
					continue;
				}

				string tagName = node.Name.ToLowerInvariant();

				// Explicit Page Divider handling for perfect round-trip
				if (tagName == "div" && node.GetAttributeValue("class", "") == "page-divider")
				{
					// Start new page
					page = doc.AddPage();
					gfx = XGraphics.FromPdfPage(page);
					currentY = MarginTop;
					continue;
				}

				switch (tagName)
				{
					case "h1":
					case "h2":
					case "h3":
					case "h4":
					case "h5":
					case "h6":
						{
							double size = 18;
							if (tagName == "h1") size = 20;
							else if (tagName == "h2") size = 16;
							else if (tagName == "h3") size = 14;
							else if (tagName == "h4") size = 12;
							else if (tagName == "h5" || tagName == "h6") size = 11;

							var font = new XFont("Arial", size, XFontStyle.Bold);
							currentY += 8; // spacing before heading
							DrawParagraph(node.InnerText?.Trim() ?? "", font, doc, ref page, ref gfx, ref currentY, printableWidth);
							currentY += 6; // spacing after heading
							break;
						}
					case "p":
					case "div":
						{
							currentY += 4; // paragraph spacing

							// Check if this node contains any img elements
							bool hasImages = node.SelectNodes(".//img") != null;

							if (hasImages)
							{
								// Recursively process child nodes to handle both text and images
								ProcessNodes(node.ChildNodes, doc, ref page, ref gfx, ref currentY);
							}
							else
							{
								// No images, render as plain text paragraph
								var styleAttr = node.GetAttributeValue("style", "");
								bool isBold = styleAttr.Contains("font-weight: bold") || styleAttr.Contains("font-weight:bold");
								var fontStyle = isBold ? XFontStyle.Bold : XFontStyle.Regular;
								var font = new XFont("Arial", 10, fontStyle);

								DrawParagraph(node.InnerText?.Trim() ?? "", font, doc, ref page, ref gfx, ref currentY, printableWidth);
							}

							currentY += 4;
							break;
						}
					case "ul":
					case "ol":
						{
							int index = 1;
							var liNodes = node.SelectNodes(".//li");
							if (liNodes != null)
							{
								var font = new XFont("Arial", 10, XFontStyle.Regular);
								foreach (var li in liNodes)
								{
									string bullet = tagName == "ol" ? $"{index++}. " : "• ";
									string text = bullet + (li.InnerText?.Trim() ?? "");
									DrawParagraph(text, font, doc, ref page, ref gfx, ref currentY, printableWidth, indent: 15);
								}
							}
							break;
						}
					case "table":
						{
							ProcessTableNode(node, doc, ref page, ref gfx, ref currentY, printableWidth);
							break;
						}
					case "img":
						{
							ProcessImageNode(node, doc, ref page, ref gfx, ref currentY, printableWidth);
							break;
						}
					case "br":
						{
							currentY += 12; // line height
							break;
						}
					default:
						// Skip containers or recursively process child nodes
						if (node.HasChildNodes)
						{
							ProcessNodes(node.ChildNodes, doc, ref page, ref gfx, ref currentY);
						}
						break;
				}
			}
		}

		private void DrawParagraph(string text, XFont font, PdfDocument doc, ref PdfPage page, ref XGraphics gfx, ref double currentY, double printableWidth, double indent = 0)
		{
			if (string.IsNullOrEmpty(text)) return;

			double drawWidth = printableWidth - indent;
			var lines = WrapText(text, font, gfx, drawWidth);

			double lineHeight = font.Size * 1.25;

			foreach (var line in lines)
			{
				EnsureVerticalSpace(lineHeight, doc, ref page, ref gfx, ref currentY);
				gfx.DrawString(line, font, XBrushes.Black, MarginLeft + indent, currentY + font.Size);
				currentY += lineHeight;
			}
		}

		private void ProcessTableNode(HtmlNode tableNode, PdfDocument doc, ref PdfPage page, ref XGraphics gfx, ref double currentY, double printableWidth)
		{
			var rows = tableNode.SelectNodes(".//tr");
			if (rows == null || rows.Count == 0) return;

			// Determine max columns, taking colspan into account
			int maxCols = 0;
			foreach (var row in rows)
			{
				var cells = row.SelectNodes("th|td");
				if (cells != null)
				{
					int colCount = 0;
					foreach (var cell in cells)
					{
						colCount += cell.GetAttributeValue("colspan", 1);
					}
					maxCols = Math.Max(maxCols, colCount);
				}
			}

			if (maxCols == 0) return;

			double colWidth = printableWidth / maxCols;
			double cellPadding = 4;

			currentY += 10; // spacing before table

			var regularFont = new XFont("Arial", 9, XFontStyle.Regular);
			var headerFont = new XFont("Arial", 9, XFontStyle.Bold);

			foreach (var row in rows)
			{
				var cells = row.SelectNodes("th|td");
				if (cells == null) continue;

				// We need to measure row height first
				double maxRowHeight = 0;
				var cellLinesList = new List<List<string>>();

				for (int i = 0; i < cells.Count; i++)
				{
					var cell = cells[i];
					bool isHeader = cell.Name.ToLowerInvariant() == "th";
					var font = isHeader ? headerFont : regularFont;
					int colspan = cell.GetAttributeValue("colspan", 1);
					double cellWidth = colWidth * colspan;

					string text = cell.InnerText?.Trim() ?? "";
					var lines = WrapText(text, font, gfx, cellWidth - (cellPadding * 2));
					cellLinesList.Add(lines);

					double cellHeight = (lines.Count * font.Size * 1.25) + (cellPadding * 2);

					// Also check for images inside the cell and add their height!
					var imgNodes = cell.SelectNodes(".//img");
					if (imgNodes != null)
					{
						foreach (var imgNode in imgNodes)
						{
							double imgHeight = GetImageHeightForCell(imgNode, cellWidth - (cellPadding * 2));
							cellHeight += imgHeight + 4; // spacing
						}
					}

					maxRowHeight = Math.Max(maxRowHeight, cellHeight);
				}

				// If we don't have cells, skip row
				if (cellLinesList.Count == 0) continue;

				// Ensure space for the row
				EnsureVerticalSpace(maxRowHeight, doc, ref page, ref gfx, ref currentY);

				// Draw cells
				double currentX = MarginLeft;
				for (int i = 0; i < cells.Count; i++)
				{
					var cell = cells[i];
					bool isHeader = cell.Name.ToLowerInvariant() == "th";
					var font = isHeader ? headerFont : regularFont;
					var lines = cellLinesList[i];
					int colspan = cell.GetAttributeValue("colspan", 1);
					double cellWidth = colWidth * colspan;

					// Draw cell borders
					gfx.DrawRectangle(XPens.DarkGray, currentX, currentY, cellWidth, maxRowHeight);

					// Draw cell text
					double textY = currentY + cellPadding;
					double lineHeight = font.Size * 1.25;
					foreach (var line in lines)
					{
						gfx.DrawString(line, font, XBrushes.Black, currentX + cellPadding, textY + font.Size);
						textY += lineHeight;
					}

					// Draw cell images below text
					var imgNodes = cell.SelectNodes(".//img");
					if (imgNodes != null)
					{
						double imgY = textY;
						foreach (var imgNode in imgNodes)
						{
							DrawImageInCell(imgNode, gfx, currentX + cellPadding, ref imgY, cellWidth - (cellPadding * 2));
						}
					}

					currentX += cellWidth;
				}

				currentY += maxRowHeight;
			}

			currentY += 10; // spacing after table
		}

		private void ProcessImageNode(HtmlNode node, PdfDocument doc, ref PdfPage page, ref XGraphics gfx, ref double currentY, double printableWidth)
		{
			if (TryGetImageDimensions(node, out byte[] imageBytes, out double width, out double height))
			{
				try
				{
					if (width > printableWidth)
					{
						double ratio = printableWidth / width;
						width = printableWidth;
						height *= ratio;
					}

					currentY += 8; // spacing before image
					EnsureVerticalSpace(height, doc, ref page, ref gfx, ref currentY);

					XImage xImage = XImage.FromStream(() => new MemoryStream(imageBytes));
					gfx.DrawImage(xImage, MarginLeft, currentY, width, height);

					currentY += height + 8; // spacing after image
				}
				catch (Exception ex)
				{
					var font = new XFont("Arial", 9, XFontStyle.Italic);
					DrawParagraph($"[Image layout error: {ex.Message}]", font, doc, ref page, ref gfx, ref currentY, printableWidth);
				}
			}
		}

		private bool TryGetImageDimensions(HtmlNode node, out byte[] imageBytes, out double width, out double height)
		{
			imageBytes = null;
			width = 0;
			height = 0;

			string src = node.GetAttributeValue("src", "");
			if (string.IsNullOrEmpty(src)) return false;

			try
			{
				imageBytes = DocumentConverter.Helpers.ImageHelper.GetImageBytes(src);

				if (imageBytes != null && imageBytes.Length > 0)
				{
					width = node.GetAttributeValue("width", 0);
					height = node.GetAttributeValue("height", 0);

					string styleAttr = node.GetAttributeValue("style", "");
					if (!string.IsNullOrEmpty(styleAttr))
					{
						var wMatch = Regex.Match(styleAttr, @"width\s*:\s*(?<val>\d+)px");
						if (wMatch.Success) double.TryParse(wMatch.Groups["val"].Value, out width);

						var hMatch = Regex.Match(styleAttr, @"height\s*:\s*(?<val>\d+)px");
						if (hMatch.Success) double.TryParse(hMatch.Groups["val"].Value, out height);
					}

					if (width <= 0 || height <= 0)
					{
						if (DocumentConverter.Helpers.ImageHelper.TryGetImageDimensions(imageBytes, out int w, out int h))
						{
							width = w;
							height = h;
						}
						else
						{
							if (width <= 0) width = 300;
							if (height <= 0) height = 200;
						}
					}
					return true;
				}
			}
			catch { }

			return false;
		}

		private double GetImageHeightForCell(HtmlNode node, double cellWidth)
		{
			if (TryGetImageDimensions(node, out _, out double width, out double height))
			{
				if (width > cellWidth)
				{
					double ratio = cellWidth / width;
					height *= ratio;
				}
				return height;
			}
			return 0;
		}

		private void DrawImageInCell(HtmlNode node, XGraphics gfx, double x, ref double y, double cellWidth)
		{
			if (TryGetImageDimensions(node, out byte[] imageBytes, out double width, out double height))
			{
				if (width > cellWidth)
				{
					double ratio = cellWidth / width;
					width = cellWidth;
					height *= ratio;
				}

				try
				{
					XImage xImage = XImage.FromStream(() => new MemoryStream(imageBytes));
					gfx.DrawImage(xImage, x, y, width, height);
					y += height + 4; // increment Y position in cell
				}
				catch { }
			}
		}

		private List<string> WrapText(string text, XFont font, XGraphics gfx, double maxWidth)
		{
			var lines = new List<string>();
			if (string.IsNullOrEmpty(text)) return lines;

			string[] words = text.Split(new[] { ' ', '\r', '\n', '\t' }, StringSplitOptions.RemoveEmptyEntries);
			if (words.Length == 0) return lines;

			var currentLine = new StringBuilder();

			foreach (var word in words)
			{
				int prevLen = currentLine.Length;
				if (prevLen > 0) currentLine.Append(" ");
				currentLine.Append(word);

				string testLine = currentLine.ToString();
				double width = gfx.MeasureString(testLine, font).Width;

				if (width > maxWidth)
				{
					if (prevLen > 0)
					{
						currentLine.Length = prevLen; // backtrack
						lines.Add(currentLine.ToString());
						currentLine.Clear();
						currentLine.Append(word);
					}
					else
					{
						// The word itself is longer than the printable width, force split it
						lines.Add(word);
						currentLine.Clear();
					}
				}
			}

			if (currentLine.Length > 0)
			{
				lines.Add(currentLine.ToString());
			}

			return lines;
		}

		private void EnsureVerticalSpace(double heightNeeded, PdfDocument doc, ref PdfPage page, ref XGraphics gfx, ref double currentY)
		{
			double maxAvailableY = page.Height - MarginBottom;
			if (currentY + heightNeeded > maxAvailableY)
			{
				// Create new page and reset Y
				page = doc.AddPage();
				gfx = XGraphics.FromPdfPage(page);
				currentY = MarginTop;
			}
		}
	}
}
