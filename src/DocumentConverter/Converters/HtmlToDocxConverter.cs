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
	using NPOI.XWPF.UserModel;

	/// <summary>
	/// Converts HTML content back to Word document (.docx) format using NPOI.
	/// </summary>
	public class HtmlToDocxConverter : IHtmlToDocumentConverter
	{
		public Result<byte[]> Convert(string html)
		{
			if (string.IsNullOrEmpty(html))
			{
				return Result<byte[]>.Failure("HTML content cannot be null or empty.");
			}

			try
			{
				var doc = new XWPFDocument();
				var htmlDoc = new HtmlDocument();
				htmlDoc.LoadHtml(html);

				var bodyNode = htmlDoc.DocumentNode.SelectSingleNode("//body") ?? htmlDoc.DocumentNode;

				// Process elements starting from body or root
				ProcessNodes(bodyNode.ChildNodes, doc, null);

				using (var ms = new MemoryStream())
				{
					doc.Write(ms);
					return Result<byte[]>.Success(ms.ToArray());
				}
			}
			catch (Exception ex)
			{
				return Result<byte[]>.Failure($"HTML to DOCX conversion failed: {ex.Message}");
			}
		}

		private void ProcessNodes(HtmlNodeCollection nodes, XWPFDocument doc, XWPFTable tableContext)
		{
			if (nodes == null) return;

			foreach (var node in nodes)
			{
				if (node.NodeType == HtmlNodeType.Text)
				{
					string text = node.InnerText?.Trim();
					if (string.IsNullOrEmpty(text)) continue;

					// Write orphaned text to a default paragraph
					var p = doc.CreateParagraph();
					var run = p.CreateRun();
					run.SetText(text);
					continue;
				}

				string tagName = node.Name.ToLowerInvariant();
				switch (tagName)
				{
					case "p":
					case "h1":
					case "h2":
					case "h3":
					case "h4":
					case "h5":
					case "h6":
					case "div":
						{
							var p = doc.CreateParagraph();
							ApplyParagraphStyles(node, p);
							if (tagName.StartsWith("h"))
							{
								p.Style = "Heading" + tagName.Substring(1);
							}
							ProcessInlineNodes(node.ChildNodes, p, doc);
							break;
						}
					case "ul":
					case "ol":
						{
							ProcessList(node, doc, tagName == "ol");
							break;
						}
					case "table":
						{
							ProcessTable(node, doc);
							break;
						}
					case "img":
						{
							var p = doc.CreateParagraph();
							ProcessImageNode(node, p, doc);
							break;
						}
					case "br":
						{
							// Orphaned br, create a spacing/empty paragraph
							doc.CreateParagraph();
							break;
						}
					default:
						// If we encounter standard formatting or unknown tags at block level, wrap them in a paragraph
						if (IsInlineElement(node))
						{
							var p = doc.CreateParagraph();
							ProcessInlineNode(node, p, doc, new TextStyle());
						}
						else
						{
							// Recursively process children
							ProcessNodes(node.ChildNodes, doc, tableContext);
						}
						break;
				}
			}
		}

		private void ProcessList(HtmlNode listNode, XWPFDocument doc, bool isOrdered)
		{
			int index = 1;
			foreach (var li in listNode.SelectNodes(".//li") ?? Enumerable.Empty<HtmlNode>())
			{
				var p = doc.CreateParagraph();
				var run = p.CreateRun();
				if (isOrdered)
				{
					run.SetText($"{index++}. ");
				}
				else
				{
					run.SetText("• ");
				}
				ProcessInlineNodes(li.ChildNodes, p, doc);
			}
		}

		private void ProcessTable(HtmlNode tableNode, XWPFDocument doc)
		{
			var rows = tableNode.SelectNodes(".//tr");
			if (rows == null || rows.Count == 0) return;

			// We need to determine the maximum number of columns across all rows to initialize the table correctly
			int maxCols = 0;
			foreach (var row in rows)
			{
				var cells = row.SelectNodes("th|td");
				if (cells != null)
				{
					int colCount = 0;
					foreach (var cell in cells)
					{
						int colspan = cell.GetAttributeValue("colspan", 1);
						colCount += colspan;
					}
					maxCols = Math.Max(maxCols, colCount);
				}
			}

			if (maxCols == 0) maxCols = 1;

			var xwpfTable = doc.CreateTable(rows.Count, maxCols);

			for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++)
			{
				var rowNode = rows[rowIndex];
				var xwpfRow = xwpfTable.GetRow(rowIndex);
				var cellNodes = rowNode.SelectNodes("th|td");
				if (cellNodes == null) continue;

				int cellColIndex = 0;
				foreach (var cellNode in cellNodes)
				{
					int colspan = cellNode.GetAttributeValue("colspan", 1);

					if (cellColIndex < xwpfRow.GetTableCells().Count)
					{
						var xwpfCell = xwpfRow.GetCell(cellColIndex);
						if (xwpfCell == null)
						{
							xwpfCell = xwpfRow.CreateCell();
						}

						// Apply gridSpan for colspan
						if (colspan > 1)
						{
							var tcPr = xwpfCell.GetCTTc().tcPr ?? xwpfCell.GetCTTc().AddNewTcPr();
							var gridSpan = new NPOI.OpenXmlFormats.Wordprocessing.CT_DecimalNumber();
							gridSpan.val = colspan.ToString();
							tcPr.gridSpan = gridSpan;

							// Remove C - 1 cells from the end
							for (int i = 0; i < colspan - 1; i++)
							{
								if (xwpfRow.GetTableCells().Count > cellColIndex + 1)
								{
									xwpfRow.RemoveCell(xwpfRow.GetTableCells().Count - 1);
								}
							}
						}

						// Apply styling (e.g. bold header th)
						bool isHeader = cellNode.Name.ToLowerInvariant() == "th";

						// Cell text and paragraphs
						var p = xwpfCell.Paragraphs.FirstOrDefault() ?? xwpfCell.AddParagraph();
						var style = new TextStyle { IsBold = isHeader };

						// Parse cell style attribute
						var styleAttr = cellNode.GetAttributeValue("style", "");
						if (!string.IsNullOrEmpty(styleAttr))
						{
							if (styleAttr.Contains("font-weight: bold") || styleAttr.Contains("font-weight:700"))
							{
								style.IsBold = true;
							}
							if (styleAttr.Contains("font-style: italic"))
							{
								style.IsItalic = true;
							}
						}

						// Process contents inside cell
						if (cellNode.HasChildNodes)
						{
							ProcessInlineNodes(cellNode.ChildNodes, p, doc, style);
						}
						else
						{
							var run = p.CreateRun();
							run.SetText(cellNode.InnerText);
							if (style.IsBold) run.IsBold = true;
						}

						cellColIndex++;
					}
				}

				// Remove any remaining extra cells in this row
				while (xwpfRow.GetTableCells().Count > cellColIndex)
				{
					xwpfRow.RemoveCell(xwpfRow.GetTableCells().Count - 1);
				}
			}
		}

		private void ProcessInlineNodes(HtmlNodeCollection nodes, XWPFParagraph p, XWPFDocument doc, TextStyle inheritedStyle = null)
		{
			if (nodes == null) return;
			var style = inheritedStyle ?? new TextStyle();

			foreach (var node in nodes)
			{
				ProcessInlineNode(node, p, doc, style);
			}
		}

		private void ProcessInlineNode(HtmlNode node, XWPFParagraph p, XWPFDocument doc, TextStyle currentStyle)
		{
			if (node.NodeType == HtmlNodeType.Text)
			{
				string text = node.InnerText;
				if (string.IsNullOrEmpty(text)) return;

				var run = p.CreateRun();
				run.SetText(text);
				ApplyTextStyleToRun(run, currentStyle);
				return;
			}

			string tagName = node.Name.ToLowerInvariant();
			switch (tagName)
			{
				case "strong":
				case "b":
					{
						var nextStyle = currentStyle.Clone();
						nextStyle.IsBold = true;
						ProcessInlineNodes(node.ChildNodes, p, doc, nextStyle);
						break;
					}
				case "em":
				case "i":
					{
						var nextStyle = currentStyle.Clone();
						nextStyle.IsItalic = true;
						ProcessInlineNodes(node.ChildNodes, p, doc, nextStyle);
						break;
					}
				case "u":
					{
						var nextStyle = currentStyle.Clone();
						nextStyle.IsUnderline = true;
						ProcessInlineNodes(node.ChildNodes, p, doc, nextStyle);
						break;
					}
				case "span":
					{
						var nextStyle = currentStyle.Clone();
						ParseInlineStyle(node.GetAttributeValue("style", ""), nextStyle);
						ProcessInlineNodes(node.ChildNodes, p, doc, nextStyle);
						break;
					}
				case "br":
					{
						p.CreateRun().AddCarriageReturn();
						break;
					}
				case "img":
					{
						ProcessImageNode(node, p, doc);
						break;
					}
				default:
					// Treat any other tag as transparent container and process its children
					ProcessInlineNodes(node.ChildNodes, p, doc, currentStyle);
					break;
			}
		}

		private void ProcessImageNode(HtmlNode node, XWPFParagraph p, XWPFDocument doc)
		{
			string src = node.GetAttributeValue("src", "");
			if (string.IsNullOrEmpty(src)) return;

			try
			{
				byte[] imageBytes = DocumentConverter.Helpers.ImageHelper.GetImageBytes(src);
				if (imageBytes != null)
				{
					string extension = "png";
					if (src.StartsWith("data:", StringComparison.OrdinalIgnoreCase))
					{
						var match = Regex.Match(src, @"data:image/(?<ext>[^;]+);base64,(?<data>.+)");
						if (match.Success)
						{
							extension = match.Groups["ext"].Value;
						}
					}
					else
					{
						string pathOnly = src.Split('?')[0];
						extension = Path.GetExtension(pathOnly).TrimStart('.');
						if (string.IsNullOrEmpty(extension))
						{
							extension = "png";
						}
					}

					int pictureType = MapExtensionToPictureType(extension);
					var run = p.CreateRun();

					// Default width and height (EMU: English Metric Unit. 1 pixel = 9525 EMU)
					int widthPx = node.GetAttributeValue("width", 300);
					int heightPx = node.GetAttributeValue("height", 200);

					// Try to parse from inline style if present
					string styleAttr = node.GetAttributeValue("style", "");
					if (!string.IsNullOrEmpty(styleAttr))
					{
						var wMatch = Regex.Match(styleAttr, @"width\s*:\s*(?<val>\d+)px");
						if (wMatch.Success) int.TryParse(wMatch.Groups["val"].Value, out widthPx);

						var hMatch = Regex.Match(styleAttr, @"height\s*:\s*(?<val>\d+)px");
						if (hMatch.Success) int.TryParse(hMatch.Groups["val"].Value, out heightPx);
					}

					// Let's fallback to reading image metadata if System.Drawing is available and widthPx/heightPx are defaults
					if (widthPx == 300 && heightPx == 200)
					{
						if (DocumentConverter.Helpers.ImageHelper.TryGetImageDimensions(imageBytes, out int w, out int h))
						{
							widthPx = w;
							heightPx = h;
						}
					}

					int widthEmu = widthPx * 9525;
					int heightEmu = heightPx * 9525;

					using (var imgStream = new MemoryStream(imageBytes))
					{
						run.AddPicture(imgStream, pictureType, $"image.{extension}", widthEmu, heightEmu);
					}
				}
			}
			catch (Exception ex)
			{
				// Keep converting, just add a fallback warning text in the document
				var run = p.CreateRun();
				run.SetText($"[Error loading image: {ex.Message}]");
				run.IsItalic = true;
			}
		}

		private static int MapExtensionToPictureType(string ext)
		{
			ext = ext.ToLowerInvariant();
			switch (ext)
			{
				case "jpeg":
				case "jpg":
					return (int)NPOI.XWPF.UserModel.PictureType.JPEG;
				case "gif":
					return (int)NPOI.XWPF.UserModel.PictureType.GIF;
				case "tiff":
				case "tif":
					return (int)NPOI.XWPF.UserModel.PictureType.TIFF;
				case "bmp":
					return (int)NPOI.XWPF.UserModel.PictureType.BMP;
				default:
					return (int)NPOI.XWPF.UserModel.PictureType.PNG;
			}
		}

		private bool IsInlineElement(HtmlNode node)
		{
			string name = node.Name.ToLowerInvariant();
			return name == "span" || name == "strong" || name == "b" || name == "em" || name == "i" || name == "u" || name == "a" || name == "br" || name == "font";
		}

		private void ApplyParagraphStyles(HtmlNode node, XWPFParagraph p)
		{
			string style = node.GetAttributeValue("style", "");
			if (string.IsNullOrEmpty(style)) return;

			if (style.Contains("text-align: center") || style.Contains("text-align:center"))
			{
				p.Alignment = ParagraphAlignment.CENTER;
			}
			else if (style.Contains("text-align: right") || style.Contains("text-align:right"))
			{
				p.Alignment = ParagraphAlignment.RIGHT;
			}
			else if (style.Contains("text-align: justify") || style.Contains("text-align:justify"))
			{
				p.Alignment = ParagraphAlignment.DISTRIBUTE;
			}
		}

		private void ParseInlineStyle(string styleAttr, TextStyle style)
		{
			if (string.IsNullOrEmpty(styleAttr)) return;

			if (styleAttr.Contains("font-weight: bold") || styleAttr.Contains("font-weight:bold") || styleAttr.Contains("font-weight: 700") || styleAttr.Contains("font-weight:700"))
			{
				style.IsBold = true;
			}
			if (styleAttr.Contains("font-style: italic") || styleAttr.Contains("font-style:italic"))
			{
				style.IsItalic = true;
			}
			if (styleAttr.Contains("text-decoration: underline") || styleAttr.Contains("text-decoration:underline"))
			{
				style.IsUnderline = true;
			}

			// Color parsing (e.g. color: #ff0000 or color: red)
			var colorMatch = Regex.Match(styleAttr, @"color\s*:\s*(?<val>[^;]+)");
			if (colorMatch.Success)
			{
				string colorVal = colorMatch.Groups["val"].Value.Trim().Replace("#", "");
				// Word expects hex color code without #
				if (colorVal.Length == 6 || colorVal.Length == 3)
				{
					style.ColorHex = colorVal;
				}
			}

			// Font size parsing (e.g. font-size: 14px or font-size: 12pt)
			var sizeMatch = Regex.Match(styleAttr, @"font-size\s*:\s*(?<val>\d+)(?<unit>px|pt)");
			if (sizeMatch.Success)
			{
				if (double.TryParse(sizeMatch.Groups["val"].Value, out double size))
				{
					if (sizeMatch.Groups["unit"].Value == "px")
					{
						// convert px to pt approximately
						style.FontSize = size * 0.75;
					}
					else
					{
						style.FontSize = size;
					}
				}
			}
		}

		private void ApplyTextStyleToRun(XWPFRun run, TextStyle style)
		{
			if (style.IsBold) run.IsBold = true;
			if (style.IsItalic) run.IsItalic = true;
			if (style.IsUnderline) run.SetUnderline(UnderlinePatterns.Single);
			if (!string.IsNullOrEmpty(style.ColorHex)) run.SetColor(style.ColorHex);
			if (style.FontSize > 0) run.FontSize = (int)Math.Round(style.FontSize);
		}

		private class TextStyle
		{
			public bool IsBold { get; set; }
			public bool IsItalic { get; set; }
			public bool IsUnderline { get; set; }
			public string ColorHex { get; set; }
			public double FontSize { get; set; }

			public TextStyle Clone()
			{
				return new TextStyle
				{
					IsBold = this.IsBold,
					IsItalic = this.IsItalic,
					IsUnderline = this.IsUnderline,
					ColorHex = this.ColorHex,
					FontSize = this.FontSize
				};
			}
		}
	}
}
