namespace DocumentConverter.Converters
{
	using System;
	using System.Collections.Generic;
	using System.IO;
	using System.Linq;
	using System.Text.RegularExpressions;
	using DocumentConverter.Abstractions;
	using DocumentConverter.Models;
	using HtmlAgilityPack;
	using NPOI.SS.UserModel;
	using NPOI.SS.Util;
	using NPOI.XSSF.UserModel;
	using NPOI.XSSF.UserModel.Extensions;

	/// <summary>
	/// Converts HTML content containing table elements back to Excel document (.xlsx) format using NPOI.
	/// </summary>
	public class HtmlToExcelConverter : IHtmlToDocumentConverter
	{
		private readonly Dictionary<string, ICellStyle> _styleCache = new Dictionary<string, ICellStyle>();

		public Result<byte[]> Convert(string html)
		{
			if (string.IsNullOrEmpty(html))
			{
				return Result<byte[]>.Failure("HTML content cannot be null or empty.");
			}

			try
			{
				_styleCache.Clear();
				var workbook = new XSSFWorkbook();
				var htmlDoc = new HtmlDocument();
				htmlDoc.LoadHtml(html);

				var tableNodes = htmlDoc.DocumentNode.SelectNodes("//table");

				if (tableNodes == null || tableNodes.Count == 0)
				{
					// Fallback: If no tables exist, write all text content into a single cell in a default sheet
					var sheet = workbook.CreateSheet("Sheet1");
					var row = sheet.CreateRow(0);
					var cell = row.CreateCell(0);
					string plainText = htmlDoc.DocumentNode.InnerText?.Trim() ?? "";
					cell.SetCellValue(plainText);

					// Apply auto wrap style
					var style = workbook.CreateCellStyle();
					style.WrapText = true;
					cell.CellStyle = style;
					sheet.SetColumnWidth(0, 10000); // give it some width

					// Scan entire document for images and add them to the sheet
					var allImgNodes = htmlDoc.DocumentNode.SelectNodes("//img");
					if (allImgNodes != null && allImgNodes.Count > 0)
					{
						var images = new List<(byte[] Data, int PictureType)>();
						foreach (var imgNode in allImgNodes)
						{
							var imageData = ExtractImageData(imgNode);
							if (imageData.HasValue)
							{
								images.Add(imageData.Value);
							}
						}
						if (images.Count > 0)
						{
							AddImagesToSheet(images, sheet, workbook, 1);
						}
					}
				}
				else
				{
					int sheetIndex = 1;
					for (int tIdx = 0; tIdx < tableNodes.Count; tIdx++)
					{
						var tableNode = tableNodes[tIdx];

						// Try to get sheet name from a header or table attribute, or use default Sheet1, Sheet2...
						string sheetName = tableNode.GetAttributeValue("name", "");
						if (string.IsNullOrEmpty(sheetName))
						{
							var captionNode = tableNode.SelectSingleNode(".//caption");
							if (captionNode != null)
							{
								sheetName = captionNode.InnerText?.Trim();
							}
						}

						if (string.IsNullOrEmpty(sheetName))
						{
							sheetName = $"Sheet{sheetIndex++}";
						}

						// Sanitize sheet name (Excel names must be unique, <= 31 chars, and cannot contain certain chars)
						sheetName = SanitizeSheetName(sheetName, workbook);

						var sheet = workbook.CreateSheet(sheetName);
						int lastDataRow = ProcessTableToSheet(tableNode, sheet, workbook);

						// For the first table, also collect images that appear before it
						if (tIdx == 0)
						{
							var imagesBefore = CollectImagesFromSiblings(tableNode, false, null);
							if (imagesBefore.Count > 0)
							{
								AddImagesToSheet(imagesBefore, sheet, workbook, lastDataRow + 1);
								lastDataRow += imagesBefore.Count * 12; // approximate row offset
							}
						}

						// Collect images that appear after this table and before the next table
						var images = CollectImagesFromSiblings(tableNode, true, tIdx < tableNodes.Count - 1 ? tableNodes[tIdx + 1] : null);
						if (images.Count > 0)
						{
							AddImagesToSheet(images, sheet, workbook, lastDataRow + 1);
						}
					}
				}

				using (var ms = new MemoryStream())
				{
					workbook.Write(ms);
					return Result<byte[]>.Success(ms.ToArray());
				}
			}
			catch (Exception ex)
			{
				return Result<byte[]>.Failure($"HTML to Excel conversion failed: {ex.Message}");
			}
		}

		/// <summary>
		/// Processes a table node and writes its data to the given sheet.
		/// Returns the index of the last row written.
		/// </summary>
		private int ProcessTableToSheet(HtmlNode tableNode, ISheet sheet, IWorkbook workbook)
		{
			var rows = tableNode.SelectNodes(".//tr");
			if (rows == null || rows.Count == 0) return 0;

			// Grid of occupied cell positions: tracks where merged regions are so we don't place cells over them
			var occupied = new HashSet<(int r, int c)>();

			// Collect (imgNode, rowIndex, colIndex, colSpan, rowSpan) so images are anchored to the correct cell boundaries
			var cellImagesWithCoords = new List<(HtmlNode ImgNode, int RowIndex, int ColIndex, int ColSpan, int RowSpan)>();

			for (int rIndex = 0; rIndex < rows.Count; rIndex++)
			{
				var rowNode = rows[rIndex];
				var cells = rowNode.SelectNodes("th|td");
				if (cells == null) continue;

				IRow excelRow = sheet.GetRow(rIndex) ?? sheet.CreateRow(rIndex);
				int cIndex = 0;

				foreach (var cellNode in cells)
				{
					// Find next unoccupied column index in the current row
					while (occupied.Contains((rIndex, cIndex)))
					{
						cIndex++;
					}

					// Get cell row and column span
					int colspan = cellNode.GetAttributeValue("colspan", 1);
					int rowspan = cellNode.GetAttributeValue("rowspan", 1);

					// Apply merges for colspan or rowspan
					if (colspan > 1 || rowspan > 1)
					{
						int lastRow = rIndex + rowspan - 1;
						int lastCol = cIndex + colspan - 1;

						// Mark all merged coordinates as occupied
						for (int r = rIndex; r <= lastRow; r++)
						{
							for (int c = cIndex; c <= lastCol; c++)
							{
								occupied.Add((r, c));
							}
						}

						sheet.AddMergedRegion(new CellRangeAddress(rIndex, lastRow, cIndex, lastCol));
					}
					else
					{
						occupied.Add((rIndex, cIndex));
					}

					// Collect any img tags in this cell, tagging them with this row index and column/span coordinates
					var imgNodes = cellNode.SelectNodes(".//img");
					if (imgNodes != null)
					{
						foreach (var imgNode in imgNodes)
						{
							cellImagesWithCoords.Add((imgNode, rIndex, cIndex, colspan, rowspan));
						}
					}

					// Create cell
					ICell excelCell = excelRow.CreateCell(cIndex);

					// Cell type and styling
					bool isHeader = cellNode.Name.ToLowerInvariant() == "th";
					string text = cellNode.InnerText?.Trim() ?? "";

					// Parse cell style attributes
					string styleAttr = cellNode.GetAttributeValue("style", "");
					ApplyCellStyle(excelCell, workbook, isHeader, styleAttr);

					// Set cell value with type inference
					SetCellValueInferType(excelCell, text);

					cIndex += colspan;
				}
			}

			// Auto-size columns
			int maxCol = occupied.Count > 0 ? occupied.Max(p => p.c) : 0;
			for (int i = 0; i <= maxCol; i++)
			{
				try
				{
					sheet.AutoSizeColumn(i);
				}
				catch
				{
					// Fallback to default width if auto-size fails
					sheet.SetColumnWidth(i, 4000);
				}
			}

			// Embed images at their source cell boundaries
			if (cellImagesWithCoords.Count > 0)
			{
				// Create drawing patriarch ONCE for the sheet
				IDrawing drawing = sheet.CreateDrawingPatriarch();

				foreach (var (imgNode, rowIdx, colIdx, colSpan, rowSpan) in cellImagesWithCoords)
				{
					var imageData = ExtractImageData(imgNode);
					if (!imageData.HasValue) continue;

					EmbedImageInCell(drawing, workbook, sheet, imageData.Value.Data, imageData.Value.PictureType, rowIdx, colIdx, colSpan, rowSpan);
				}
			}

			return rows.Count - 1;
		}

		/// <summary>
		/// Collects img elements from siblings in the DOM.
		/// </summary>
		private List<(byte[] Data, int PictureType)> CollectImagesFromSiblings(HtmlNode startNode, bool forward, HtmlNode stopAtNode = null)
		{
			var images = new List<(byte[] Data, int PictureType)>();
			try
			{
				var sibling = forward ? startNode.NextSibling : startNode.PreviousSibling;
				var imgNodesFound = new List<HtmlNode>();

				while (sibling != null)
				{
					if (sibling == stopAtNode) break;
					if (sibling.Name?.ToLowerInvariant() == "table") break;

					if (sibling.Name?.ToLowerInvariant() == "img")
					{
						imgNodesFound.Add(sibling);
					}
					else
					{
						var nestedImgs = sibling.SelectNodes(".//img");
						if (nestedImgs != null)
						{
							imgNodesFound.AddRange(nestedImgs);
						}
					}

					sibling = forward ? sibling.NextSibling : sibling.PreviousSibling;
				}

				if (!forward)
				{
					// Reverse to maintain original DOM order
					imgNodesFound.Reverse();
				}

				foreach (var imgNode in imgNodesFound)
				{
					var imageData = ExtractImageData(imgNode);
					if (imageData.HasValue)
					{
						images.Add(imageData.Value);
					}
				}
			}
			catch
			{
				// Ignore errors in image collection
			}

			return images;
		}

		/// <summary>
		/// Extracts image bytes and NPOI picture type from an img HTML node.
		/// </summary>
		private (byte[] Data, int PictureType)? ExtractImageData(HtmlNode imgNode)
		{
			string src = imgNode.GetAttributeValue("src", "");
			if (string.IsNullOrEmpty(src)) return null;

			try
			{
				byte[] imageBytes = DocumentConverter.Helpers.ImageHelper.GetImageBytes(src);
				if (imageBytes != null && imageBytes.Length > 0)
				{
					string extension = "png";
					if (src.StartsWith("data:", StringComparison.OrdinalIgnoreCase))
					{
						var match = Regex.Match(src, @"data:image/(?<ext>[^;]+);base64,(?<data>.+)");
						if (match.Success)
						{
							extension = match.Groups["ext"].Value.ToLowerInvariant();
						}
					}
					else
					{
						string pathOnly = src.Split('?')[0];
						extension = Path.GetExtension(pathOnly).TrimStart('.').ToLowerInvariant();
						if (string.IsNullOrEmpty(extension))
						{
							extension = "png";
						}
					}

					int pictureType = MapExtensionToNpoiPictureType(extension);
					return (imageBytes, pictureType);
				}
			}
			catch
			{
				// Ignore individual image extraction errors
			}

			return null;
		}

		/// <summary>
		/// Adds a list of images to the sheet, starting at the given row index.
		/// </summary>
		private void AddImagesToSheet(List<(byte[] Data, int PictureType)> images, ISheet sheet, IWorkbook workbook, int startRow)
		{
			if (images == null || images.Count == 0) return;

			int currentRow = startRow;
			IDrawing drawing = sheet.CreateDrawingPatriarch();

			foreach (var image in images)
			{
				int rowSpan = EmbedImageOnSheet(drawing, workbook, sheet, image.Data, image.PictureType, currentRow);
				currentRow += rowSpan + 1; // Move to next available row after this image
			}
		}

		/// <summary>
		/// Embeds a single image on the sheet and returns the number of rows it spanned.
		/// </summary>
		private int EmbedImageOnSheet(IDrawing drawing, IWorkbook workbook, ISheet sheet, byte[] imgBytes, int pictureType, int startRow)
		{
			try
			{
				int pictureIndex = workbook.AddPicture(imgBytes, (PictureType)pictureType);

				// Determine image dimensions
				int imgWidthPx = 300, imgHeightPx = 200;
				if (DocumentConverter.Helpers.ImageHelper.TryGetImageDimensions(imgBytes, out int w, out int h))
				{
					imgWidthPx = w;
					imgHeightPx = h;
				}

				// Constrain width
				const int maxWidth = 800;
				if (imgWidthPx > maxWidth)
				{
					double ratio = (double)maxWidth / imgWidthPx;
					imgWidthPx = maxWidth;
					imgHeightPx = (int)(imgHeightPx * ratio);
				}

				int colSpan = Math.Max(1, (imgWidthPx / 64) + 1);
				int rowSpan = Math.Max(1, (imgHeightPx / 20) + 1);

				// Expand row heights
				for (int r = startRow; r < startRow + rowSpan; r++)
				{
					IRow existingRow = sheet.GetRow(r) ?? sheet.CreateRow(r);
					float neededHeight = Math.Max(existingRow.HeightInPoints, imgHeightPx / (float)rowSpan);
					existingRow.HeightInPoints = neededHeight;
				}

				// Anchor the image
				IClientAnchor anchor = workbook.GetCreationHelper().CreateClientAnchor();
				anchor.Col1 = 0;
				anchor.Row1 = startRow;
				anchor.Col2 = colSpan;
				anchor.Row2 = startRow + rowSpan;
				anchor.AnchorType = AnchorType.MoveAndResize;

				drawing.CreatePicture(anchor, pictureIndex);
				return rowSpan;
			}
			catch
			{
				return 0;
			}
		}

		/// <summary>
		/// Embeds a single image inside the boundaries of a specific table cell.
		/// </summary>
		private void EmbedImageInCell(IDrawing drawing, IWorkbook workbook, ISheet sheet, byte[] imgBytes, int pictureType, int rowIdx, int colIdx, int colSpan, int rowSpan)
		{
			try
			{
				int pictureIndex = workbook.AddPicture(imgBytes, (PictureType)pictureType);

				// Determine image dimensions
				int imgWidthPx = 300, imgHeightPx = 200;
				if (DocumentConverter.Helpers.ImageHelper.TryGetImageDimensions(imgBytes, out int w, out int h))
				{
					imgWidthPx = w;
					imgHeightPx = h;
				}

				// Expand the row heights of the spanned cell rows so the image fits inside
				for (int r = rowIdx; r < rowIdx + rowSpan; r++)
				{
					IRow existingRow = sheet.GetRow(r) ?? sheet.CreateRow(r);
					float neededHeight = Math.Max(existingRow.HeightInPoints, imgHeightPx / (float)rowSpan);
					existingRow.HeightInPoints = neededHeight;
				}

				// Anchor the image exactly inside the cell boundaries
				IClientAnchor anchor = workbook.GetCreationHelper().CreateClientAnchor();
				anchor.Col1 = colIdx;
				anchor.Row1 = rowIdx;
				anchor.Col2 = colIdx + colSpan;
				anchor.Row2 = rowIdx + rowSpan;
				anchor.AnchorType = AnchorType.MoveAndResize;

				drawing.CreatePicture(anchor, pictureIndex);
			}
			catch
			{
				// Skip image on failure
			}
		}

		private static int MapExtensionToNpoiPictureType(string ext)
		{
			ext = ext?.ToLowerInvariant() ?? "png";
			switch (ext)
			{
				case "jpeg":
				case "jpg":
					return (int)PictureType.JPEG;
				case "gif":
					return (int)PictureType.GIF;
				case "bmp":
				case "x-bmp":
					return (int)PictureType.BMP;
				case "tiff":
				case "tif":
					return (int)PictureType.TIFF;
				case "emf":
				case "x-emf":
					return (int)PictureType.EMF;
				case "wmf":
				case "x-wmf":
					return (int)PictureType.WMF;
				case "png":
				default:
					return (int)PictureType.PNG;
			}
		}

		private void ApplyCellStyle(ICell cell, IWorkbook workbook, bool isHeader, string styleAttr)
		{
			string cacheKey = $"{isHeader}|{styleAttr}";
			if (_styleCache.TryGetValue(cacheKey, out var cachedStyle))
			{
				cell.CellStyle = cachedStyle;
				return;
			}

			ICellStyle cellStyle = workbook.CreateCellStyle();
			IFont font = workbook.CreateFont();

			if (isHeader)
			{
				font.IsBold = true;
				// Light gray background for headers
				cellStyle.FillForegroundColor = IndexedColors.Grey25Percent.Index;
				cellStyle.FillPattern = FillPattern.SolidForeground;
			}

			// Parse inline styles
			if (!string.IsNullOrEmpty(styleAttr))
			{
				if (styleAttr.Contains("font-weight: bold") || styleAttr.Contains("font-weight:bold") || styleAttr.Contains("font-weight: 700") || styleAttr.Contains("font-weight:700"))
				{
					font.IsBold = true;
				}

				if (styleAttr.Contains("font-style: italic") || styleAttr.Contains("font-style:italic"))
				{
					font.IsItalic = true;
				}

				// Check text alignments
				if (styleAttr.Contains("text-align: center") || styleAttr.Contains("text-align:center"))
				{
					cellStyle.Alignment = HorizontalAlignment.Center;
				}
				else if (styleAttr.Contains("text-align: right") || styleAttr.Contains("text-align:right"))
				{
					cellStyle.Alignment = HorizontalAlignment.Right;
				}
				else if (styleAttr.Contains("text-align: left") || styleAttr.Contains("text-align:left"))
				{
					cellStyle.Alignment = HorizontalAlignment.Left;
				}

				// Check background color
				var bgMatch = Regex.Match(styleAttr, @"background(-color)?\s*:\s*(?<val>[^;]+)");
				if (bgMatch.Success)
				{
					string colorVal = bgMatch.Groups["val"].Value.Trim().ToLowerInvariant();
					if (colorVal.StartsWith("#"))
					{
						ApplyHexColorToStyle(workbook, cellStyle, colorVal.Replace("#", ""));
					}
				}
			}

			// Add borders
			cellStyle.BorderBottom = BorderStyle.Thin;
			cellStyle.BorderTop = BorderStyle.Thin;
			cellStyle.BorderLeft = BorderStyle.Thin;
			cellStyle.BorderRight = BorderStyle.Thin;

			cellStyle.SetFont(font);
			_styleCache[cacheKey] = cellStyle;
			cell.CellStyle = cellStyle;
		}

		private void ApplyHexColorToStyle(IWorkbook workbook, ICellStyle cellStyle, string hexColor)
		{
			hexColor = hexColor.Trim().Replace("#", "");
			if (cellStyle is XSSFCellStyle xssfStyle)
			{
				byte[] rgb = ParseHexColor(hexColor);
				if (rgb != null)
				{
					xssfStyle.SetFillForegroundColor(new XSSFColor(rgb));
					xssfStyle.FillPattern = FillPattern.SolidForeground;
					return;
				}
			}

			// Fallback for HSSF/other styles to standard IndexedColors
			if (hexColor.StartsWith("ff0000", StringComparison.OrdinalIgnoreCase) || hexColor.Equals("f00", StringComparison.OrdinalIgnoreCase))
			{
				cellStyle.FillForegroundColor = IndexedColors.Red.Index;
				cellStyle.FillPattern = FillPattern.SolidForeground;
			}
			else if (hexColor.StartsWith("00ff00", StringComparison.OrdinalIgnoreCase) || hexColor.Equals("0f0", StringComparison.OrdinalIgnoreCase))
			{
				cellStyle.FillForegroundColor = IndexedColors.BrightGreen.Index;
				cellStyle.FillPattern = FillPattern.SolidForeground;
			}
			else if (hexColor.StartsWith("0000ff", StringComparison.OrdinalIgnoreCase) || hexColor.Equals("00f", StringComparison.OrdinalIgnoreCase))
			{
				cellStyle.FillForegroundColor = IndexedColors.Blue.Index;
				cellStyle.FillPattern = FillPattern.SolidForeground;
			}
			else if (hexColor.StartsWith("ffff00", StringComparison.OrdinalIgnoreCase) || hexColor.Equals("ff0", StringComparison.OrdinalIgnoreCase))
			{
				cellStyle.FillForegroundColor = IndexedColors.Yellow.Index;
				cellStyle.FillPattern = FillPattern.SolidForeground;
			}
			else if (hexColor.StartsWith("cccccc", StringComparison.OrdinalIgnoreCase) || hexColor.StartsWith("d3d3d3", StringComparison.OrdinalIgnoreCase))
			{
				cellStyle.FillForegroundColor = IndexedColors.Grey25Percent.Index;
				cellStyle.FillPattern = FillPattern.SolidForeground;
			}
		}

		private byte[] ParseHexColor(string hexColor)
		{
			try
			{
				hexColor = hexColor.Trim().Replace("#", "");
				if (hexColor.Length == 3)
				{
					char r = hexColor[0];
					char g = hexColor[1];
					char b = hexColor[2];
					hexColor = new string(new[] { r, r, g, g, b, b });
				}

				if (hexColor.Length == 6)
				{
					byte r = System.Convert.ToByte(hexColor.Substring(0, 2), 16);
					byte g = System.Convert.ToByte(hexColor.Substring(2, 2), 16);
					byte b = System.Convert.ToByte(hexColor.Substring(4, 2), 16);
					return new byte[] { r, g, b };
				}
			}
			catch { }
			return null;
		}

		private void SetCellValueInferType(ICell cell, string text)
		{
			if (string.IsNullOrEmpty(text))
			{
				cell.SetCellValue("");
				return;
			}

			// Try to parse as double/int
			if (double.TryParse(text, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double dVal))
			{
				cell.SetCellValue(dVal);
				cell.SetCellType(CellType.Numeric);
				return;
			}

			// Try to parse as bool
			if (bool.TryParse(text, out bool bVal))
			{
				cell.SetCellValue(bVal);
				cell.SetCellType(CellType.Boolean);
				return;
			}

			// Try to parse as DateTime (standard ISO format)
			if (DateTime.TryParse(text, System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out DateTime dtVal))
			{
				cell.SetCellValue(dtVal);
				return;
			}

			// Fallback to text
			cell.SetCellValue(text);
			cell.SetCellType(CellType.String);
		}

		private string SanitizeSheetName(string name, IWorkbook workbook)
		{
			if (string.IsNullOrEmpty(name)) return "Sheet";

			// Remove forbidden characters: \ / ? * : [ ]
			string clean = Regex.Replace(name, @"[\\/\?\*\:\[\]]", "");

			// Limit length to 30 characters
			if (clean.Length > 30)
			{
				clean = clean.Substring(0, 30);
			}

			// Ensure unique sheet name in workbook
			string uniqueName = clean;
			int counter = 1;
			while (workbook.GetSheet(uniqueName) != null)
			{
				string suffix = $"({counter++})";
				if (clean.Length + suffix.Length > 30)
				{
					uniqueName = clean.Substring(0, 30 - suffix.Length) + suffix;
				}
				else
				{
					uniqueName = clean + suffix;
				}
			}

			return uniqueName;
		}
	}
}
