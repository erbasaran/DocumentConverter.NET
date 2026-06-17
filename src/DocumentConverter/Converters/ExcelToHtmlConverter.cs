namespace DocumentConverter.Converters
{
	using System;
	using System.Collections.Generic;
	using System.IO;
	using System.Security.Cryptography;
	using System.Text;
	using DocumentConverter.Abstractions;
	using DocumentConverter.Models;
	using NPOI.HSSF.UserModel;
	using NPOI.HSSF.Util;
	using NPOI.SS.UserModel;
	using NPOI.XSSF.UserModel;

	/// <summary>
	/// Converts legacy (.xls) and modern (.xlsx) Excel documents to HTML format.
	/// </summary>
	public class ExcelToHtmlConverter : IDocumentConverter
	{
		private readonly bool _includeHtmlWrapper;

		public ExcelToHtmlConverter(bool includeHtmlWrapper = false)
		{
			_includeHtmlWrapper = includeHtmlWrapper;
		}

		private static string GetPictureHash(byte[] rawBytes)
		{
			using (var sha = SHA1.Create())
			{
				return System.Convert.ToBase64String(sha.ComputeHash(rawBytes));
			}
		}

		private class RowObject
		{
			public int RowIndex { get; set; }
			public RowObjectType Type { get; set; }
			public string Text { get; set; }
			public string Style { get; set; }
			public byte[] PictureData { get; set; }
			public string PictureMime { get; set; }
			public string RawHash { get; set; }
			public IRow Row { get; set; }
		}

		private enum RowObjectType
		{
			Text,
			Picture,
			Data
		}

		public Result<string> Convert(Stream stream)
		{
			var resourcesToClose = new System.Collections.Generic.List<object>();
			try
			{
				byte[] rawBytes;
				using (var memoryStream = new MemoryStream())
				{
					stream.CopyTo(memoryStream);
					rawBytes = memoryStream.ToArray();
				}

				IWorkbook workbook;
				try
				{
					var ms = new MemoryStream(rawBytes);
					resourcesToClose.Add(ms);
					workbook = WorkbookFactory.Create(ms);
					resourcesToClose.Add(workbook);
				}
				catch (Exception ex)
				{
					try
					{
						var poifs = new NPOI.POIFS.FileSystem.POIFSFileSystem(new MemoryStream(rawBytes));
						resourcesToClose.Add(poifs);

						string entryName = null;
						if (poifs.Root.HasEntry("Workbook"))
						{
							entryName = "Workbook";
						}
						else if (poifs.Root.HasEntry("Book"))
						{
							entryName = "Book";
						}
						if (entryName != null)
						{
							byte[] workbookData;
							using (var documentStream = poifs.CreateDocumentInputStream(entryName))
							{
								workbookData = new byte[documentStream.Length];
								documentStream.Read(workbookData, 0, workbookData.Length);
							}
							byte[] cleanedData = CleanWorkbookStream(workbookData);
							var cleanPoifs = new NPOI.POIFS.FileSystem.POIFSFileSystem();
							resourcesToClose.Add(cleanPoifs);

							cleanPoifs.CreateDocument(new MemoryStream(cleanedData), entryName);
							CopyPoifsEntry(poifs, cleanPoifs, "\x05SummaryInformation");
							CopyPoifsEntry(poifs, cleanPoifs, "\x05DocumentSummaryInformation");

							var outputMs = new MemoryStream();
							resourcesToClose.Add(outputMs);
							cleanPoifs.WriteFileSystem(outputMs);
							outputMs.Position = 0;

							var hssfWorkbook = new HSSFWorkbook(outputMs);
							resourcesToClose.Add(hssfWorkbook);
							workbook = hssfWorkbook;
						}
						else
						{
							throw new Exception("Workbook entry not found in OLE2 container.", ex);
						}
					}
					catch (Exception innerEx)
					{
						return Result<string>.Failure($"Excel conversion failed: {ex.Message} (Fallback cleanup failed: {innerEx.Message})");
					}
				}

				StringBuilder htmlBuilder = new StringBuilder();

				if (_includeHtmlWrapper)
				{
					htmlBuilder.AppendLine("<!DOCTYPE html>");
					htmlBuilder.AppendLine("<html>");
					htmlBuilder.AppendLine("<head>");
					htmlBuilder.AppendLine("<meta charset=\"utf-8\" />");
					htmlBuilder.AppendLine("<title>Converted Excel Document</title>");
					htmlBuilder.AppendLine("<style>");
					htmlBuilder.AppendLine("  body { font-family: -apple-system, BlinkMacSystemFont, \"Segoe UI\", Roboto, Helvetica, Arial, sans-serif; color: #334155; padding: 20px; max-width: 1200px; margin: 0 auto; }");
					htmlBuilder.AppendLine("  h2 { color: #0f172a; font-size: 1.5rem; font-weight: 600; margin-top: 24px; margin-bottom: 12px; }");
					htmlBuilder.AppendLine("  h3 { color: #1e293b; font-size: 1.2rem; font-weight: 600; margin-top: 20px; margin-bottom: 8px; }");
					htmlBuilder.AppendLine("  p { line-height: 1.6; margin-bottom: 16px; }");
					htmlBuilder.AppendLine("  table { border-collapse: collapse; width: auto; margin: 24px auto; min-width: 60%; box-shadow: 0 1px 3px 0 rgba(0, 0, 0, 0.1), 0 1px 2px 0 rgba(0, 0, 0, 0.06); border-radius: 4px; overflow: hidden; }");
					htmlBuilder.AppendLine("  th, td { padding: 12px 16px; min-width: 60px; vertical-align: middle; font-size: 0.9rem; border-bottom: 1px solid #e2e8f0; text-align: left; }");
					htmlBuilder.AppendLine("  th { background-color: #f8fafc; font-weight: 600; color: #334155; border-bottom: 2px solid #cbd5e1; }");
					htmlBuilder.AppendLine("  tr:last-child td { border-bottom: none; }");
					htmlBuilder.AppendLine("  tr:hover td { background-color: #f8fafc; }");
					htmlBuilder.AppendLine("</style>");
					htmlBuilder.AppendLine("</head>");
					htmlBuilder.AppendLine("<body>");
				}

				DataFormatter dataFormatter = new DataFormatter();
				HashSet<string> renderedPicHashes = new HashSet<string>();

				for (int s = 0; s < workbook.NumberOfSheets; s++)
				{
					ISheet sheet = workbook.GetSheetAt(s);
					if (sheet == null) continue;

					int lastRowNum = sheet.LastRowNum;
					if (lastRowNum < 0)
					{
						htmlBuilder.AppendLine("  <p><em>This sheet is empty.</em></p>");
						continue;
					}

					// 1. Extract drawing pictures for this sheet
					var sheetPictures = GetSheetPictures(sheet);

					// Find max columns to keep table structure grid uniform
					int maxCols = 0;
					for (int r = 0; r <= lastRowNum; r++)
					{
						IRow row = sheet.GetRow(r);
						if (row != null && row.LastCellNum > maxCols)
						{
							maxCols = row.LastCellNum;
						}
					}

					// Find active row indices
					var activeRows = new List<int>();
					for (int r = 0; r <= lastRowNum; r++)
					{
						if (IsRowNeeded(sheet, r, maxCols))
						{
							activeRows.Add(r);
						}
					}

					if (activeRows.Count == 0)
					{
						htmlBuilder.AppendLine("  <p><em>This sheet is empty.</em></p>");
						continue;
					}

					// Find active columns range [minCol, maxCol]
					int minCol = maxCols;
					int maxCol = 0;
					foreach (int r in activeRows)
					{
						IRow row = sheet.GetRow(r);
						if (row == null) continue;
						for (int c = 0; c < maxCols; c++)
						{
							ICell cell = row.GetCell(c);
							if (cell != null && !string.IsNullOrWhiteSpace(cell.ToString()))
							{
								if (c < minCol) minCol = c;
								if (c > maxCol) maxCol = c;
							}
						}
					}

					if (minCol > maxCol)
					{
						htmlBuilder.AppendLine("  <p><em>This sheet is empty.</em></p>");
						continue;
					}

					int numMerged = sheet.NumMergedRegions;

					// Render table
					htmlBuilder.AppendLine("  <table>");
					foreach (int r in activeRows)
					{
						IRow row = sheet.GetRow(r);
						htmlBuilder.AppendLine("    <tr>");

						// Check if this row has exactly one non-empty cell
						int cellCount = 0;
						ICell singleCell = null;
						int singleCellCol = -1;
						if (row != null)
						{
							for (int c = minCol; c <= maxCol; c++)
							{
								ICell cell = row.GetCell(c);
								if (cell != null && !string.IsNullOrWhiteSpace(cell.ToString()))
								{
									cellCount++;
									singleCell = cell;
									singleCellCol = c;
								}
							}
						}

						// Check if the single cell is part of any merged region
						bool isSingleCellMerged = false;
						if (cellCount == 1 && singleCell != null)
						{
							for (int i = 0; i < numMerged; i++)
							{
								var region = sheet.GetMergedRegion(i);
								if (region.IsInRange(r, singleCellCol))
								{
									isSingleCellMerged = true;
									break;
								}
							}
						}

						if (cellCount == 1 && singleCell != null && !isSingleCellMerged)
						{
							// Render empty cells for columns before singleCellCol
							for (int c = minCol; c < singleCellCol; c++)
							{
								htmlBuilder.AppendLine("      <td></td>");
							}

							string cellText = dataFormatter.FormatCellValue(singleCell);
							string encodedText = SafeHtmlEncode(cellText);
							int colSpan = maxCol - singleCellCol + 1;
							string spans = colSpan > 1 ? $" colspan=\"{colSpan}\"" : "";

							string bgStyle = GetCellBackgroundColor(singleCell);
							string fontStyle = GetCellFontStyle(singleCell);
							string alignStyle = GetCellAlignment(singleCell);

							List<string> styles = new List<string>();
							if (!string.IsNullOrEmpty(bgStyle)) styles.Add(bgStyle);
							if (!string.IsNullOrEmpty(fontStyle)) styles.Add(fontStyle);
							if (!string.IsNullOrEmpty(alignStyle)) styles.Add(alignStyle);

							string styleAttr = styles.Count > 0 ? $" style=\"{string.Join(" ", styles)}\"" : "";
							htmlBuilder.AppendLine($"      <td{spans}{styleAttr}>{encodedText}</td>");
						}
						else
						{
							for (int c = minCol; c <= maxCol; c++)
							{
								if (GetMergedCellInfo(sheet, activeRows, minCol, maxCol, r, c, out int rowSpan, out int colSpan, out bool isTopLeft))
								{
									if (!isTopLeft)
									{
										continue;
									}
								}

								ICell cell = row?.GetCell(c);
								string cellText = cell != null ? dataFormatter.FormatCellValue(cell) : string.Empty;
								string encodedText = SafeHtmlEncode(cellText);

								string spans = "";
								if (rowSpan > 1) spans += $" rowspan=\"{rowSpan}\"";
								if (colSpan > 1) spans += $" colspan=\"{colSpan}\"";

								string bgStyle = GetCellBackgroundColor(cell);
								string fontStyle = GetCellFontStyle(cell);
								string alignStyle = GetCellAlignment(cell);

								List<string> styles = new List<string>();
								if (!string.IsNullOrEmpty(bgStyle)) styles.Add(bgStyle);
								if (!string.IsNullOrEmpty(fontStyle)) styles.Add(fontStyle);
								if (!string.IsNullOrEmpty(alignStyle)) styles.Add(alignStyle);

								string styleAttr = styles.Count > 0 ? $" style=\"{string.Join(" ", styles)}\"" : "";
								htmlBuilder.AppendLine($"      <td{spans}{styleAttr}>{encodedText}</td>");
							}
						}

						htmlBuilder.AppendLine("    </tr>");
					}
					htmlBuilder.AppendLine("  </table>");

					// Render sheet pictures at the bottom of the table
					if (sheetPictures != null && sheetPictures.Count > 0)
					{
						foreach (var pic in sheetPictures.Values)
						{
							if (!renderedPicHashes.Contains(pic.RawHash))
							{
								string base64 = System.Convert.ToBase64String(pic.Data);
								htmlBuilder.AppendLine($"  <div style=\"text-align: center; margin: 24px 0;\">");
								htmlBuilder.AppendLine($"    <img src=\"data:{pic.Mime};base64,{base64}\" style=\"max-width: 100%; height: auto; display: inline-block;\" />");
								htmlBuilder.AppendLine($"  </div>");
								renderedPicHashes.Add(pic.RawHash);
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
				return Result<string>.Failure($"Excel conversion failed: {ex.Message}");
			}
			finally
			{
				foreach (var res in resourcesToClose)
				{
					if (res == null) continue;
					try
					{
						if (res is IDisposable disp)
						{
							disp.Dispose();
						}
						else
						{
							var closeMethod = res.GetType().GetMethod("Close");
							closeMethod?.Invoke(res, null);
						}
					}
					catch { }
				}
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

		private bool IsCellInvisible(ICell cell)
		{
			if (cell == null) return true;
			if (string.IsNullOrWhiteSpace(cell.ToString())) return true;

			var style = cell.CellStyle;
			if (style == null) return false;

			var font = cell.Sheet.Workbook.GetFontAt(style.FontIndex);
			if (font == null) return false;

			string fontColor = GetFontColor(font, cell.Sheet.Workbook);
			if (fontColor == null) return false;

			if (fontColor.Equals("#FFFFFF", StringComparison.OrdinalIgnoreCase))
			{
				string bgColor = GetCellBackgroundColor(cell);
				if (bgColor == null || bgColor.IndexOf("#FFFFFF", StringComparison.OrdinalIgnoreCase) >= 0)
				{
					return true;
				}
			}

			return false;
		}

		private bool IsRowNeeded(ISheet sheet, int rowIdx, int maxCols)
		{
			IRow row = sheet.GetRow(rowIdx);
			if (row == null) return false;

			bool hasVisibleContent = false;

			// Check if there are any cells with text in this row
			for (int c = 0; c < maxCols; c++)
			{
				ICell cell = row.GetCell(c);
				if (cell != null && !string.IsNullOrWhiteSpace(cell.ToString()))
				{
					if (!IsCellInvisible(cell))
					{
						hasVisibleContent = true;
						break;
					}
				}
			}

			if (hasVisibleContent) return true;

			// Check if this row is part of any non-empty merged region
			int numMerged = sheet.NumMergedRegions;
			for (int i = 0; i < numMerged; i++)
			{
				var region = sheet.GetMergedRegion(i);
				if (rowIdx >= region.FirstRow && rowIdx <= region.LastRow)
				{
					ICell topLeftCell = sheet.GetRow(region.FirstRow)?.GetCell(region.FirstColumn);
					if (topLeftCell != null && !string.IsNullOrWhiteSpace(topLeftCell.ToString()))
					{
						if (!IsCellInvisible(topLeftCell))
						{
							return true;
						}
					}
				}
			}

			return false;
		}

		private Dictionary<string, (byte[] Data, string Mime, string RawHash)> GetSheetPictures(ISheet sheet)
		{
			var dict = new Dictionary<string, (byte[] Data, string Mime, string RawHash)>();
			try
			{
				if (sheet is HSSFSheet hssfSheet)
				{
					var patriarch = hssfSheet.DrawingPatriarch as HSSFPatriarch;
					if (patriarch != null)
					{
						var shapes = patriarch.GetShapes();
						if (shapes != null)
						{
							foreach (var shape in shapes)
							{
								ExtractHSSFPictures(shape, dict, hssfSheet.Workbook);
							}
						}
					}
				}
				else if (sheet is XSSFSheet xssfSheet)
				{
					var patriarch = xssfSheet.DrawingPatriarch as XSSFDrawing;
					if (patriarch != null)
					{
						var shapes = patriarch.GetShapes();
						if (shapes != null)
						{
							foreach (var shape in shapes)
							{
								ExtractXSSFPictures(shape, dict);
							}
						}
					}
				}
			}
			catch
			{
				// Ignore drawing extraction failures
			}
			return dict;
		}

		private (byte[] Data, string Mime) ProcessPictureData(byte[] rawData, PictureType type)
		{
			string ext = GetExtensionFromType(type);
			string mime;
			byte[] processedData = DocumentConverter.Helpers.ImageHelper.ProcessImage(rawData, ext, out mime);
			return (processedData, mime);
		}

		private void ExtractHSSFPictures(HSSFShape shape, Dictionary<string, (byte[] Data, string Mime, string RawHash)> dict, IWorkbook workbook)
		{
			if (shape is HSSFPicture picture && !(shape is HSSFObjectData))
			{
				var anchor = picture.ClientAnchor;
				if (anchor != null)
				{
					int r = anchor.Row1;
					int c = anchor.Col1;
					byte[] rawBytes = picture.PictureData.Data;
					var processed = ProcessPictureData(rawBytes, picture.PictureData.PictureType);
					string rawHash = GetPictureHash(rawBytes);
					dict[$"{r}_{c}"] = (processed.Data, processed.Mime, rawHash);
				}
			}
			else if (shape is HSSFShapeGroup group)
			{
				if (group.Children != null)
				{
					foreach (var child in group.Children)
					{
						ExtractHSSFPictures(child, dict, workbook);
					}
				}
			}
		}

		private void ExtractXSSFPictures(IShape shape, Dictionary<string, (byte[] Data, string Mime, string RawHash)> dict)
		{
			if (shape is XSSFPicture picture)
			{
				var anchor = picture.ClientAnchor;
				if (anchor != null)
				{
					int r = anchor.Row1;
					int c = anchor.Col1;
					byte[] rawBytes = picture.PictureData.Data;
					var processed = ProcessPictureData(rawBytes, picture.PictureData.PictureType);
					string rawHash = GetPictureHash(rawBytes);
					dict[$"{r}_{c}"] = (processed.Data, processed.Mime, rawHash);
				}
			}
		}

		private int GetChartShapeIndex(HSSFShape shape)
		{
			var patriarch = shape.Patriarch;
			if (patriarch == null) return -1;

			var shapes = patriarch.GetShapes();
			int index = 0;
			if (FindChartShapeIndexRecursive(shapes, shape, ref index))
			{
				return index;
			}
			return -1;
		}

		private bool FindChartShapeIndexRecursive(IList<HSSFShape> shapes, HSSFShape target, ref int index)
		{
			if (shapes == null) return false;
			foreach (var s in shapes)
			{
				if (s == target)
				{
					return true;
				}
				if (s is HSSFSimpleShape simple && simple.ShapeType == 201)
				{
					index++;
				}
				else if (s is HSSFShapeGroup group)
				{
					if (FindChartShapeIndexRecursive(group.Children, target, ref index))
					{
						return true;
					}
				}
			}
			return false;
		}

		private string GetExtensionFromType(PictureType type)
		{
			switch (type)
			{
				case PictureType.EMF: return "emf";
				case PictureType.WMF: return "wmf";
				case PictureType.PICT: return "pict";
				case PictureType.JPEG: return "jpeg";
				case PictureType.PNG: return "png";
				case PictureType.BMP: return "bmp";
				default: return "png";
			}
		}

		private bool GetMergedCellInfo(ISheet sheet, List<int> activeRows, int minCol, int maxCol, int rowIdx, int colIdx, out int rowSpan, out int colSpan, out bool isTopLeft)
		{
			rowSpan = 1;
			colSpan = 1;
			isTopLeft = true;

			int numMerged = sheet.NumMergedRegions;
			for (int i = 0; i < numMerged; i++)
			{
				var region = sheet.GetMergedRegion(i);
				if (region.IsInRange(rowIdx, colIdx))
				{
					// Calculate rowspan based on active rows in range
					int countActive = 0;
					foreach (int ar in activeRows)
					{
						if (ar >= region.FirstRow && ar <= region.LastRow)
						{
							countActive++;
						}
					}
					rowSpan = countActive;

					int firstCol = Math.Max(region.FirstColumn, minCol);
					int lastCol = Math.Min(region.LastColumn, maxCol);
					colSpan = Math.Max(1, lastCol - firstCol + 1);

					isTopLeft = (region.FirstRow == rowIdx && firstCol == colIdx);
					return true;
				}
			}

			return false;
		}

		private string GetCellBackgroundColor(ICell cell)
		{
			if (cell == null || cell.CellStyle == null || cell.CellStyle.FillPattern == FillPattern.NoFill)
				return null;

			var style = cell.CellStyle;

			if (cell.Sheet.Workbook is HSSFWorkbook hssfWorkbook)
			{
				short colorIdx = style.FillForegroundColor;
				HSSFPalette palette = hssfWorkbook.GetCustomPalette();
				HSSFColor color = palette.GetColor(colorIdx);
				if (color != null && colorIdx != 64) // 64 is automatic/default
				{
					byte[] rgb = color.GetTriplet();
					return $"background-color: #{rgb[0]:X2}{rgb[1]:X2}{rgb[2]:X2};";
				}
			}
			else if (cell.Sheet.Workbook is XSSFWorkbook xssfWorkbook)
			{
				var xssfStyle = style as XSSFCellStyle;
				var color = xssfStyle?.FillForegroundColorColor as XSSFColor;
				if (color != null && !color.IsAuto)
				{
					byte[] rgb = color.RGB;
					if (rgb != null && rgb.Length >= 3)
					{
						if (rgb.Length == 4)
							return $"background-color: #{rgb[1]:X2}{rgb[2]:X2}{rgb[3]:X2};";
						return $"background-color: #{rgb[0]:X2}{rgb[1]:X2}{rgb[2]:X2};";
					}
				}
			}

			return null;
		}

		private string GetCellFontStyle(ICell cell)
		{
			if (cell == null || cell.CellStyle == null)
				return null;

			var style = cell.CellStyle;
			var font = cell.Sheet.Workbook.GetFontAt(style.FontIndex);
			if (font == null)
				return null;

			List<string> styles = new List<string>();
			if (font.IsBold) styles.Add("font-weight: bold;");
			if (font.IsItalic) styles.Add("font-style: italic;");
			if (font.Underline != FontUnderlineType.None) styles.Add("text-decoration: underline;");
			if (font.FontHeightInPoints > 0) styles.Add($"font-size: {font.FontHeightInPoints}pt;");
			if (!string.IsNullOrEmpty(font.FontName)) styles.Add($"font-family: '{font.FontName}', sans-serif;");

			string fontColor = GetFontColor(font, cell.Sheet.Workbook);
			if (!string.IsNullOrEmpty(fontColor)) styles.Add($"color: {fontColor};");

			return styles.Count > 0 ? string.Join(" ", styles) : null;
		}

		private string GetFontColor(IFont font, IWorkbook workbook)
		{
			if (font is HSSFFont hssfFont)
			{
				short colorIdx = hssfFont.Color;
				HSSFWorkbook hssfWorkbook = workbook as HSSFWorkbook;
				if (hssfWorkbook != null && colorIdx != 32767) // 32767 is normal/automatic
				{
					HSSFPalette palette = hssfWorkbook.GetCustomPalette();
					HSSFColor color = palette.GetColor(colorIdx);
					if (color != null)
					{
						byte[] rgb = color.GetTriplet();
						return $"#{rgb[0]:X2}{rgb[1]:X2}{rgb[2]:X2}";
					}
				}
			}
			else if (font is XSSFFont xssfFont)
			{
				XSSFColor color = xssfFont.GetXSSFColor();
				if (color != null && !color.IsAuto)
				{
					byte[] rgb = color.RGB;
					if (rgb != null && rgb.Length >= 3)
					{
						if (rgb.Length == 4)
							return $"#{rgb[1]:X2}{rgb[2]:X2}{rgb[3]:X2}";
						return $"#{rgb[0]:X2}{rgb[1]:X2}{rgb[2]:X2}";
					}
				}
			}
			return null;
		}

		private string GetCellAlignment(ICell cell)
		{
			if (cell == null || cell.CellStyle == null)
				return null;

			switch (cell.CellStyle.Alignment)
			{
				case HorizontalAlignment.Center:
					return "text-align: center;";
				case HorizontalAlignment.Right:
					return "text-align: right;";
				case HorizontalAlignment.Justify:
					return "text-align: justify;";
				case HorizontalAlignment.Left:
					return "text-align: left;";
				default:
					return null;
			}
		}

		private static byte[] CleanWorkbookStream(byte[] data)
		{
			using (var msOutput = new MemoryStream())
			{
				int length = data.Length;
				int offset = 0;
				bool skipChartSubstream = false;
				while (offset < length - 4)
				{
					ushort code = BitConverter.ToUInt16(data, offset);
					ushort len = BitConverter.ToUInt16(data, offset + 2);
					if (offset + 4 + len > length)
					{
						break;
					}

					bool skip = false;

					if (code == 0x0809) // BOF
					{
						if (len >= 6)
						{
							ushort type = BitConverter.ToUInt16(data, offset + 6);
							if (type == 0x0020) // Chart BOF
							{
								skipChartSubstream = true;
							}
						}
					}

					if (skipChartSubstream)
					{
						skip = true;
						if (code == 0x000A) // EOF of chart substream
						{
							skipChartSubstream = false;
						}
					}
					else
					{
						// We do not skip drawings or other records, only chart substreams
						skip = false;
					}

					if (!skip)
					{
						msOutput.Write(data, offset, 4 + len);
					}
					offset += 4 + len;
				}
				return msOutput.ToArray();
			}
		}

		private static void CopyPoifsEntry(NPOI.POIFS.FileSystem.POIFSFileSystem source, NPOI.POIFS.FileSystem.POIFSFileSystem target, string name)
		{
			try
			{
				if (source.Root.HasEntry(name))
				{
					using (var stream = source.CreateDocumentInputStream(name))
					{
						byte[] buffer = new byte[stream.Length];
						stream.Read(buffer, 0, buffer.Length);
						target.CreateDocument(new MemoryStream(buffer), name);
					}
				}
			}
			catch
			{
				// Ignore optional metadata stream errors
			}
		}
	}
}
