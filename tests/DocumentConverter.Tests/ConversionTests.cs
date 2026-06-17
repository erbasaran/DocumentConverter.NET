namespace DocumentConverter.Tests
{
	using System;
	using System.IO;
	using Xunit;

	public class NonClosingMemoryStream : MemoryStream
	{
		public override void Close()
		{
			// Do not close
		}

		protected override void Dispose(bool disposing)
		{
			// Do not dispose
		}

		public void ReallyDispose()
		{
			base.Dispose(true);
		}
	}

	public class ConversionTests
	{
		private string GetWorkspaceRoot()
		{
			string root = AppDomain.CurrentDomain.BaseDirectory;
			while (!string.IsNullOrEmpty(root))
			{
				if (File.Exists(Path.Combine(root, "DocumentConverter.slnx")))
				{
					return root;
				}
				root = Path.GetDirectoryName(root);
			}
			throw new FileNotFoundException("Workspace root with solution not found.");
		}

		private string GetTestFilePath(string fileName)
		{
			string root = GetWorkspaceRoot();
			string filesPath = Path.Combine(root, "tests", "DocumentConverter.Tests", "Files", fileName);
			if (File.Exists(filesPath))
			{
				return filesPath;
			}
			return Path.Combine(root, fileName);
		}

		[Fact]
		public void TestConvertDocToHtml()
		{
			string root = GetWorkspaceRoot();
			string docPath = GetTestFilePath("test.doc");
			string outputPath = Path.Combine(root, "test_doc_output.html");

			var service = new DocumentConverterService();
			var result = service.ConvertToHtml(docPath);

			Assert.True(result.IsSuccess, $"Failed to convert DOC file: {result.ErrorMessage}");
			Assert.NotNull(result.Value);
			Assert.NotEmpty(result.Value);

			File.WriteAllText(outputPath, result.Value);
		}

		[Fact]
		public void TestConvertDoc2ToHtml()
		{
			string root = GetWorkspaceRoot();
			string docPath = GetTestFilePath("test2.doc");
			string outputPath = Path.Combine(root, "test2_doc_output.html");

			if (File.Exists(docPath))
			{
				var service = new DocumentConverterService();
				var result = service.ConvertToHtml(docPath);

				Assert.True(result.IsSuccess, $"Failed to convert DOC2 file: {result.ErrorMessage}");
				Assert.NotNull(result.Value);
				Assert.NotEmpty(result.Value);

				File.WriteAllText(outputPath, result.Value);
			}
		}

		[Fact]
		public void TestConvertDocToHtmlWithHeaderFooter()
		{
			string root = GetWorkspaceRoot();
			string docPath = GetTestFilePath("test.doc");
			string outputPath = Path.Combine(root, "test_doc_header_footer_output.html");

			var service = new DocumentConverterService();
			var result = service.ConvertToHtml(docPath, includeHeaderFooter: true);

			Assert.True(result.IsSuccess, $"Failed to convert DOC file with header/footer: {result.ErrorMessage}");
			Assert.NotNull(result.Value);
			Assert.NotEmpty(result.Value);

			File.WriteAllText(outputPath, result.Value);
		}


		[Fact]
		public void TestConvertXlsToHtml()
		{
			string root = GetWorkspaceRoot();
			string xlsPath = GetTestFilePath("test.xls");
			string outputPath = Path.Combine(root, "test_xls_output.html");

			var service = new DocumentConverterService();
			var result = service.ConvertToHtml(xlsPath);

			Assert.True(result.IsSuccess, $"Failed to convert XLS file: {result.ErrorMessage}");
			Assert.NotNull(result.Value);
			Assert.NotEmpty(result.Value);

			File.WriteAllText(outputPath, result.Value);
		}

		[Fact]
		public void TestConvertDocxToHtml()
		{
			var doc = new NPOI.XWPF.UserModel.XWPFDocument();
			var p1 = doc.CreateParagraph();
			p1.Alignment = NPOI.XWPF.UserModel.ParagraphAlignment.CENTER;
			var r1 = p1.CreateRun();
			r1.SetText("Hello World from DOCX");
			r1.IsBold = true;
			r1.FontSize = 14;

			var table = doc.CreateTable(2, 2);
			table.GetRow(0).GetCell(0).SetText("Cell 1-1");
			table.GetRow(0).GetCell(1).SetText("Cell 1-2");
			table.GetRow(1).GetCell(0).SetText("Cell 2-1");
			table.GetRow(1).GetCell(1).SetText("Cell 2-2");

			var ms = new NonClosingMemoryStream();
			try
			{
				doc.Write(ms);
				ms.Position = 0;

				var service = new DocumentConverterService();
				var result = service.ConvertToHtml(ms, ".docx");

				Assert.True(result.IsSuccess, $"Failed to convert DOCX stream: {result.ErrorMessage}");
				Assert.NotNull(result.Value);
				Assert.Contains("Hello World from DOCX", result.Value);
				Assert.Contains("Cell 1-1", result.Value);
			}
			finally
			{
				ms.ReallyDispose();
			}
		}

		[Fact]
		public void TestConvertDocxToHtmlWithHeaderFooter()
		{
			var doc = new NPOI.XWPF.UserModel.XWPFDocument();

			// Create header/footer policy
			var policy = doc.CreateHeaderFooterPolicy();
			var header = policy.CreateHeader(NPOI.XWPF.Model.XWPFHeaderFooterPolicy.DEFAULT);
			header.CreateParagraph().CreateRun().SetText("Header Content");

			// Create footer
			var footer = policy.CreateFooter(NPOI.XWPF.Model.XWPFHeaderFooterPolicy.DEFAULT);
			footer.CreateParagraph().CreateRun().SetText("Footer Content");

			var p1 = doc.CreateParagraph();
			p1.CreateRun().SetText("Body Content");

			var ms = new NonClosingMemoryStream();
			try
			{
				doc.Write(ms);
				ms.Position = 0;

				var service = new DocumentConverterService();

				// Test without header/footer (default)
				var resultDefault = service.ConvertToHtml(ms, ".docx");
				Assert.True(resultDefault.IsSuccess);
				Assert.Contains("Body Content", resultDefault.Value);
				Assert.DoesNotContain("Header Content", resultDefault.Value);
				Assert.DoesNotContain("Footer Content", resultDefault.Value);

				// Test with header/footer
				ms.Position = 0;
				var resultWithHeaderFooter = service.ConvertToHtml(ms, ".docx", includeHeaderFooter: true);
				Assert.True(resultWithHeaderFooter.IsSuccess);
				Assert.Contains("Body Content", resultWithHeaderFooter.Value);
				Assert.Contains("Header Content", resultWithHeaderFooter.Value);
				Assert.Contains("Footer Content", resultWithHeaderFooter.Value);
			}
			finally
			{
				ms.ReallyDispose();
			}
		}

		[Fact]
		public void TestConvertXlsxToHtml()
		{
			var workbook = new NPOI.XSSF.UserModel.XSSFWorkbook();
			var sheet = workbook.CreateSheet("TestSheet");
			var row1 = sheet.CreateRow(0);
			row1.CreateCell(0).SetCellValue("Header 1");
			row1.CreateCell(1).SetCellValue("Header 2");

			var row2 = sheet.CreateRow(1);
			row2.CreateCell(0).SetCellValue("Data 1");
			row2.CreateCell(1).SetCellValue("Data 2");

			// Merged cell
			sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(2, 2, 0, 1));
			var row3 = sheet.CreateRow(2);
			row3.CreateCell(0).SetCellValue("Merged Cell");

			var ms = new NonClosingMemoryStream();
			try
			{
				workbook.Write(ms);
				ms.Position = 0;

				var service = new DocumentConverterService();
				var result = service.ConvertToHtml(ms, ".xlsx");

				Assert.True(result.IsSuccess, $"Failed to convert XLSX stream: {result.ErrorMessage}");
				Assert.NotNull(result.Value);
				Assert.Contains("Header 1", result.Value);
				Assert.Contains("Data 2", result.Value);
				Assert.Contains("Merged Cell", result.Value);
			}
			finally
			{
				ms.ReallyDispose();
			}
		}

		[Fact]
		public void TestXlsPictures()
		{
			string root = GetWorkspaceRoot();
			string xlsPath = GetTestFilePath("test.xls");

			var service = new DocumentConverterService();
			var result = service.ConvertToHtml(xlsPath);

			Assert.True(result.IsSuccess, $"Failed to convert XLS file: {result.ErrorMessage}");
			Assert.NotNull(result.Value);
			Assert.DoesNotContain("<img src=\"data:", result.Value);

			File.WriteAllText(Path.Combine(root, "test_xls_output.html"), result.Value);
		}

		[Fact]
		public void TestConvertPdfToHtml()
		{
			string root = GetWorkspaceRoot();
			string pdfPath = GetTestFilePath("test.pdf");
			string outputPath = Path.Combine(root, "test_pdf_output.html");

			var service = new DocumentConverterService();
			var result = service.ConvertToHtml(pdfPath);

			Assert.True(result.IsSuccess, $"Failed to convert PDF file: {result.ErrorMessage}");
			Assert.NotNull(result.Value);
			Assert.NotEmpty(result.Value);

			File.WriteAllText(outputPath, result.Value);
		}

		[Fact]
		public void TestMoveAndConvertFiles()
		{
			string root = GetWorkspaceRoot();
			string filesDir = Path.Combine(root, "tests", "DocumentConverter.Tests", "Files");
			string outputDir = Path.Combine(root, "output");

			Directory.CreateDirectory(filesDir);
			Directory.CreateDirectory(outputDir);

			// Move documents from root to Files directory
			string[] extensions = { ".doc", ".docx", ".xls", ".xlsx", ".pdf" };
			foreach (var file in Directory.GetFiles(root))
			{
				string ext = Path.GetExtension(file).ToLower();
				if (extensions.Contains(ext))
				{
					string dest = Path.Combine(filesDir, Path.GetFileName(file));
					if (File.Exists(dest)) File.Delete(dest);
					File.Move(file, dest);
				}
			}

			// Convert all files in Files directory to output folder
			var service = new DocumentConverterService();
			foreach (var file in Directory.GetFiles(filesDir))
			{
				string ext = Path.GetExtension(file).ToLower();
				if (extensions.Contains(ext))
				{
					var result = service.ConvertToHtml(file);
					if (result.IsSuccess)
					{
						string extName = ext.TrimStart('.');
						string outputFileName = $"{Path.GetFileNameWithoutExtension(file)}_{extName}.html";
						string outputPath = Path.Combine(outputDir, outputFileName);
						File.WriteAllText(outputPath, result.Value);
					}
				}
			}
		}

		[Fact]
		public void TestHtmlWrapperOption()
		{
			var doc = new NPOI.XWPF.UserModel.XWPFDocument();
			var p1 = doc.CreateParagraph();
			var r1 = p1.CreateRun();
			r1.SetText("Test paragraph content");

			var ms = new NonClosingMemoryStream();
			try
			{
				doc.Write(ms);
				ms.Position = 0;

				var service = new DocumentConverterService();

				// 1. Test default behavior (should be false, no wrapper)
				var resultDefault = service.ConvertToHtml(ms, ".docx");
				Assert.True(resultDefault.IsSuccess);
				Assert.Contains("Test paragraph content", resultDefault.Value);
				Assert.DoesNotContain("<!DOCTYPE html>", resultDefault.Value);
				Assert.DoesNotContain("<html>", resultDefault.Value);
				Assert.DoesNotContain("</html>", resultDefault.Value);

				// 2. Test explicit true (should include wrapper)
				ms.Position = 0;
				var resultWrapped = service.ConvertToHtml(ms, ".docx", includeHtmlWrapper: true);
				Assert.True(resultWrapped.IsSuccess);
				Assert.Contains("Test paragraph content", resultWrapped.Value);
				Assert.Contains("<!DOCTYPE html>", resultWrapped.Value);
				Assert.Contains("<html>", resultWrapped.Value);
				Assert.Contains("</html>", resultWrapped.Value);
			}
			finally
			{
				ms.ReallyDispose();
			}
		}
	}
}
