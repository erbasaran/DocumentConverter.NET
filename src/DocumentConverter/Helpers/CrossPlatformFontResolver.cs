namespace DocumentConverter.Helpers
{
	using System;
	using System.Collections.Generic;
	using System.IO;
	using PdfSharp.Fonts;

	/// <summary>
	/// Cross-platform font resolver for PDFsharp that resolves fonts from standard system directories.
	/// </summary>
	internal class CrossPlatformFontResolver : IFontResolver
	{
		private static readonly object SyncLock = new object();
		private static bool _isRegistered = false;
		private static readonly Dictionary<string, byte[]> FontDataCache = new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase);
		private static readonly List<string> FontDirectories = new List<string>();

		static CrossPlatformFontResolver()
		{
			// Windows fonts directory
			try
			{
				string winFonts = Environment.GetFolderPath(Environment.SpecialFolder.Fonts);
				if (!string.IsNullOrEmpty(winFonts) && Directory.Exists(winFonts))
				{
					FontDirectories.Add(winFonts);
				}
			}
			catch { }

			// Standard Linux font directories
			string[] linuxDirs = {
				"/usr/share/fonts",
				"/usr/local/share/fonts",
				"/usr/share/fonts/truetype",
				"/usr/share/fonts/truetype/dejavu",
				"/usr/share/fonts/truetype/liberation",
				"/usr/share/fonts/truetype/msttcorefonts"
			};
			foreach (var dir in linuxDirs)
			{
				try
				{
					if (Directory.Exists(dir))
					{
						FontDirectories.Add(dir);
					}
				}
				catch { }
			}

			// macOS font directories
			string[] macDirs = {
				"/System/Library/Fonts",
				"/Library/Fonts"
			};
			foreach (var dir in macDirs)
			{
				try
				{
					if (Directory.Exists(dir))
					{
						FontDirectories.Add(dir);
					}
				}
				catch { }
			}
		}

		/// <summary>
		/// Ensures this font resolver is registered with PDFsharp once.
		/// </summary>
		public static void EnsureRegistered()
		{
			if (!_isRegistered)
			{
				lock (SyncLock)
				{
					if (!_isRegistered)
					{
						try
						{
							if (GlobalFontSettings.FontResolver == null)
							{
								GlobalFontSettings.FontResolver = new CrossPlatformFontResolver();
							}
						}
						catch
						{
							// Ignore if already registered by host application
						}
						_isRegistered = true;
					}
				}
			}
		}

		public FontResolverInfo ResolveTypeface(string familyName, bool isBold, bool isItalic)
		{
			string name = familyName?.Trim() ?? "Arial";

			// Try to find the exact font or a sensible fallback
			string faceName = GetFaceName(name, isBold, isItalic);
			if (faceName != null)
			{
				return new FontResolverInfo(faceName);
			}

			// Fallback to Arial if not already tried
			if (!name.Equals("Arial", StringComparison.OrdinalIgnoreCase))
			{
				faceName = GetFaceName("Arial", isBold, isItalic);
				if (faceName != null)
				{
					return new FontResolverInfo(faceName);
				}
			}

			// Generic fallback: find any valid TTF font in font directories
			faceName = GetAnyAvailableFontFace();
			if (faceName != null)
			{
				return new FontResolverInfo(faceName);
			}

			return null;
		}

		public byte[] GetFont(string faceName)
		{
			if (string.IsNullOrEmpty(faceName)) return null;

			lock (FontDataCache)
			{
				if (FontDataCache.TryGetValue(faceName, out var cached))
				{
					return cached;
				}

				string filePath = FindFontFile(faceName);
				if (filePath != null && File.Exists(filePath))
				{
					try
					{
						byte[] data = File.ReadAllBytes(filePath);
						FontDataCache[faceName] = data;
						return data;
					}
					catch { }
				}
			}

			return null;
		}

		private string GetFaceName(string familyName, bool isBold, bool isItalic)
		{
			string prefix = familyName.ToLowerInvariant();
			string suffix;

			if (isBold && isItalic) suffix = "bi";
			else if (isBold) suffix = "bd";
			else if (isItalic) suffix = "i";
			else suffix = "";

			// Check standard candidates (e.g. arial.ttf, arialbd.ttf, etc.)
			List<string> candidates = new List<string>();

			if (prefix == "arial")
			{
				if (suffix == "bi") candidates.AddRange(new[] { "arialbi.ttf", "LiberationSans-BoldItalic.ttf", "DejaVuSans-BoldOblique.ttf" });
				else if (suffix == "bd") candidates.AddRange(new[] { "arialbd.ttf", "LiberationSans-Bold.ttf", "DejaVuSans-Bold.ttf" });
				else if (suffix == "i") candidates.AddRange(new[] { "ariali.ttf", "LiberationSans-Italic.ttf", "DejaVuSans-Oblique.ttf" });
				else candidates.AddRange(new[] { "arial.ttf", "LiberationSans-Regular.ttf", "DejaVuSans.ttf" });
			}
			else
			{
				string candidateBase = $"{prefix}{suffix}.ttf";
				candidates.Add(candidateBase);
				candidates.Add($"{prefix}.ttf");
			}

			foreach (var candidate in candidates)
			{
				if (FindFontFile(candidate) != null)
				{
					return candidate;
				}
			}

			return null;
		}

		private string GetAnyAvailableFontFace()
		{
			foreach (var dir in FontDirectories)
			{
				try
				{
					if (Directory.Exists(dir))
					{
						var files = Directory.GetFiles(dir, "*.ttf", SearchOption.TopDirectoryOnly);
						if (files.Length > 0)
						{
							return Path.GetFileName(files[0]);
						}
					}
				}
				catch { }
			}
			return null;
		}

		private string FindFontFile(string fileName)
		{
			foreach (var dir in FontDirectories)
			{
				try
				{
					string path = Path.Combine(dir, fileName);
					if (File.Exists(path))
					{
						return path;
					}

					// Also check case-insensitive match on Linux
					var files = Directory.GetFiles(dir, fileName, SearchOption.AllDirectories);
					if (files.Length > 0)
					{
						return files[0];
					}
				}
				catch { }
			}

			return null;
		}
	}
}
