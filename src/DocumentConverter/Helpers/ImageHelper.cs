namespace DocumentConverter.Helpers
{
	using System;
	using System.Collections.Generic;
	using System.IO;
	using System.Net;
	using System.Net.Http;
	using System.Runtime.InteropServices;
	using System.Text;

	public static class ImageHelper
	{
		public static int MaxImageSizeBytes { get; set; } = 25 * 1024 * 1024; // 25 MB
		public static bool AllowRemoteImages { get; set; } = true;
		public static bool AllowLocalFiles { get; set; } = true;
		public static string AllowedLocalDirectory { get; set; } = null;

		private static readonly HashSet<string> AllowedExtensions = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
		{
			".jpg", ".jpeg", ".png", ".gif", ".bmp", ".tiff", ".tif", ".svg", ".emf", ".wmf", ".webp"
		};

		private static readonly Dictionary<string, string> MimeTypes = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
		{
			{ "jpg", "image/jpeg" },
			{ "jpeg", "image/jpeg" },
			{ "png", "image/png" },
			{ "gif", "image/gif" },
			{ "bmp", "image/bmp" },
			{ "tiff", "image/tiff" },
			{ "tif", "image/tiff" },
			{ "svg", "image/svg+xml" },
			{ "emf", "image/x-emf" },
			{ "wmf", "image/x-wmf" },
			{ "webp", "image/webp" }
		};

		public static byte[] GetImageBytes(string src)
		{
			if (string.IsNullOrEmpty(src)) return null;

			try
			{
				// 1. Data URI scheme
				if (src.StartsWith("data:", StringComparison.OrdinalIgnoreCase))
				{
					return ParseDataUri(src);
				}

				// 2. HTTP/HTTPS Remote URLs
				if (src.StartsWith("http://", StringComparison.OrdinalIgnoreCase) || src.StartsWith("https://", StringComparison.OrdinalIgnoreCase))
				{
					if (!AllowRemoteImages) return null;
					return DownloadRemoteImageSafely(src);
				}

				// 3. Local Filesystem
				if (AllowLocalFiles)
				{
					return LoadLocalImageSafely(src);
				}
			}
			catch
			{
				// Ignore load/download failures and return null
			}

			return null;
		}

		private static byte[] ParseDataUri(string src)
		{
			int commaIdx = src.IndexOf(',');
			if (commaIdx <= 0 || !src.StartsWith("data:image/", StringComparison.OrdinalIgnoreCase))
			{
				return null;
			}

			string header = src.Substring(0, commaIdx);
			if (header.IndexOf(";base64", StringComparison.OrdinalIgnoreCase) < 0)
			{
				return null;
			}

			string base64Data = src.Substring(commaIdx + 1).Trim();
			if (base64Data.Length > (MaxImageSizeBytes * 4 / 3) + 1024)
			{
				return null;
			}

			byte[] bytes = Convert.FromBase64String(base64Data);
			if (bytes != null && bytes.Length <= MaxImageSizeBytes && IsValidImage(bytes))
			{
				return bytes;
			}

			return null;
		}

		private static byte[] DownloadRemoteImageSafely(string url)
		{
			if (!Uri.TryCreate(url, UriKind.Absolute, out Uri targetUri))
			{
				return null;
			}

			int maxRedirects = 3;
			Uri currentUri = targetUri;

			for (int redirectCount = 0; redirectCount <= maxRedirects; redirectCount++)
			{
				if (!IsSafeRemoteUri(currentUri))
				{
					return null;
				}

				using (var handler = new HttpClientHandler { AllowAutoRedirect = false })
				using (var client = new HttpClient(handler) { Timeout = TimeSpan.FromSeconds(8) })
				using (var request = new HttpRequestMessage(HttpMethod.Get, currentUri))
				{
					request.Headers.UserAgent.ParseAdd("DocumentConverter.NET/1.0");

					using (var response = client.SendAsync(request, HttpCompletionOption.ResponseHeadersRead).GetAwaiter().GetResult())
					{
						int statusCode = (int)response.StatusCode;

						// Handle safe redirects manually to verify target IP at every hop
						if (statusCode >= 300 && statusCode <= 399)
						{
							var location = response.Headers.Location;
							if (location == null) return null;

							if (!location.IsAbsoluteUri)
							{
								location = new Uri(currentUri, location);
							}
							currentUri = location;
							continue;
						}

						if (!response.IsSuccessStatusCode)
						{
							return null;
						}

						if (response.Content.Headers.ContentLength.HasValue &&
							response.Content.Headers.ContentLength.Value > MaxImageSizeBytes)
						{
							return null;
						}

						using (var stream = response.Content.ReadAsStreamAsync().GetAwaiter().GetResult())
						using (var ms = new MemoryStream())
						{
							byte[] buffer = new byte[8192];
							int bytesRead;
							int totalBytes = 0;

							while ((bytesRead = stream.Read(buffer, 0, buffer.Length)) > 0)
							{
								totalBytes += bytesRead;
								if (totalBytes > MaxImageSizeBytes)
								{
									return null; // Exceeded maximum size, abort
								}
								ms.Write(buffer, 0, bytesRead);
							}

							byte[] downloadedBytes = ms.ToArray();
							if (IsValidImage(downloadedBytes))
							{
								return downloadedBytes;
							}
							return null;
						}
					}
				}
			}

			return null;
		}

		private static byte[] LoadLocalImageSafely(string src)
		{
			// Check for illegal or dangerous paths
			if (src.IndexOfAny(Path.GetInvalidPathChars()) >= 0)
			{
				return null;
			}

			string extension = Path.GetExtension(src);
			if (string.IsNullOrEmpty(extension) || !AllowedExtensions.Contains(extension))
			{
				return null;
			}

			string fullPath;
			try
			{
				fullPath = Path.GetFullPath(src);
			}
			catch
			{
				return null;
			}

			// If an allowed directory is enforced, verify path traversal containment
			if (!string.IsNullOrEmpty(AllowedLocalDirectory))
			{
				string allowedFull = Path.GetFullPath(AllowedLocalDirectory);
				if (!fullPath.StartsWith(allowedFull, StringComparison.OrdinalIgnoreCase))
				{
					return null;
				}
			}

			if (!File.Exists(fullPath))
			{
				return null;
			}

			var fileInfo = new FileInfo(fullPath);
			if (fileInfo.Length > MaxImageSizeBytes)
			{
				return null;
			}

			byte[] bytes = File.ReadAllBytes(fullPath);
			if (IsValidImage(bytes, extension))
			{
				return bytes;
			}

			return null;
		}

		/// <summary>
		/// Verifies that the given URI does not point to internal, private, loopback, or metadata addresses (SSRF mitigation).
		/// </summary>
		public static bool IsSafeRemoteUri(Uri uri)
		{
			if (uri == null) return false;
			if (uri.Scheme != Uri.UriSchemeHttp && uri.Scheme != Uri.UriSchemeHttps)
			{
				return false;
			}

			string host = uri.DnsSafeHost;
			if (string.IsNullOrEmpty(host)) return false;

			IPAddress[] addresses;
			try
			{
				addresses = Dns.GetHostAddresses(host);
			}
			catch
			{
				return false;
			}

			if (addresses == null || addresses.Length == 0) return false;

			foreach (var address in addresses)
			{
				if (IsPrivateOrRestrictedIp(address))
				{
					return false;
				}
			}

			return true;
		}

		private static bool IsPrivateOrRestrictedIp(IPAddress ip)
		{
			if (IPAddress.IsLoopback(ip)) return true;

			IPAddress checkIp = ip;
			if (checkIp.IsIPv4MappedToIPv6)
			{
				checkIp = checkIp.MapToIPv4();
			}

			byte[] bytes = checkIp.GetAddressBytes();

			if (checkIp.AddressFamily == System.Net.Sockets.AddressFamily.InterNetwork)
			{
				// 0.0.0.0/8 (Current network)
				if (bytes[0] == 0) return true;

				// 10.0.0.0/8 (Private network)
				if (bytes[0] == 10) return true;

				// 100.64.0.0/10 (Shared address / CGNAT)
				if (bytes[0] == 100 && (bytes[1] >= 64 && bytes[1] <= 127)) return true;

				// 127.0.0.0/8 (Loopback)
				if (bytes[0] == 127) return true;

				// 169.254.0.0/16 (Link-local / Cloud metadata: AWS, Azure, GCP)
				if (bytes[0] == 169 && bytes[1] == 254) return true;

				// 172.16.0.0/12 (Private network)
				if (bytes[0] == 172 && (bytes[1] >= 16 && bytes[1] <= 31)) return true;

				// 192.0.0.0/24 (IETF Protocol Assignments)
				if (bytes[0] == 192 && bytes[1] == 0 && bytes[2] == 0) return true;

				// 192.0.2.0/24 (TEST-NET-1)
				if (bytes[0] == 192 && bytes[1] == 0 && bytes[2] == 2) return true;

				// 192.168.0.0/16 (Private network)
				if (bytes[0] == 192 && bytes[1] == 168) return true;

				// 198.18.0.0/15 (Benchmarking)
				if (bytes[0] == 198 && (bytes[1] == 18 || bytes[1] == 19)) return true;

				// 198.51.100.0/24 (TEST-NET-2)
				if (bytes[0] == 198 && bytes[1] == 51 && bytes[2] == 100) return true;

				// 203.0.113.0/24 (TEST-NET-3)
				if (bytes[0] == 203 && bytes[1] == 0 && bytes[2] == 113) return true;

				// 224.0.0.0/4 (Multicast) & 240.0.0.0/4 (Reserved)
				if (bytes[0] >= 224) return true;

				// 255.255.255.255 (Broadcast)
				if (bytes[0] == 255 && bytes[1] == 255 && bytes[2] == 255 && bytes[3] == 255) return true;
			}
			else if (checkIp.AddressFamily == System.Net.Sockets.AddressFamily.InterNetworkV6)
			{
				if (checkIp.Equals(IPAddress.IPv6Loopback) || checkIp.Equals(IPAddress.IPv6None) || checkIp.Equals(IPAddress.IPv6Any))
				{
					return true;
				}

				if (checkIp.IsIPv6LinkLocal || checkIp.IsIPv6SiteLocal || checkIp.IsIPv6Multicast)
				{
					return true;
				}

				// Unique Local Addresses (fc00::/7)
				if ((bytes[0] & 0xFE) == 0xFC)
				{
					return true;
				}
			}

			return false;
		}

		/// <summary>
		/// Validates whether the byte array matches genuine image format headers.
		/// </summary>
		public static bool IsValidImage(byte[] bytes, string expectedExtension = null)
		{
			if (bytes == null || bytes.Length < 4) return false;

			// 1. PNG: 89 50 4E 47 0D 0A 1A 0A
			if (bytes.Length >= 8 && bytes[0] == 0x89 && bytes[1] == 0x50 && bytes[2] == 0x4E && bytes[3] == 0x47 &&
				bytes[4] == 0x0D && bytes[5] == 0x0A && bytes[6] == 0x1A && bytes[7] == 0x0A)
			{
				return true;
			}

			// 2. JPEG: FF D8 FF
			if (bytes.Length >= 3 && bytes[0] == 0xFF && bytes[1] == 0xD8 && bytes[2] == 0xFF)
			{
				return true;
			}

			// 3. GIF: GIF87a or GIF89a
			if (bytes.Length >= 6 && bytes[0] == 0x47 && bytes[1] == 0x49 && bytes[2] == 0x46 && bytes[3] == 0x38 &&
				(bytes[4] == 0x37 || bytes[4] == 0x39) && bytes[5] == 0x61)
			{
				return true;
			}

			// 4. BMP: BM
			if (bytes.Length >= 2 && bytes[0] == 0x42 && bytes[1] == 0x4D)
			{
				return true;
			}

			// 5. TIFF: II*\0 or MM\0*
			if (bytes.Length >= 4 &&
				((bytes[0] == 0x49 && bytes[1] == 0x49 && bytes[2] == 0x2A && bytes[3] == 0x00) ||
				 (bytes[0] == 0x4D && bytes[1] == 0x4D && bytes[2] == 0x00 && bytes[3] == 0x2A)))
			{
				return true;
			}

			// 6. WebP: RIFF....WEBP
			if (bytes.Length >= 12 &&
				bytes[0] == 0x52 && bytes[1] == 0x49 && bytes[2] == 0x46 && bytes[3] == 0x46 &&
				bytes[8] == 0x57 && bytes[9] == 0x45 && bytes[10] == 0x42 && bytes[11] == 0x50)
			{
				return true;
			}

			// 7. EMF / WMF
			if (bytes.Length >= 44 && BitConverter.ToUInt32(bytes, 0) == 1 &&
				bytes[40] == 0x20 && bytes[41] == 0x45 && bytes[42] == 0x4D && bytes[43] == 0x46)
			{
				return true;
			}

			if (bytes.Length >= 4 &&
				((bytes[0] == 0xD7 && bytes[1] == 0xCD && bytes[2] == 0xC6 && bytes[3] == 0x9A) ||
				 (bytes[0] == 0x01 && bytes[1] == 0x00 && bytes[2] == 0x09 && bytes[3] == 0x00)))
			{
				return true;
			}

			// 8. SVG check (text/xml)
			if (bytes.Length >= 10)
			{
				int checkLen = Math.Min(bytes.Length, 512);
				string prefix = Encoding.UTF8.GetString(bytes, 0, checkLen).TrimStart();
				if (prefix.IndexOf("<svg", StringComparison.OrdinalIgnoreCase) >= 0 ||
					(prefix.StartsWith("<?xml", StringComparison.OrdinalIgnoreCase) && prefix.IndexOf("<svg", StringComparison.OrdinalIgnoreCase) >= 0))
				{
					return true;
				}
			}

			return false;
		}

		public static bool TryGetImageDimensions(byte[] bytes, out int width, out int height)
		{
			width = 0;
			height = 0;
			if (bytes == null || bytes.Length < 8) return false;

			try
			{
				// 1. PNG check: starts with 89 50 4E 47 0D 0A 1A 0A
				if (bytes.Length >= 24 && bytes[0] == 0x89 && bytes[1] == 0x50 && bytes[2] == 0x4E && bytes[3] == 0x47)
				{
					width = (bytes[16] << 24) | (bytes[17] << 16) | (bytes[18] << 8) | bytes[19];
					height = (bytes[20] << 24) | (bytes[21] << 16) | (bytes[22] << 8) | bytes[23];
					return width > 0 && height > 0;
				}

				// 2. GIF check: starts with 'GIF87a' or 'GIF89a'
				if (bytes.Length >= 10 && bytes[0] == 0x47 && bytes[1] == 0x49 && bytes[2] == 0x46 && bytes[3] == 0x38)
				{
					width = bytes[6] | (bytes[7] << 8);
					height = bytes[8] | (bytes[9] << 8);
					return width > 0 && height > 0;
				}

				// 3. BMP check: starts with 'BM'
				if (bytes.Length >= 26 && bytes[0] == 0x42 && bytes[1] == 0x4D)
				{
					width = bytes[18] | (bytes[19] << 8) | (bytes[20] << 16) | (bytes[21] << 24);
					height = bytes[22] | (bytes[23] << 8) | (bytes[24] << 16) | (bytes[25] << 24);
					height = Math.Abs(height);
					return width > 0 && height > 0;
				}

				// 4. JPEG check: starts with FF D8
				if (bytes.Length >= 4 && bytes[0] == 0xFF && bytes[1] == 0xD8)
				{
					int offset = 2;
					while (offset < bytes.Length - 8)
					{
						byte markerStart = bytes[offset];
						byte markerType = bytes[offset + 1];
						if (markerStart != 0xFF) break;

						if (markerType == 0xFF)
						{
							offset++;
							continue;
						}

						int segmentLength = (bytes[offset + 2] << 8) | bytes[offset + 3];

						if ((markerType >= 0xC0 && markerType <= 0xCF) && 
							markerType != 0xC4 && markerType != 0xC8 && markerType != 0xCC)
						{
							if (offset + 8 < bytes.Length)
							{
								height = (bytes[offset + 5] << 8) | bytes[offset + 6];
								width = (bytes[offset + 7] << 8) | bytes[offset + 8];
								return width > 0 && height > 0;
							}
						}

						offset += 2 + segmentLength;
					}
				}
			}
			catch { }

			return false;
		}

		public static byte[] ProcessImage(byte[] rawData, string extension, out string mimeType)
		{
			mimeType = GetMimeTypeFromExtension(extension);
			if (rawData == null || rawData.Length == 0) return rawData;

			string ext = extension?.TrimStart('.').ToLowerInvariant() ?? "";

			if (ext == "emf" || ext == "wmf" || ext == "x-emf" || ext == "x-wmf")
			{
				byte[] pngBytes = ConvertMetafileToPngIfWindows(rawData, out string convertedMime);
				if (convertedMime == "image/png")
				{
					mimeType = "image/png";
					return pngBytes;
				}

				if (ext == "emf" || ext == "x-emf")
				{
					byte[] rasterBytes = ExtractRasterFromEmf(rawData, out string rasterMime);
					if (rasterMime != "image/x-emf")
					{
						mimeType = rasterMime;
						return rasterBytes;
					}
				}
			}

			return rawData;
		}

		public static string GetMimeTypeFromExtension(string extension)
		{
			if (string.IsNullOrEmpty(extension)) return "image/png";
			extension = extension.TrimStart('.').ToLowerInvariant();
			if (MimeTypes.TryGetValue(extension, out var mimeType))
			{
				return mimeType;
			}
			return $"image/{extension}";
		}

		private static byte[] ConvertMetafileToPngIfWindows(byte[] bytes, out string mimeType)
		{
			mimeType = "image/x-metafile";

			if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows) || bytes == null || bytes.Length == 0)
			{
				return bytes;
			}

			try
			{
				System.Reflection.Assembly drawingAssembly = null;
				try
				{
					drawingAssembly = System.Reflection.Assembly.Load("System.Drawing, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a");
				}
				catch
				{
					try
					{
						drawingAssembly = System.Reflection.Assembly.Load("System.Drawing.Common");
					}
					catch { }
				}

				if (drawingAssembly != null)
				{
					var imageType = drawingAssembly.GetType("System.Drawing.Image");
					var imageFormatType = drawingAssembly.GetType("System.Drawing.Imaging.ImageFormat");

					if (imageType != null && imageFormatType != null)
					{
						var fromStreamMethod = imageType.GetMethod("FromStream", new[] { typeof(Stream) });
						var saveMethod = imageType.GetMethod("Save", new[] { typeof(Stream), imageFormatType });
						var pngProperty = imageFormatType.GetProperty("Png");

						if (fromStreamMethod != null && saveMethod != null && pngProperty != null)
						{
							using (var msInput = new MemoryStream(bytes))
							{
								var imgInstance = fromStreamMethod.Invoke(null, new object[] { msInput });
								if (imgInstance != null)
								{
									try
									{
										var bitmapType = drawingAssembly.GetType("System.Drawing.Bitmap");
										var graphicsType = drawingAssembly.GetType("System.Drawing.Graphics");
										var colorType = drawingAssembly.GetType("System.Drawing.Color");
										var rectType = drawingAssembly.GetType("System.Drawing.Rectangle");
										var graphicsUnitType = drawingAssembly.GetType("System.Drawing.GraphicsUnit");

										if (bitmapType != null && graphicsType != null && colorType != null && rectType != null && graphicsUnitType != null)
										{
											var fromImageMethod = graphicsType.GetMethod("FromImage", new[] { imageType });
											var clearMethod = graphicsType.GetMethod("Clear", new[] { colorType });
											var drawImageMethod = graphicsType.GetMethod("DrawImage", new[] { imageType, typeof(int), typeof(int), typeof(int), typeof(int) });
											var drawImageRectMethod = graphicsType.GetMethod("DrawImage", new[] { imageType, rectType, rectType, graphicsUnitType });
											var getPixelMethod = bitmapType.GetMethod("GetPixel", new[] { typeof(int), typeof(int) });

											var aProperty = colorType.GetProperty("A");
											var rProperty = colorType.GetProperty("R");
											var gProperty = colorType.GetProperty("G");
											var bProperty = colorType.GetProperty("B");

											var transparentColor = colorType.GetProperty("Transparent").GetValue(null);
											var pngFormat = pngProperty.GetValue(null);

											var widthProp = imageType.GetProperty("Width");
											var heightProp = imageType.GetProperty("Height");
											int width = (int)widthProp.GetValue(imgInstance);
											int height = (int)heightProp.GetValue(imgInstance);

											// Dimension sanity check to prevent decompression bomb OutOfMemoryException
											if (width > 0 && height > 0 && width <= 8192 && height <= 8192 &&
												fromImageMethod != null && clearMethod != null && drawImageMethod != null && getPixelMethod != null)
											{
												var bmp = Activator.CreateInstance(bitmapType, new object[] { width, height });
												using (var g = fromImageMethod.Invoke(null, new object[] { bmp }) as IDisposable)
												{
													if (g != null)
													{
														clearMethod.Invoke(g, new object[] { transparentColor });
														drawImageMethod.Invoke(g, new object[] { imgInstance, 0, 0, width, height });
													}
												}

												Func<int, int, bool> isInk = (x, y) =>
												{
													var color = getPixelMethod.Invoke(bmp, new object[] { x, y });
													byte a = (byte)aProperty.GetValue(color);
													byte r = (byte)rProperty.GetValue(color);
													byte g = (byte)gProperty.GetValue(color);
													byte b = (byte)bProperty.GetValue(color);
													return a >= 10 && !(r > 250 && g > 250 && b > 250);
												};

												int minY = -1;
												for (int y = 0; y < height; y++)
												{
													for (int x = 0; x < width; x++)
													{
														if (isInk(x, y)) { minY = y; break; }
													}
													if (minY != -1) break;
												}

												int maxY = -1;
												if (minY != -1)
												{
													for (int y = height - 1; y >= 0; y--)
													{
														for (int x = 0; x < width; x++)
														{
															if (isInk(x, y)) { maxY = y; break; }
														}
														if (maxY != -1) break;
													}
												}

												int minX = -1;
												if (minY != -1 && maxY != -1)
												{
													for (int x = 0; x < width; x++)
													{
														for (int y = minY; y <= maxY; y++)
														{
															if (isInk(x, y)) { minX = x; break; }
														}
														if (minX != -1) break;
													}
												}

												int maxX = -1;
												if (minY != -1 && maxY != -1 && minX != -1)
												{
													for (int x = width - 1; x >= 0; x--)
													{
														for (int y = minY; y <= maxY; y++)
														{
															if (isInk(x, y)) { maxX = x; break; }
														}
														if (maxX != -1) break;
													}
												}

												object finalBmp = bmp;
												bool cropped = false;

												if (minY != -1 && maxY != -1 && minX != -1 && maxX != -1)
												{
													int pad = 4;
													minX = Math.Max(0, minX - pad);
													minY = Math.Max(0, minY - pad);
													maxX = Math.Min(width - 1, maxX + pad);
													maxY = Math.Min(height - 1, maxY + pad);

													int cropWidth = maxX - minX + 1;
													int cropHeight = maxY - minY + 1;

													if ((cropWidth < width || cropHeight < height) && cropWidth > 0 && cropHeight > 0 && drawImageRectMethod != null)
													{
														var croppedBmp = Activator.CreateInstance(bitmapType, new object[] { cropWidth, cropHeight });
														using (var gCrop = fromImageMethod.Invoke(null, new object[] { croppedBmp }) as IDisposable)
														{
															if (gCrop != null)
															{
																clearMethod.Invoke(gCrop, new object[] { transparentColor });

																var pixelUnit = Enum.Parse(graphicsUnitType, "Pixel");
																var destRect = Activator.CreateInstance(rectType, new object[] { 0, 0, cropWidth, cropHeight });
																var srcRect = Activator.CreateInstance(rectType, new object[] { minX, minY, cropWidth, cropHeight });

																drawImageRectMethod.Invoke(gCrop, new object[] { bmp, destRect, srcRect, pixelUnit });
															}
														}
														finalBmp = croppedBmp;
														cropped = true;
													}
												}

												using (var msOutput = new MemoryStream())
												{
													saveMethod.Invoke(finalBmp, new object[] { msOutput, pngFormat });

													if (cropped)
													{
														((IDisposable)finalBmp).Dispose();
													}
													((IDisposable)bmp).Dispose();
													((IDisposable)imgInstance).Dispose();

													mimeType = "image/png";
													return msOutput.ToArray();
												}
											}
										}
									}
									catch
									{
										// Fallback silently to direct save
									}

									using (var msOutput = new MemoryStream())
									{
										var pngFormat = pngProperty.GetValue(null);
										saveMethod.Invoke(imgInstance, new object[] { msOutput, pngFormat });

										((IDisposable)imgInstance).Dispose();

										mimeType = "image/png";
										return msOutput.ToArray();
									}
								}
							}
						}
					}
				}
			}
			catch
			{
				// Fallback silently
			}

			return bytes;
		}

		private static byte[] ExtractRasterFromEmf(byte[] emfBytes, out string mimeType)
		{
			mimeType = "image/x-emf";

			if (emfBytes == null || emfBytes.Length > MaxImageSizeBytes)
			{
				return emfBytes;
			}

			try
			{
				int offset = 0;
				while (offset < emfBytes.Length - 8)
				{
					uint type = BitConverter.ToUInt32(emfBytes, offset);
					uint size = BitConverter.ToUInt32(emfBytes, offset + 4);
					if (size < 8 || offset + size > emfBytes.Length)
					{
						break;
					}

					if (type == 81) // EMR_STRETCHDIBITS
					{
						int recordStart = offset;
						if (size >= 64)
						{
							uint offBmi = BitConverter.ToUInt32(emfBytes, recordStart + 48);
							uint cbBmi = BitConverter.ToUInt32(emfBytes, recordStart + 52);
							uint offBits = BitConverter.ToUInt32(emfBytes, recordStart + 56);
							uint cbBits = BitConverter.ToUInt32(emfBytes, recordStart + 60);

							if (cbBmi >= 40 && offBmi > 0 && recordStart + offBmi + cbBmi <= emfBytes.Length)
							{
								int bmiStart = (int)(recordStart + offBmi);
								uint compression = BitConverter.ToUInt32(emfBytes, bmiStart + 16);

								if (compression == 4 || compression == 5) // BI_JPEG = 4, BI_PNG = 5
								{
									if (cbBits > 0 && recordStart + offBits + cbBits <= emfBytes.Length)
									{
										byte[] imgBytes = new byte[cbBits];
										Array.Copy(emfBytes, (int)(recordStart + offBits), imgBytes, 0, (int)cbBits);
										mimeType = (compression == 4) ? "image/jpeg" : "image/png";
										return imgBytes;
									}
								}
								else // Raw bitmap (BI_RGB, BI_BITFIELDS, etc.)
								{
									if (cbBits > 0 && recordStart + offBits + cbBits <= emfBytes.Length)
									{
										byte[] bmpBytes = new byte[14 + cbBmi + cbBits];

										// BITMAPFILEHEADER
										bmpBytes[0] = 0x42; // 'B'
										bmpBytes[1] = 0x4D; // 'M'

										uint bfSize = (uint)(14 + cbBmi + cbBits);
										Array.Copy(BitConverter.GetBytes(bfSize), 0, bmpBytes, 2, 4);

										uint bfOffBits = (uint)(14 + cbBmi);
										Array.Copy(BitConverter.GetBytes(bfOffBits), 0, bmpBytes, 10, 4);

										// Copy BITMAPINFO
										Array.Copy(emfBytes, (int)(recordStart + offBmi), bmpBytes, 14, (int)cbBmi);

										// Copy Bits
										Array.Copy(emfBytes, (int)(recordStart + offBits), bmpBytes, (int)(14 + cbBmi), (int)cbBits);

										mimeType = "image/bmp";
										return bmpBytes;
									}
								}
							}
						}
					}

					offset += (int)size;
				}
			}
			catch
			{
				// Fallback to returning original EMF if parsing fails
			}

			return emfBytes;
		}
	}
}
