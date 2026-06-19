namespace DocumentConverter.Helpers
{
	using System;
	using System.IO;
	using System.Runtime.InteropServices;

	public static class ImageHelper
	{
		private static readonly System.Net.Http.HttpClient HttpClientInstance = new System.Net.Http.HttpClient();

		public static byte[] GetImageBytes(string src)
		{
			if (string.IsNullOrEmpty(src)) return null;

			try
			{
				if (src.StartsWith("data:", StringComparison.OrdinalIgnoreCase))
				{
					var match = System.Text.RegularExpressions.Regex.Match(src, @"data:image/[^;]+;base64,(?<data>.+)");
					if (match.Success)
					{
						return System.Convert.FromBase64String(match.Groups["data"].Value);
					}
				}
				else if (src.StartsWith("http://", StringComparison.OrdinalIgnoreCase) || src.StartsWith("https://", StringComparison.OrdinalIgnoreCase))
				{
					using (var cts = new System.Threading.CancellationTokenSource(TimeSpan.FromSeconds(5)))
					{
						var responseTask = HttpClientInstance.GetAsync(src, cts.Token);
						responseTask.Wait(cts.Token);
						using (var response = responseTask.Result)
						{
							response.EnsureSuccessStatusCode();
							var readTask = response.Content.ReadAsByteArrayAsync();
							readTask.Wait(cts.Token);
							return readTask.Result;
						}
					}
				}
				else if (File.Exists(src))
				{
					return File.ReadAllBytes(src);
				}
			}
			catch
			{
				// Ignore load/download failures and return null
			}

			return null;
		}

		private static readonly System.Collections.Generic.Dictionary<string, string> MimeTypes = new System.Collections.Generic.Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
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
			{ "wmf", "image/x-wmf" }
		};

		public static bool TryGetImageDimensions(byte[] bytes, out int width, out int height)
		{
			width = 0;
			height = 0;
			if (bytes == null || bytes.Length < 8) return false;

			try
			{
				// 1. PNG check: starts with 89 50 4E 47 0D 0A 1A 0A
				if (bytes[0] == 0x89 && bytes[1] == 0x50 && bytes[2] == 0x4E && bytes[3] == 0x47)
				{
					// IHDR block starts at offset 12. Width is at 16 (4 bytes), Height is at 20 (4 bytes), big-endian
					if (bytes.Length >= 24)
					{
						width = (bytes[16] << 24) | (bytes[17] << 16) | (bytes[18] << 8) | bytes[19];
						height = (bytes[20] << 24) | (bytes[21] << 16) | (bytes[22] << 8) | bytes[23];
						return width > 0 && height > 0;
					}
				}

				// 2. GIF check: starts with 'GIF87a' or 'GIF89a' (47 49 46 38 37/39 61)
				if (bytes[0] == 0x47 && bytes[1] == 0x49 && bytes[2] == 0x46 && bytes[3] == 0x38)
				{
					// Width is at offset 6 (2 bytes, little-endian), Height is at 8 (2 bytes, little-endian)
					if (bytes.Length >= 10)
					{
						width = bytes[6] | (bytes[7] << 8);
						height = bytes[8] | (bytes[9] << 8);
						return width > 0 && height > 0;
					}
				}

				// 3. BMP check: starts with 'BM' (42 4D)
				if (bytes[0] == 0x42 && bytes[1] == 0x4D)
				{
					// Width is at offset 18 (4 bytes, little-endian), Height is at 22 (4 bytes, little-endian)
					if (bytes.Length >= 26)
					{
						width = bytes[18] | (bytes[19] << 8) | (bytes[20] << 16) | (bytes[21] << 24);
						height = bytes[22] | (bytes[23] << 8) | (bytes[24] << 16) | (bytes[25] << 24);
						height = Math.Abs(height);
						return width > 0 && height > 0;
					}
				}

				// 4. JPEG check: starts with FF D8
				if (bytes[0] == 0xFF && bytes[1] == 0xD8)
				{
					int offset = 2;
					while (offset < bytes.Length - 8)
					{
						byte markerStart = bytes[offset];
						byte markerType = bytes[offset + 1];
						if (markerStart != 0xFF) break;

						// Skip padding
						if (markerType == 0xFF)
						{
							offset++;
							continue;
						}

						int segmentLength = (bytes[offset + 2] << 8) | bytes[offset + 3];

						// SOF0 (Start Of Frame 0) marker is 0xC0, SOF2 is 0xC2, others 0xC1-0xCF except 0xC4, 0xC8, 0xCC
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
			string ext = extension.TrimStart('.').ToLowerInvariant();

			if (ext == "emf" || ext == "wmf" || ext == "x-emf" || ext == "x-wmf")
			{
				// Try to convert EMF/WMF to PNG if on Windows
				byte[] pngBytes = ConvertMetafileToPngIfWindows(rawData, out string convertedMime);
				if (convertedMime == "image/png")
				{
					mimeType = "image/png";
					return pngBytes;
				}

				// If not on Windows or conversion fails, try to extract raster from EMF
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
			mimeType = "image/x-metafile"; // Default fallback

			if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
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

											if (width > 0 && height > 0 && fromImageMethod != null && clearMethod != null && drawImageMethod != null && getPixelMethod != null)
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

												// Local function to check if pixel is active (ink) content
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
													// Add 4px safety padding
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
			mimeType = "image/x-emf"; // Default fallback if no DIB is found

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
