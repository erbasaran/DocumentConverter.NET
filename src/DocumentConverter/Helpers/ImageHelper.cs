namespace DocumentConverter.Helpers
{
	using System;
	using System.IO;
	using System.Runtime.InteropServices;

	public static class ImageHelper
	{
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
				case "emf":
					return "image/x-emf";
				case "wmf":
					return "image/x-wmf";
				default:
					return $"image/{extension}";
			}
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
