using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using Ghostscript.NET;
using Ghostscript.NET.Rasterizer;

namespace ZorgMetZorg.WebApplication.Utilities
{
    /// <summary>
    /// Utility class for converting PDF files to images using Ghostscript.NET library.
    /// </summary>
    public static class PDFToImageConverter
    {
        private const int DefaultDpi = 96;
        private static readonly ImageFormat DefaultImageFormat = ImageFormat.Png;

        #region Public

        /// <summary>
        /// Converts a document represented by its raw data (e.g., byte array or base64-encoded string)
        /// to a list of images in the specified format.
        /// The method utilizes Ghostscript for rasterization, allowing customization of DPI and image format.
        /// If the input document is null, the method throws an ArgumentException.
        /// </summary>
        /// <typeparam name="TInput">The type of the input document (e.g., byte[], string).</typeparam>
        /// <typeparam name="TOutput">The type of the output (e.g., byte[], string).</typeparam>
        /// <param name="input">The raw data representing the document.</param>
        /// <param name="dpi">The resolution in dots per inch (DPI) for rendering the images. Default value is 96.</param>
        /// <param name="imageFormat">The format in which the images should be saved (e.g., ImageFormat.Png). Default value is Png.</param>
        /// <returns>
        /// A list of objects representing images in the specified format (e.g., byte array, base64-encoded string).
        /// </returns>
        /// <exception cref="ArgumentException">
        /// Thrown when the input document is null.
        /// </exception>
        /// <exception cref="InvalidOperationException">
        /// Thrown when the conversion to the specified output type is not supported.
        /// </exception>
        public static List<TOutput> ConvertPdfDocumentToListOfImages<TInput, TOutput>(TInput input,
                                                                                      int dpi = DefaultDpi,
                                                                                      ImageFormat imageFormat = null)
        {
            if (input == null)
            {
                throw new ArgumentException("The input document cannot be null.", nameof(input));
            }

            // Set default ImageFormat to Png if not specified.
            imageFormat ??= DefaultImageFormat;

            var versionInfo = GetGhostscriptVersionInfo();
            var result = new List<TOutput>();

            using (var pdfStream = new MemoryStream(ConvertToByteArray(input)))
            using (var rasterizer = new GhostscriptRasterizer())
            {
                rasterizer.Open(pdfStream, versionInfo, true);

                for (int pageNumber = 1; pageNumber <= rasterizer.PageCount; pageNumber++)
                {
                    using (var pageImage = rasterizer.GetPage(dpi, pageNumber))
                    using (var memoryStream = new MemoryStream())
                    {
                        pageImage.Save(memoryStream, imageFormat);

                        if (typeof(TOutput) == typeof(string))
                        {
                            result.Add((TOutput)(object)Convert.ToBase64String(memoryStream.ToArray()));
                        }
                        else
                        {
                            result.Add((TOutput)(object)memoryStream.ToArray());
                        }
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// Converts a document represented by its raw data (e.g., byte array or base64-encoded string)
        /// to a single image in the specified format.
        /// The method utilizes Ghostscript for rasterization, allowing customization of DPI and image format.
        /// If the input document is null, the method throws an ArgumentException.
        /// </summary>
        /// <typeparam name="TInput">The type of the input document (e.g., byte[], string).</typeparam>
        /// <typeparam name="TOutput">The type of the output (e.g., byte[], string).</typeparam>
        /// <param name="input">The raw data representing the document.</param>
        /// <param name="dpi">The resolution in dots per inch (DPI) for rendering the images. Default value is 96.</param>
        /// <param name="imageFormat">The format in which the images should be saved (e.g., ImageFormat.Png). Default value is Png.</param>
        /// <returns>
        /// An object representing the image in the specified format (e.g., byte array, base64-encoded string).
        /// If the input document is null, the method throws an ArgumentException.
        /// </returns>
        /// <exception cref="ArgumentException">
        /// Thrown when the input document is null.
        /// </exception>
        /// <exception cref="InvalidOperationException">
        /// Thrown when the conversion to the specified output type is not supported.
        /// </exception>
        public static TOutput ConvertPdfDocumentToSingleImage<TInput, TOutput>(TInput input,
                                                                               int dpi = DefaultDpi,
                                                                               ImageFormat imageFormat = null)
        {
            if (input == null)
            {
                throw new ArgumentException("The input document cannot be null.", nameof(input));
            }

            // Set default ImageFormat to Png if not specified.
            imageFormat ??= DefaultImageFormat;

            var versionInfo = GetGhostscriptVersionInfo();

            using (var documentStream = new MemoryStream(ConvertToByteArray(input)))
            using (var rasterizer = new GhostscriptRasterizer())
            {
                rasterizer.Open(documentStream, versionInfo, true);

                int width = rasterizer.GetPage(dpi, 1).Width;
                int height = rasterizer.GetPage(dpi, 1).Height;

                Bitmap mergedImage = new Bitmap(width, height * rasterizer.PageCount);

                using (var graphics = Graphics.FromImage(mergedImage))
                {
                    for (int pageNumber = 1; pageNumber <= rasterizer.PageCount; pageNumber++)
                    {
                        using (var pageImage = rasterizer.GetPage(dpi, pageNumber))
                        {
                            graphics.DrawImage(pageImage, new Point(0, height * (pageNumber - 1)));
                        }
                    }
                }

                using (var memoryStream = new MemoryStream())
                {
                    mergedImage.Save(memoryStream, imageFormat);

                    if (typeof(TOutput) == typeof(string))
                    {
                        return (TOutput)(object)Convert.ToBase64String(memoryStream.ToArray());
                    }
                    else if (typeof(TOutput) == typeof(byte[]))
                    {
                        return (TOutput)(object)memoryStream.ToArray();
                    }

                    // Implement conversion logic for other types if needed.
                    throw new InvalidOperationException($"Conversion to type {typeof(TOutput)} is not supported.");
                }
            }
        }

        /// <summary>
        /// Reads the content of a PDF file specified by its file path into a byte array.
        /// </summary>
        /// <param name="filePath">The file path of the PDF file.</param>
        /// <returns>A byte array representing the content of the PDF file.</returns>
        /// <exception cref="FileNotFoundException">Thrown if the specified PDF file does not exist.</exception>
        public static byte[] GetPdfFileByteArray(string filePath)
        {
            // Check if the file exists
            if (!File.Exists(filePath))
            {
                throw new FileNotFoundException("The specified PDF file does not exist.", filePath);
            }

            // Read the content of the PDF file into a byte array
            return File.ReadAllBytes(filePath);
        }

        #endregion

        #region Private

        /// <summary>
        /// Converts the input of type T to a byte array.
        /// Supports conversion from base64-encoded strings or directly from byte arrays.
        /// </summary>
        /// <typeparam name="T">The type of input to be converted.</typeparam>
        /// <param name="input">The input value to be converted.</param>
        /// <returns>A byte array representation of the input.</returns>
        /// <exception cref="InvalidOperationException">
        /// Thrown if the input type is not supported for conversion.
        /// </exception>
        private static byte[] ConvertToByteArray<T>(T input)
        {
            if (input is string inputString)
            {
                return Convert.FromBase64String(inputString);
            }
            else if (input is byte[] inputByteArray)
            {
                return inputByteArray;
            }

            // Implement conversion logic for other types if needed.
            throw new InvalidOperationException($"Conversion from type {typeof(T)} is not supported.");
        }

        /// <summary>
        /// Retrieves version information for the Ghostscript library.
        /// </summary>
        /// <returns>
        /// The version information for the Ghostscript library.
        /// </returns>
        /// <remarks>
        /// The method constructs the file path to the Ghostscript library based on the current
        /// application's root directory. The appropriate Ghostscript DLL (gsdll64.dll or gsdll32.dll)
        /// is chosen depending on whether the current process is running as a 64-bit or 32-bit application.
        /// The Ghostscript DLL is expected to be located in a subdirectory named "Ghostscript" under
        /// the application's root directory.
        /// </remarks>
        private static GhostscriptVersionInfo GetGhostscriptVersionInfo()
        {
            string rootDirectory = AppDomain.CurrentDomain.BaseDirectory;
            var ghostDllFilePath = Environment.Is64BitProcess
                ? Path.Combine(rootDirectory, @"Ghostscript\gsdll64.dll")
                : Path.Combine(rootDirectory, @"Ghostscript\gsdll32.dll");
            return new GhostscriptVersionInfo(ghostDllFilePath);
        }

        #endregion
    }
}