using ImageMagick;
using ImageMagick.Formats;
using System;
using System.IO;
using System.Threading.Tasks;

namespace ManagerPdf.Services
{
    public class PdfToTiffService
    {
        public async Task<byte[]> ConvertPdfTiffAsync(Stream pdfStream, string outputFileName)
        {
            string tempPdfPath = Path.Combine(Path.GetTempPath(), outputFileName);
            string tempTiffPath = Path.Combine(Path.GetTempPath(), $"{Path.GetFileNameWithoutExtension(outputFileName)}.tiff");

            try
            {
                // Guardar el stream PDF en un archivo temporal
                using (var fileStream = new FileStream(tempPdfPath, FileMode.Create, FileAccess.Write))
                {
                    await pdfStream.CopyToAsync(fileStream);
                }

                // Configuración de ImageMagick para convertir PDF a TIFF
                using (var images = new MagickImageCollection())
                {
                    // Configurar opciones de lectura
                    var settings = new MagickReadSettings
                    {
                        Density = new Density(200, 200), // Resolución de 200 DPI
                        ColorType = ColorType.Bilevel, // Opción monocromática
                        BackgroundColor = MagickColors.White, // Fondo blanco
                    };

                    // Leer el archivo PDF
                    images.Read(tempPdfPath, settings);

                    foreach (var image in images)
                    {
                        image.Alpha(AlphaOption.Remove);
                        image.Format = MagickFormat.Tiff;
                        image.Settings.Compression = CompressionMethod.Group4;
                    }

                    // Combinar todas las imágenes en un solo archivo TIFF
                    //using (var tiffImage = images.Mosaic())
                    //{
                    //    // Guardar el resultado en el archivo TIFF temporal
                    //}
                        images.Write(tempTiffPath);
                }

                // Leer el archivo TIFF temporal y devolverlo como un byte array
                byte[] tiffBytes = await File.ReadAllBytesAsync(tempTiffPath);

                // Limpieza de archivos temporales
                File.Delete(tempPdfPath);
                File.Delete(tempTiffPath);

                return tiffBytes;
            }
            catch (Exception ex)
            {
                // Manejo de errores aquí (ejemplo: registrar el error)
                Console.WriteLine($"Error al convertir archivos PDF a TIFF: {ex.Message}");
                throw;
            }
        }
    }
}
