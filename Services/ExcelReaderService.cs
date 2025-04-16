using OfficeOpenXml;
using System.Collections.Generic;
using System.IO;


namespace ManagerPdf.Services
{

    public class ExcelReaderService
    {
        public Dictionary<string, (string Factura, string Soporte, string Autorizacion)> GetFileNames(IFormFile excelFile)
        {
            //var fileNames = new Dictionary<string, string>();
            var fileNames = new Dictionary<string, (string Factura, string Soporte, string Autorizacion)>();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var stream = new MemoryStream())
            {
                excelFile.CopyTo(stream);
                using (var package = new ExcelPackage(stream))
                {
                    var worksheet = package.Workbook.Worksheets[0];
                    int rowCount = worksheet.Dimension.Rows;
                    for (int row = 2; row <= rowCount; row++)
                    {
                        var factura = worksheet.Cells[row, 1].Text.Trim();
                        var soporte = worksheet.Cells[row, 2].Text.Trim();
                        var autorizacion = worksheet.Cells[row, 3].Text.Trim();
                        if (!string.IsNullOrEmpty(factura))
                        {
                            fileNames[factura] = (factura, soporte, autorizacion);
                        }
                    }
                }
            }

            return fileNames;
        }
    }
}