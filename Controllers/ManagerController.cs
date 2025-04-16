using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using PuppeteerSharp.Media;
using PuppeteerSharp;
using PdfOptions = PuppeteerSharp.PdfOptions;
using System.IO.Compression;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using ManagerPdf.Services;
using static PdfSharp.Snippets.Drawing.ImageHelper;
using System.Net;
using System.Runtime.Intrinsics.Arm;
using System.Threading;
using ManagerPdf.Data.Dtos;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.AspNetCore.Authorization;
using System.IO;

namespace ManagerPdf.Controllers
{
    [Route("api/[controller]")]
    [Authorize]
    [ApiController]
    public class ManagerController : ControllerBase
    {
        //se instancia la clase para poder usarla
        private readonly PdfToTiffService _pdfService;
        private readonly ExcelReaderService _readerExcelService;

        public ManagerController()
        {
            _pdfService = new PdfToTiffService();
            _readerExcelService = new ExcelReaderService();
        }


        [HttpPost("getPdf")]
        public async Task<IActionResult> GetPdf()
        {
            List<UrlDocumentsDTO> tasksErrors = new List<UrlDocumentsDTO>();
            IBrowser ErrorBrowser = null;
            SemaphoreSlim ErrorSemaphore = null;

            string cobrosFolderPath = "C:/UNIFIQUEPDF/cobros/";
            if (!Directory.Exists(cobrosFolderPath))
            {
                Directory.CreateDirectory(cobrosFolderPath);
            }

            // Eliminar archivos existentes en la carpeta
            foreach (var existingFile in Directory.GetFiles(cobrosFolderPath))
            {
                System.IO.File.Delete(existingFile);
            }

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            IFormFile file = Request.Form.Files[0];
            if (file == null || file.Length == 0)
                return Content("File not selected");

            using (var memoryStream = new MemoryStream())
            {
                await file.CopyToAsync(memoryStream);
                using (var package = new ExcelPackage(memoryStream))
                {
                    var worksheet = package.Workbook.Worksheets[0];
                    int rowCount = worksheet.Dimension.Rows;

                    // Inicializar el navegador
                    await new BrowserFetcher().DownloadAsync();
                    var browser = await Puppeteer.LaunchAsync(new LaunchOptions
                    {
                        Headless = true,
                        Args = new[] { "--no-sandbox" }
                    });

                    ErrorBrowser = browser;

                    // Iniciar sesión
                    var page = await browser.NewPageAsync();
                    await page.GoToAsync("http://operacion.etsg.com.co/accounts/login/", new NavigationOptions { Timeout = 60000 });
                    await page.TypeAsync("#user", "automatizacion@etsg.com.co");
                    await page.TypeAsync("#password", "123");
                    await page.ClickAsync(".btn.bg-Cprimary.text-white.btn-sm.btn-block");
                    //await page.WaitForSelectorAsync(".fa.fa-print.fa-2x.text-Csecondary.text-end"); // Espera a un selector específico después de iniciar sesión

                    // Configuración de concurrencia y semáforo
                    List<Task> tasks = new List<Task>();
                    int maxConcurrentRequests = Math.Min(10, rowCount);
                    SemaphoreSlim semaphore = new SemaphoreSlim(maxConcurrentRequests);
                    ErrorSemaphore = semaphore;


                    for (int row = 2; row <= rowCount; row++)
                    {
                        await semaphore.WaitAsync();

                        var cellData1 = worksheet.Cells[row, 1]?.Value?.ToString();
                        var cellData2 = worksheet.Cells[row, 2]?.Value?.ToString();
                        if (string.IsNullOrEmpty(cellData1) || string.IsNullOrEmpty(cellData2))
                        {
                            semaphore.Release();
                            continue;
                        }

                        var filePath = Path.Combine(cobrosFolderPath, $"{cellData2}.pdf");

                        tasks.Add(Task.Run(async () =>
                        {
                            int retryCount = 0;
                            bool pdfGenerated = false;

                            while (retryCount < 3 && !pdfGenerated)
                            {
                                try
                                {
                                    // Crear una nueva página para cada PDF
                                    var newPage = await browser.NewPageAsync();

                                    // Navegar a la URL y esperar a que cargue completamente
                                    await newPage.GoToAsync(
                                        $"http://operacion.etsg.com.co/servicios/certificadoSolicitud?cliente={cellData1}&solicitud={cellData2}",
                                        new NavigationOptions { Timeout = 60000 }
                                    );

                                    string bodyText = await newPage.EvaluateExpressionAsync<string>("document.body.innerText.trim()");

                                    // Comprobar si solo contiene el texto esperado
                                    if (bodyText == "El cliente seleccionado no corresponde al número de solicitud proporcionado" &&
                                        await newPage.EvaluateExpressionAsync<int>("document.body.children.length") <= 1)
                                    {
                                        Console.WriteLine($"La solicitud {cellData2} no contiene información");
                                        await newPage.CloseAsync();
                                        continue;
                                    }
                                    else
                                    {
                                        // Esperar a que un elemento clave esté cargado antes de generar el PDF
                                        await newPage.WaitForSelectorAsync(".fa.fa-print.fa-2x.text-Csecondary.text-end", new WaitForSelectorOptions { Timeout = 30000 });

                                        // Generar el PDF
                                        var pdfOptions = new PdfOptions { Format = PaperFormat.A4 };
                                        await newPage.PdfAsync(filePath, pdfOptions);
                                        await newPage.CloseAsync();

                                        pdfGenerated = true; // Marcar como exitoso si el PDF se genera correctamente
                                    }

                                }
                                catch (Exception ex)
                                {
                                    retryCount++;
                                    Console.WriteLine($"Error al generar el PDF para {cellData2}, intento {retryCount}: {ex.Message}");
                                    if (retryCount == 3)
                                    {
                                        tasksErrors.Add(new UrlDocumentsDTO
                                        {
                                            client = cellData1,
                                            solicitud = cellData2
                                        });
                                    }
                                }
                                finally
                                {
                                    semaphore.Release();
                                }
                            }
                        }));
                    }

                    await Task.WhenAll(tasks);

                    // Reintentos de tareas fallidas
                    if (tasksErrors.Count > 0)
                    {
                        await GenerateTasksError(tasksErrors, ErrorBrowser, ErrorSemaphore);
                    }

                    // Cerrar el navegador
                    await browser.CloseAsync();
                }
            }

            string zipPath = cobrosFolderPath + "result.zip";
            string tempZipPath = cobrosFolderPath + "temp_result.zip";

            try
            {
                using (FileStream zipToOpen = new FileStream(tempZipPath, FileMode.Create, FileAccess.ReadWrite, FileShare.None))
                {
                    using (var zip = new ZipArchive(zipToOpen, ZipArchiveMode.Update))
                    {
                        foreach (var file1 in Directory.EnumerateFiles(cobrosFolderPath, "*.pdf"))
                        {
                            zip.CreateEntryFromFile(file1, Path.GetFileName(file1));
                        }
                    }
                }
            }
            catch (IOException ex)
            {
                Console.WriteLine($"Error al acceder al archivo ZIP: {ex.Message}");
                return StatusCode(StatusCodes.Status500InternalServerError, "Error al acceder al archivo ZIP.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error inesperado: {ex.Message}");
                return StatusCode(StatusCodes.Status500InternalServerError, "Error inesperado.");
            }

            byte[] fileBytes;
            using (FileStream stream = new FileStream(tempZipPath, FileMode.Open))
            {
                fileBytes = new BinaryReader(stream).ReadBytes((int)stream.Length);
            }

            return File(fileBytes, "application/zip", "result.zip");
        }

        [HttpPost("unifique")]
        public async Task<IActionResult> Unifique()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            IFormFile file = Request.Form.Files[0];


            if (file == null || file.Length == 0)
                return Content("file not selected");

            using (var memoryStream = new MemoryStream())
            {
                await file.CopyToAsync(memoryStream);
                using (var package = new ExcelPackage(memoryStream))
                {
                    var worksheet = package.Workbook.Worksheets[0]; // Selecciona la primera hoja
                    int rowCount = worksheet.Dimension.Rows; // Obtiene el número de filas

                    for (int row = 2; row <= rowCount; row++)
                    {
                        // Lee los datos de cada celda en la fila
                        var cellData1 = worksheet.Cells[row, 1].Value.ToString();
                        var cellData2 = worksheet.Cells[row, 2].Value.ToString();
                        var cellData3 = worksheet.Cells[row, 3].Value.ToString();
                        var cellData4 = worksheet.Cells[row, 4].Value.ToString();

                        Directory.CreateDirectory("T:/Facturacion/UNIFIQUEPDF/RESULTS/" + cellData4);
                        var outputDocument = new PdfSharp.Pdf.PdfDocument();

                        Directory.CreateDirectory("T:/Facturacion/UNIFIQUEPDF/RESULTS/" + cellData4 + "/" + cellData3);

                        // Obtener todos los archivos PDF en la carpeta
                        var pdfFiles = Directory.GetFiles("T:/Facturacion/UNIFIQUEPDF/" + cellData4 + "/" + cellData1, "*.pdf");

                        foreach (var pdfFile in pdfFiles)
                        {
                            // Abrir el archivo PDF
                            var inputDocument = PdfSharp.Pdf.IO.PdfReader.Open(pdfFile, PdfSharp.Pdf.IO.PdfDocumentOpenMode.Import);

                            // Añadir cada página del archivo PDF al documento de salida
                            for (int i = 0; i < inputDocument.PageCount; i++)
                            {
                                var page = inputDocument.Pages[i];
                                outputDocument.AddPage(page);
                            }
                        }

                        // Guardar el documento de salida
                        outputDocument.Save("T:/Facturacion/UNIFIQUEPDF/RESULTS/" + cellData4 + "/" + cellData3 + "/" + cellData2 + ".pdf");


                    }
                }
            }

            return Ok(new { message = "Operacion realizada satisfactoriamente" });

        }

        [HttpPost("unifiquePdf")] 
        public async Task<IActionResult> UnifiquePdf(IFormFileCollection pdfFiles, string nameNewFile)
        {
            if (pdfFiles == null || !pdfFiles.Any())
                return Content("file not selected");
            try
            {
                using(PdfDocument outputDocument = new PdfSharp.Pdf.PdfDocument())
                {
                    foreach (var pdfFile in pdfFiles)
                    {
                        if (pdfFile == null || pdfFile.Length == 0)
                            continue;

                        using (var memoryStream = new MemoryStream())
                        {

                            await pdfFile.CopyToAsync(memoryStream);
                            memoryStream.Position = 0; // Resetear posición del stream

                            using(PdfDocument inputDocument = PdfReader.Open(memoryStream, PdfDocumentOpenMode.Import))
                            {
                                for (int i = 0; i < inputDocument.PageCount; i++)
                                {
                                    var page = inputDocument.Pages[i];
                                    outputDocument.AddPage(page);
                                }
                            }
                        }
                    }
                    using (MemoryStream outputStream = new MemoryStream())
                    {
                        outputDocument.Save(outputStream);
                        outputStream.Position = 0; // Resetear la posición del stream

                        // Retornar el PDF unificado como archivo descargable
                        return File(outputStream.ToArray(), "application/pdf", $"{nameNewFile}.pdf");
                    }
                }
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Error al unir los PDFs: {ex.Message}");
            }


        }

        [HttpPost("getPdfPositiva")]
        public async Task<IActionResult> GetPdfPositiva()
        {
            string supportFolderPath = "C:/UNIFIQUEPDF/RESULTS/POSITIVA_DOCS/";
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            IFormFile file = Request.Form.Files[0];
            if (file == null || file.Length == 0)
                return Content("file not selected");

            using (var memoryStream = new MemoryStream())
            {
                await file.CopyToAsync(memoryStream);
                using (var package = new ExcelPackage(memoryStream))
                {
                    var worksheet = package.Workbook.Worksheets[0]; // Selecciona la primera hoja
                    int rowCount = worksheet.Dimension.Rows; // Obtiene el número de filas

                    // List to store tasks for async operations
                    List<Task> tasks = new List<Task>();

                    int maxConcurrentRequests = Math.Min(10, rowCount);
                    SemaphoreSlim semaphore = new SemaphoreSlim(maxConcurrentRequests);

                    HashSet<string> createdFolders = new HashSet<string>();

                    for (int row = 2; row <= rowCount; row++)
                    {
                        var cellData1 = worksheet.Cells[row, 1]?.Value?.ToString();
                        var cellData2 = worksheet.Cells[row, 2]?.Value?.ToString();

                        if (string.IsNullOrEmpty(cellData1) || string.IsNullOrEmpty(cellData2))
                        {
                            continue;
                        }

                        string folderPath = Path.Combine(supportFolderPath, cellData1);

                        if (!createdFolders.Contains(folderPath))
                        {
                            Directory.CreateDirectory(folderPath);
                            createdFolders.Add(folderPath);
                        }

                        await semaphore.WaitAsync();
                        tasks.Add(Task.Run(async () =>
                        {
                            try
                            {
                                string url = $"https://positivacuida.positiva.gov.co/anexos/rs/getAnexosMultiplesTransaccionesTodasSucursales?transaccionesNos={cellData1},{cellData2}&tipoSolicitudId=1&tipoAnexo=anexo4";
                                string filePath = Path.Combine(folderPath, $"{cellData2}.pdf");

                                using (var client = new HttpClient())
                                {
                                    var response = await client.GetAsync(url);

                                    using (var fileStream = new FileStream(filePath, FileMode.Create, FileAccess.Write, FileShare.None))
                                    {
                                        await response.Content.CopyToAsync(fileStream);
                                    }
                                }

                            }
                            finally
                            {
                                semaphore.Release();
                            }
                        }));
                    }

                    await Task.WhenAll(tasks);
                }
            }

            string zipPath = supportFolderPath + "result.zip";
            try
            {
                using (FileStream zipToOpen = new FileStream(zipPath, FileMode.Create, FileAccess.ReadWrite, FileShare.None))
                {
                    using (var zip = new ZipArchive(zipToOpen, ZipArchiveMode.Update))
                    {
                        foreach (string directoryPath in Directory.GetDirectories(supportFolderPath, "*", SearchOption.TopDirectoryOnly))
                        {
                            PdfDocument outputDocument = new PdfDocument();
                            string unifiedPdfPath = Path.Combine(supportFolderPath, $"{Path.GetFileName(directoryPath)}.pdf");

                            foreach (string filePath in Directory.GetFiles(directoryPath, "*.pdf"))
                            {
                                Console.WriteLine(filePath);
                                // Abrir el documento existente
                                PdfDocument inputDocument = PdfReader.Open(filePath, PdfDocumentOpenMode.Import);

                                // Copiar cada página al documento de salida
                                for (int i = 0; i < inputDocument.PageCount; i++)
                                {
                                    PdfPage page = inputDocument.Pages[i];
                                    outputDocument.AddPage(page);
                                }
                            }

                            // Guardar el documento de salida
                            outputDocument.Save(unifiedPdfPath);

                            // Agregar el archivo unificado al archivo ZIP
                            string fileRelativePath = Path.GetRelativePath(supportFolderPath, unifiedPdfPath);
                            zip.CreateEntryFromFile(unifiedPdfPath, fileRelativePath);
                        }
                    }
                }
                foreach (string directoryPath in Directory.GetDirectories(supportFolderPath, "*", SearchOption.AllDirectories))
                {
                    Directory.Delete(directoryPath, recursive: true);
                }
            }
            catch (IOException ex)
            {
                // Manejo de excepciones, por ejemplo, loguear el error
                Console.WriteLine($"Error al acceder al archivo ZIP: {ex.Message}");
                return StatusCode(StatusCodes.Status500InternalServerError, "Error al acceder al archivo ZIP.");
            }
            catch (Exception ex)
            {
                // Manejo de otras excepciones
                Console.WriteLine($"Error inesperado: {ex.Message}");
                return StatusCode(StatusCodes.Status500InternalServerError, "Error inesperado.");
            }

            // Leer el archivo ZIP
            byte[] fileBytes;
            using (FileStream stream = new FileStream(zipPath, FileMode.Open))
            {
                fileBytes = new BinaryReader(stream).ReadBytes((int)stream.Length);
            }

            return File(fileBytes, "application/zip", "result.zip");
        }

        [HttpPost("getEditPdf")]
        public async Task<IActionResult> GetEditPdf()
        {
            string cobrosFolderPath = "C:/UNIFIQUEPDF/cobros/";

            List<UrlDocumentsPositivaDTO> tasksErrors = new List<UrlDocumentsPositivaDTO>();
            if (!Directory.Exists(cobrosFolderPath))
            {
                Directory.CreateDirectory(cobrosFolderPath);
            }
            foreach (var existingFile in Directory.GetFiles(cobrosFolderPath))
            {
                System.IO.File.Delete(existingFile);
            }
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            IFormFile file = Request.Form.Files[0];
            string routeSave = Request.Form["route"].ToString();
            if (file == null || file.Length == 0)
                return Content("file not selected");

            using (var memoryStream = new MemoryStream())
            {
                await file.CopyToAsync(memoryStream);
                using (var package = new ExcelPackage(memoryStream))
                {
                    var worksheet = package.Workbook.Worksheets[0]; // Selecciona la primera hoja
                    int rowCount = worksheet.Dimension.Rows; // Obtiene el número de filas

                    // Inicializar el navegador
                    await new BrowserFetcher().DownloadAsync();
                    var browser = await Puppeteer.LaunchAsync(new LaunchOptions
                    {
                        Headless = true,
                        Args = new[] { "--no-sandbox" }
                    });

                    // Iniciar sesión
                    var page = await browser.NewPageAsync();
                    await page.GoToAsync("http://operacion.etsg.com.co/accounts/login");
                    await page.TypeAsync("#user", "automatizacion@etsg.com.co");
                    await page.TypeAsync("#password", "123");
                    await page.ClickAsync(".btn.bg-Cprimary.text-white.btn-sm.btn-block");
                    await page.WaitForNavigationAsync();

                    string lastCellData1 = null;

                    List<Task> tasks = new List<Task>();
                    int maxConcurrentRequests = Math.Min(20, rowCount); // No más de 10 o la cantidad total de filas
                    SemaphoreSlim semaphore = new SemaphoreSlim(maxConcurrentRequests);

                    for (int row = 2; row <= rowCount; row++)
                    {
                        await semaphore.WaitAsync();

                        tasks.Add(Task.Run(async () =>
                        {
                            var cellData1 = worksheet.Cells[row, 1]?.Value?.ToString();
                            var cellData2 = worksheet.Cells[row, 2]?.Value?.ToString();
                            var cellData3 = worksheet.Cells[row, 3]?.Value?.ToString();
                            var cellData4 = worksheet.Cells[row, 4]?.Value?.ToString();
                            try
                            {

                                if (string.IsNullOrEmpty(cellData1) || string.IsNullOrEmpty(cellData2))
                                {
                                    return; // Si alguna celda está vacía, pasa a la siguiente fila
                                }

                                if (lastCellData1 == null || cellData1 != lastCellData1)
                                {
                                    if (lastCellData1 != null)
                                    {
                                        // Genera el PDF para el grupo anterior
                                        var pdfOptions = new PdfOptions { Format = PaperFormat.A4 };
                                        await page.PdfAsync(Path.Combine(cobrosFolderPath, $"{lastCellData1}.pdf"), pdfOptions);
                                    }

                                    await page.GoToAsync($"http://operacion.etsg.com.co/servicios/certificadoSolicitud?cliente=11&solicitud={cellData1}", new NavigationOptions { Timeout = 180000 });

                                    // Navega a la nueva URL
                                    lastCellData1 = cellData1;
                                }

                                string script = $@"
                                function updateOperator(id, newOperator) 
                                {{
                                    var rows = document.querySelectorAll('tbody tr');
                                    rows.forEach(function(row) 
                                    {{
                                        var idCell = row.querySelector('td:nth-child(2)');
                                        if (idCell.textContent.trim() === id) 
                                        {{
                                            var operatorCell = row.querySelector('td:nth-child(11)');
                                            operatorCell.textContent = newOperator;
                                        }}
                                    }});
                                }}
                                updateOperator('{cellData2}', '{cellData4 + "-" + cellData3}');";

                                await page.EvaluateExpressionAsync(script);
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"Error al generar el PDF para {cellData2}: {ex.Message}");
                                //tasksError.Add(cellData2,cellData1);
                                var docAdd = new UrlDocumentsPositivaDTO
                                {
                                    cierre = cellData1,
                                    id = cellData2,
                                    operador = cellData3,
                                    placa = cellData4
                                };
                                tasksErrors.Add(docAdd);
                            }
                            finally
                            {
                                semaphore.Release();
                            }
                        }));
                    }

                    await Task.WhenAll(tasks);

                    // Genera el PDF para el último grupo
                    if (lastCellData1 != null)
                    {
                        var pdfOptions = new PdfOptions { Format = PaperFormat.A4 };
                        await page.PdfAsync(Path.Combine(cobrosFolderPath, $"{lastCellData1}.pdf"), pdfOptions);
                    }

                    await page.CloseAsync();
                    await browser.CloseAsync();
                }
            }

            // Crear el archivo ZIP
            string zipPath = cobrosFolderPath + "result.zip";
            using (FileStream zipToOpen = new FileStream(zipPath, FileMode.Create, FileAccess.ReadWrite, FileShare.None))
            {
                using (var zip = new ZipArchive(zipToOpen, ZipArchiveMode.Update))
                {
                    foreach (var file1 in Directory.EnumerateFiles(cobrosFolderPath, "*.pdf"))
                    {
                        zip.CreateEntryFromFile(file1, Path.GetFileName(file1));
                    }
                }
            }

            // Leer el archivo ZIP
            byte[] fileBytes;
            using (FileStream stream = new FileStream(zipPath, FileMode.Open))
            {
                fileBytes = new BinaryReader(stream).ReadBytes((int)stream.Length);
            }

            return File(fileBytes, "application/zip", "result.zip");
        }


        [HttpPost("getFoldersPositiva")]
        public async Task<IActionResult> GetFoldersPositiva(IFormFileCollection Factura, IFormFileCollection Soporte, IFormFileCollection Autorizacion, IFormFile excelFile)
        {
            var semaphore = new SemaphoreSlim(3); // Limita a 3 las tareas simultáneas
            var tasks = new List<Task>();
            var tempFiles = new List<string>(); // Almacenar rutas de archivos temporales
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            // Crear un directorio temporal único
            string tempFolderPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString());
            Directory.CreateDirectory(tempFolderPath);

            try
            {
                using (var package = new ExcelPackage(excelFile.OpenReadStream()))
                {
                    var worksheet = package.Workbook.Worksheets[0];
                    for (int row = 2; row <= worksheet.Dimension.Rows; row++)
                    {
                        if (worksheet.Cells[row, 1].Value == null) continue;
                        var cellName = worksheet.Cells[row, 1].Value.ToString();

                        var fileCollections = new[] { Factura, Soporte, Autorizacion };
                        var standardNames = new[] { "FAC_", "EPI_", "ADS_" };

                        for (int col = 1; col <= worksheet.Dimension.Columns; col++)
                        {
                            var cellValue = worksheet.Cells[row, col].Value?.ToString().Trim();
                            if (string.IsNullOrEmpty(cellValue))
                            {
                                continue;
                            }

                            var file = fileCollections[col - 1].FirstOrDefault(f => Path.GetFileNameWithoutExtension(f.FileName) == cellValue);
                            if (file != null)
                            {
                                // Generar nombre único para el archivo
                                var standardFileName = $"{standardNames[col - 1]}900759329_{cellName}{Path.GetExtension(file.FileName)}";
                                var tempFilePath = Path.Combine(tempFolderPath, standardFileName);

                                // Guardar archivo temporalmente
                                tasks.Add(SaveFileToTempAsync(file, tempFilePath, semaphore, tempFiles));
                            }
                        }
                    }
                }

                await Task.WhenAll(tasks);

                // Crear ZIP en memoria
                var memoryStream = new MemoryStream();
                using (var zip = new ZipArchive(memoryStream, ZipArchiveMode.Create, true))
                {
                    foreach (var filePath in tempFiles)
                    {
                        var entryName = Path.GetFileName(filePath); // Solo el nombre del archivo, sin rutas
                        zip.CreateEntryFromFile(filePath, entryName);
                    }
                }

                // Limpiar archivos temporales
                foreach (var filePath in tempFiles)
                {
                    try { System.IO.File.Delete(filePath); } catch { /* Ignorar errores de eliminación */ }
                }
                try { Directory.Delete(tempFolderPath, true); } catch { /* Ignorar errores de eliminación */ }

                memoryStream.Position = 0;
                return File(memoryStream, "application/zip", "Resultados.zip");
            }
            catch
            {
                // Limpieza en caso de error
                foreach (var filePath in tempFiles)
                {
                    try { System.IO.File.Delete(filePath); } catch { }
                }
                try { Directory.Delete(tempFolderPath, true); } catch { }
                throw;
            }
        }

        /// Guarda un archivo en una ubicación temporal y registra la ruta para limpieza posterior
        private async Task SaveFileToTempAsync(IFormFile file, string filePath, SemaphoreSlim semaphore, List<string> tempFiles)
        {
            await semaphore.WaitAsync();
            try
            {
                // Asegurar nombre único
                filePath = GetUniqueFilePath(filePath);

                using (var stream = new FileStream(filePath, FileMode.Create))
                {
                    await file.CopyToAsync(stream);
                }

                lock (tempFiles)
                {
                    tempFiles.Add(filePath);
                }
            }
            finally
            {
                semaphore.Release();
            }
        }


        /// Genera una ruta de archivo única agregando un sufijo numérico si el archivo ya existe
        private string GetUniqueFilePath(string filePath)
        {
            int count = 1;
            string fileNameOnly = Path.GetFileNameWithoutExtension(filePath);
            string extension = Path.GetExtension(filePath);
            string path = Path.GetDirectoryName(filePath);
            string newFilePath = filePath;

            while (System.IO.File.Exists(newFilePath))
            {
                string tempFileName = $"{fileNameOnly}_{count++}";
                newFilePath = Path.Combine(path, tempFileName + extension);
            }
            return newFilePath;
        }

        [HttpPost("getZipSura")]
        public async Task<IActionResult> GetZipSura(IFormFileCollection Factura, IFormFileCollection Soporte, IFormFile excelFile, string date)
        {
            var semaphore = new SemaphoreSlim(3); // Limita a 3 las tareas simultáneas
            var tasks = new List<Task>();
            string cobrosFolderPath = "C:/UNIFIQUEPDF/cobros/";
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            //string date = Request.Form["date"].ToString();

            var newFolderPath = Path.Combine(cobrosFolderPath, "carpeta_facturas");
            if (!Directory.Exists(newFolderPath))
            {
                Directory.CreateDirectory(newFolderPath);
            }

            using (var package = new ExcelPackage(excelFile.OpenReadStream()))
            {
                var worksheet = package.Workbook.Worksheets[0];
                if (worksheet == null || worksheet.Dimension == null)
                {
                    throw new InvalidOperationException("Worksheet or its dimensions are not properly initialized.");
                }

                for (int row = 2; row <= worksheet.Dimension.Rows; row++)
                {
                    var facturaName = worksheet.Cells[row, 1].Value?.ToString().Trim();
                    if (string.IsNullOrEmpty(facturaName))
                    {
                        continue; // Salta filas con celdas vacías
                    }

                    var facturaGuion = facturaName.Insert(2, "_");
                    var soporteName = worksheet.Cells[row, 2].Value?.ToString().Trim();
                    var typeNp = worksheet.Cells[row, 3].Value?.ToString().Trim();
                    var valFact = worksheet.Cells[row, 4].Value?.ToString().Trim();

                    var rowFolderPath = Path.Combine(newFolderPath, $"900759329_{facturaGuion}_{valFact}_{typeNp}.zip");
                    var fileCollections = new[] { Factura, Soporte };
                    for (int col = 1; col <= 2; col++) // Ajustamos el rango a 1-2 para las dos columnas esperadas
                    {
                        var standardFileNameFact = $"IMGFACTURA_{soporteName}_{facturaName}_{date}";
                        var standardFileNameSopor = $"IMGSOPORTES_{soporteName}_{facturaName}_{date}";
                        var standardNames = new[] { standardFileNameFact, standardFileNameSopor };

                        var cellValue = worksheet.Cells[row, col].Value?.ToString().Trim(); // El nombre de lo que hay en la celda
                        if (string.IsNullOrEmpty(cellValue))
                        {
                            continue; // Salta celdas vacías
                        }

                        var file = fileCollections[col - 1]?.FirstOrDefault(f => Path.GetFileNameWithoutExtension(f.FileName) == cellValue);
                        if (file != null)
                        {
                            var standardFileName = standardNames[col - 1] + Path.GetExtension(file.FileName);
                            tasks.Add(SaveFilesZipAsync(file, rowFolderPath, standardFileName, semaphore));
                        }
                    }
                }
            }


            await Task.WhenAll(tasks);

            string zipPath = Path.Combine(Directory.GetCurrentDirectory(), "Result.zip");

            using (FileStream zipToOpen = new FileStream(zipPath, FileMode.Create))
            {
                using (ZipArchive zip = new ZipArchive(zipToOpen, ZipArchiveMode.Create))
                {
                    foreach (string filePath in Directory.GetFiles(newFolderPath))
                    {
                        string entryName = Path.GetFileName(filePath);
                        zip.CreateEntryFromFile(filePath, entryName);
                    }
                }
            }
            DirectoryInfo di = new DirectoryInfo(newFolderPath);
            foreach (FileInfo file in di.GetFiles("*.zip"))
            {
                file.Delete();
            }

            return File(System.IO.File.OpenRead(zipPath), "application/zip", "Result.zip");
        }

        private static readonly object zipLock = new object();

        private async Task SaveFilesZipAsync(IFormFile file, string zipPath, string standardFileName, SemaphoreSlim semaphore)
        {
            await semaphore.WaitAsync();
            try
            {
                using (var memoryStream = new MemoryStream())
                {
                    await file.CopyToAsync(memoryStream);
                    lock (zipLock)
                    {
                        using (var zipToOpen = new FileStream(zipPath, FileMode.OpenOrCreate))
                        {
                            using (var zip = new ZipArchive(zipToOpen, ZipArchiveMode.Update))
                            {
                                var zipEntry = zip.CreateEntry(standardFileName);
                                using (var entryStream = zipEntry.Open())
                                {
                                    memoryStream.Seek(0, SeekOrigin.Begin);
                                    memoryStream.CopyTo(entryStream);
                                }
                            }
                        }
                    }
                }
            }
            finally
            {
                semaphore.Release();
            }
        }


        [HttpPost("getPdfUnificated")]
        public IActionResult GetPdfUnificated(IFormFileCollection Factura, IFormFileCollection Soporte, IFormFileCollection Autorizacion, IFormFile excelFile)
        {
            string cobrosFolderPath = "C:/UNIFIQUEPDF/cobros/";
            string newFolderPath = Path.Combine(cobrosFolderPath, "carpeta_facturas");
            Directory.CreateDirectory(newFolderPath);
            string pathToPdf = Path.Combine(newFolderPath, "Result.pdf");

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var outputDocumentPdf = new PdfSharp.Pdf.PdfDocument();

            try
            {
                using (var package = new ExcelPackage(excelFile.OpenReadStream()))
                {
                    var worksheet = package.Workbook.Worksheets[0];
                    for (int row = 2; row <= worksheet.Dimension.Rows; row++)
                    {
                        var fileCollections = new[] { Factura, Soporte, Autorizacion };

                        for (int col = 1; col <= worksheet.Dimension.Columns; col++)
                        {
                            var cellValue = worksheet.Cells[row, col].Value?.ToString().Trim();

                            if (string.IsNullOrEmpty(cellValue)) continue;

                            var file = fileCollections[col - 1].FirstOrDefault(f => Path.GetFileNameWithoutExtension(f.FileName) == cellValue);

                            if (file != null)
                            {
                                var filePath = Path.Combine(newFolderPath, file.FileName);
                                using (var stream = new FileStream(filePath, FileMode.Create))
                                {
                                    file.CopyTo(stream);
                                }

                                var inputDocument = PdfReader.Open(filePath, PdfDocumentOpenMode.Import);
                                CopyPages(inputDocument, outputDocumentPdf);
                            }
                        }
                    }
                }

                outputDocumentPdf.Save(pathToPdf);

                if (System.IO.File.Exists(pathToPdf))
                {
                    var fileBytes = System.IO.File.ReadAllBytes(pathToPdf);
                    return File(fileBytes, "application/pdf", "Result.pdf");
                }
                else
                {
                    return StatusCode(500, "Error al generar el archivo PDF.");
                }
            }
            catch (Exception ex)
            {
                // Manejo de excepciones
                return StatusCode(500, $"Error interno del servidor: {ex.Message}");
            }
            finally
            {
                // Eliminar la carpeta y los archivos generados
                if (Directory.Exists(newFolderPath))
                {
                    Directory.Delete(newFolderPath, true);
                }
            }
        }

        void CopyPages(PdfDocument from, PdfDocument to)
        {
            for (int i = 0; i < from.PageCount; i++)
            {
                to.AddPage(from.Pages[i]);
            }
        }


        [HttpPost("convertPdfToTiff")]
        public async Task<IActionResult> ConvertPdfToTiff(IFormFileCollection Factura, IFormFileCollection Soporte, IFormFileCollection Autorizacion, IFormFile excelFile)
        {
            try
            {
                if ((Factura == null || !Factura.Any()) && (Soporte == null || !Soporte.Any()) && (Autorizacion == null || !Autorizacion.Any()) && (excelFile == null))
                {
                    return BadRequest("No files uploaded.");
                }

                var fileNames = _readerExcelService.GetFileNames(excelFile);
                var tasks = new List<Task<byte[]>>();
                var outputFileNames = new List<string>();

                // Crear semáforo con un límite de 50 tareas simultáneas
                var semaphore = new SemaphoreSlim(50);

                async Task<byte[]> ProcessWithSemaphore(IFormFile file)
                {
                    await semaphore.WaitAsync(); // Esperar hasta que haya espacio disponible
                    try
                    {
                        return await ProcessFile(file);
                    }
                    finally
                    {
                        semaphore.Release(); // Liberar espacio en el semáforo
                    }
                }

                foreach (var file in Factura ?? Enumerable.Empty<IFormFile>())
                {
                    tasks.Add(ProcessWithSemaphore(file));
                    outputFileNames.Add(GetNewFileName(file.FileName, fileNames, "FACTURA O DOCUMENTO EQUIVALENTE"));
                }

                foreach (var file in Soporte ?? Enumerable.Empty<IFormFile>())
                {
                    tasks.Add(ProcessWithSemaphore(file));
                    outputFileNames.Add(GetNewFileName(file.FileName, fileNames, "COMPROBANTE DE RECIBIDO DEL USUARIO"));
                }

                foreach (var file in Autorizacion ?? Enumerable.Empty<IFormFile>())
                {
                    tasks.Add(ProcessWithSemaphore(file));
                    outputFileNames.Add(GetNewFileName(file.FileName, fileNames, "OTROS"));
                }

                // Ejecutar todas las tareas en paralelo con el límite de 50
                var tiffFiles = await Task.WhenAll(tasks);

                var zipFile = CreateZipArchive(tiffFiles, outputFileNames);

                return File(zipFile, "application/zip", "converted_files.zip");
            }
            catch (Exception ex)
            {
                return StatusCode(StatusCodes.Status500InternalServerError, new { message = $"Error al convertir archivos PDF a TIFF y generar el archivo ZIP: {ex.Message}" });
            }
        }

        private async Task<byte[]> ProcessFile(IFormFile file)
        {
            using (var pdfStream = file.OpenReadStream())
            {
                return await _pdfService.ConvertPdfTiffAsync(pdfStream, file.FileName);
            }
        }

        private string GetNewFileName(string originalFileName, Dictionary<string, (string Factura, string Soporte, string Autorizacion)> fileNames, string suffix)
        {
            var baseName = Path.GetFileNameWithoutExtension(originalFileName);
            //Console.WriteLine($"Original file name: {originalFileName}, Base name: {baseName}");
            string nameFinal = baseName;
            foreach (var entry in fileNames)
            {
                if (entry.Value.Factura == baseName || entry.Value.Soporte == baseName || entry.Value.Autorizacion == baseName)
                {
                    string prefix = entry.Value.Factura;
                    nameFinal = $"{prefix}_900759329_{suffix}.tiff";
                }
            }
            return nameFinal;
        }

        private byte[] CreateZipArchive(byte[][] files, List<string> fileNames)
        {
            using (var memoryStream = new MemoryStream())
            {
                using (var zipArchive = new ZipArchive(memoryStream, ZipArchiveMode.Create, true))
                {
                    for (int i = 0; i < files.Length; i++)
                    {
                        var entry = zipArchive.CreateEntry(fileNames[i]);
                        using (var entryStream = entry.Open())
                        {
                            entryStream.Write(files[i], 0, files[i].Length);
                        }
                    }
                }
                return memoryStream.ToArray();
            }
        }


        [HttpPost("getPdfUnificatedMult")]
        public IActionResult GetPdfUnificatedMult(IFormFile excelFile, IFormFileCollection? pdfs0=null, IFormFileCollection? pdfs1 = null, IFormFileCollection? pdfs2 = null, IFormFileCollection? pdfs3 = null)
        {
            string cobrosFolderPath = "C:/UNIFIQUEPDF/cobros/";
            string newFolderPath = Path.Combine(cobrosFolderPath, "carpeta_facturas");
            Directory.CreateDirectory(newFolderPath);

            string generatedPdfFolderPath = Path.Combine(cobrosFolderPath, "carpeta_pdf_generados");
            Directory.CreateDirectory(generatedPdfFolderPath);


            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            

            try
            {
                using (var package = new ExcelPackage(excelFile.OpenReadStream()))
                {
                    var worksheet = package.Workbook.Worksheets[0];
                    for (int row = 2; row <= worksheet.Dimension.Rows; row++)
                    {
                        var fileCollections = new[] { pdfs0, pdfs1, pdfs2, pdfs3 };
                        if (worksheet.Cells[row, 1].Value == null) continue ;
                        var nameFact = worksheet.Cells[row, 1].Value.ToString();
                        
                        string pathToPdf = Path.Combine(newFolderPath, $"{nameFact}.pdf");
                        var outputDocumentPdf = new PdfSharp.Pdf.PdfDocument();
                        var tempFiles = new List<string>();
                        for (int col = 1; col <= worksheet.Dimension.Columns; col++)
                        {
                            var cellValue = worksheet.Cells[row, col].Value?.ToString().Trim();
                            
                            if (string.IsNullOrEmpty(cellValue)) continue;

                            var file = fileCollections[col - 1].FirstOrDefault(f => Path.GetFileNameWithoutExtension(f.FileName) == cellValue);

                            if (file != null)
                            {

                                var filePath = Path.Combine(newFolderPath, file.FileName);
                                using (var stream = new FileStream(filePath, FileMode.Create))
                                {
                                    file.CopyTo(stream);
                                }

                                var inputDocument = PdfReader.Open(filePath, PdfDocumentOpenMode.Import);
                                CopyPagesPdf(inputDocument, outputDocumentPdf);

                                tempFiles.Add(filePath);
                            }

                        }

                        string generatedPdfPath = Path.Combine(generatedPdfFolderPath, $"{nameFact}.pdf");
                        outputDocumentPdf.Save(generatedPdfPath);

                    }
                }
                string zipPath = Path.Combine(Directory.GetCurrentDirectory(), "Result.zip");
                

                using (FileStream zipToOpen = new FileStream(zipPath, FileMode.Create))
                {
                    using (ZipArchive zip = new ZipArchive(zipToOpen, ZipArchiveMode.Create))
                    {
                        foreach (string filePath in Directory.GetFiles(generatedPdfFolderPath))
                        {
                            string entryName = Path.GetFileName(filePath);
                            zip.CreateEntryFromFile(filePath, entryName);
                        }
                    }
                }
                

                return File(System.IO.File.OpenRead(zipPath), "application/zip", "Result.zip");

            }
            catch (Exception ex)
            {
                // Manejo de excepciones
                return StatusCode(500, $"Error interno del servidor: {ex.Message}");
            }
            finally
            {
                foreach (string folder in Directory.GetDirectories(cobrosFolderPath))
                {
                    Directory.Delete(folder, true);
                }
            }
        }

        [HttpPost("getPdfUnificatedMultAll")]
        public IActionResult GetPdfUnificatedMultAll(IFormFile excelFile, IFormFileCollection? pdfs0 = null, IFormFileCollection? pdfs1 = null, IFormFileCollection? pdfs2 = null, IFormFileCollection? pdfs3 = null)
        {
            string cobrosFolderPath = "C:/UNIFIQUEPDF/cobros/";
            string newFolderPath = Path.Combine(cobrosFolderPath, "carpeta_facturas");
            Directory.CreateDirectory(newFolderPath);

            string generatedPdfFolderPath = Path.Combine(cobrosFolderPath, "carpeta_pdf_generados");
            Directory.CreateDirectory(generatedPdfFolderPath);


            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;


            try
            {
                using (var package = new ExcelPackage(excelFile.OpenReadStream()))
                {
                    var worksheet = package.Workbook.Worksheets[0];
                    for (int row = 2; row <= worksheet.Dimension.Rows; row++)
                    {
                        var fileCollections = new[] { pdfs0, pdfs1, pdfs2, pdfs3 };
                        if (worksheet.Cells[row, 1].Value == null) continue;
                        var nameFact = worksheet.Cells[row, 1].Value.ToString();

                        string pathToPdf = Path.Combine(newFolderPath, $"{nameFact}.pdf");
                        var outputDocumentPdf = new PdfSharp.Pdf.PdfDocument();
                        var tempFiles = new List<string>();
                        for (int col = 1; col <= worksheet.Dimension.Columns; col++)
                        {
                            var cellValue = worksheet.Cells[row, col].Value?.ToString().Trim();

                            if (string.IsNullOrEmpty(cellValue)) continue;

                            var file = fileCollections[col - 1].FirstOrDefault(f => Path.GetFileNameWithoutExtension(f.FileName) == cellValue);

                            if (file != null)
                            {

                                var filePath = Path.Combine(newFolderPath, file.FileName);
                                using (var stream = new FileStream(filePath, FileMode.Create))
                                {
                                    file.CopyTo(stream);
                                }

                                var inputDocument = PdfReader.Open(filePath, PdfDocumentOpenMode.Import);
                                CopyPagesPdf(inputDocument, outputDocumentPdf);

                                tempFiles.Add(filePath);
                            }

                        }

                        string generatedPdfPath = Path.Combine(generatedPdfFolderPath, $"{nameFact}.pdf");
                        outputDocumentPdf.Save(generatedPdfPath);

                    }
                }
                string zipPath = Path.Combine(Directory.GetCurrentDirectory(), "Result.zip");


                using (FileStream zipToOpen = new FileStream(zipPath, FileMode.Create))
                {
                    using (ZipArchive zip = new ZipArchive(zipToOpen, ZipArchiveMode.Create))
                    {
                        foreach (string filePath in Directory.GetFiles(generatedPdfFolderPath))
                        {
                            string entryName = Path.GetFileName(filePath);
                            zip.CreateEntryFromFile(filePath, entryName);
                        }
                    }
                }


                return File(System.IO.File.OpenRead(zipPath), "application/zip", "Result.zip");

            }
            catch (Exception ex)
            {
                // Manejo de excepciones
                return StatusCode(500, $"Error interno del servidor: {ex.Message}");
            }
            finally
            {
                foreach (string folder in Directory.GetDirectories(cobrosFolderPath))
                {
                    Directory.Delete(folder, true);
                }
            }
        }



        void CopyPagesPdf(PdfDocument from, PdfDocument to)
        {
            for (int i = 0; i < from.PageCount; i++)
            {
                to.AddPage(from.Pages[i]);
            }
        }

        // Método de reintento para tareas fallidas
        async Task GenerateTasksError(List<UrlDocumentsDTO> tasksErrors, IBrowser browser, SemaphoreSlim semaphore)
        {
            string cobrosFolderPath = "C:/UNIFIQUEPDF/cobros/";
            foreach (var item in tasksErrors)
            {
                await semaphore.WaitAsync();
                var filePath = Path.Combine(cobrosFolderPath, $"{item.client}.pdf");
                try
                {
                    var newPage = await browser.NewPageAsync();
                    await newPage.GoToAsync(
                        $"http://operacion.etsg.com.co/servicios/certificadoSolicitud?cliente={item.client}&solicitud={item.solicitud}",
                        new NavigationOptions { Timeout = 60000 }
                    );
                    await newPage.WaitForSelectorAsync(".fa.fa-print.fa-2x.text-Csecondary.text-end", new WaitForSelectorOptions { Timeout = 30000 });
                    var pdfOptions = new PdfOptions { Format = PaperFormat.A4 };
                    await newPage.PdfAsync(filePath, pdfOptions);
                    await newPage.CloseAsync();
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error al generar el PDF para {item.solicitud}: {ex.Message}");
                }
                finally
                {
                    semaphore.Release();
                }
            }
        }

        async void GenerateTasksErrorPositiva(List<UrlDocumentsPositivaDTO> tasksErrors, IBrowser browser, SemaphoreSlim semaphore, string cobrosFolderPath)
        {
            string lastCellData1 = null;
            var page = await browser.NewPageAsync();
            foreach (var item in tasksErrors)
            {
                try
                {
                    if (lastCellData1 == null || item.cierre != lastCellData1)
                    {
                        if (lastCellData1 != null)
                        {
                            // Genera el PDF para el grupo anterior
                            var pdfOptions = new PdfOptions { Format = PaperFormat.A4 };
                            await page.PdfAsync(Path.Combine(cobrosFolderPath, $"{lastCellData1}.pdf"), pdfOptions);
                            await page.GoToAsync($"http://operacion.etsg.com.co/servicios/certificadoSolicitud?cliente=11&solicitud={item.cierre}", new NavigationOptions { Timeout = 180000 });
                        }

                        // Navega a la nueva URL
                        lastCellData1 = item.cierre;
                    }

                    string script = $@"
                                function updateOperator(id, newOperator) 
                                {{
                                    var rows = document.querySelectorAll('tbody tr');
                                    rows.forEach(function(row) 
                                    {{
                                        var idCell = row.querySelector('td:nth-child(2)');
                                        if (idCell.textContent.trim() === id) 
                                        {{
                                            var operatorCell = row.querySelector('td:nth-child(11)');
                                            operatorCell.textContent = newOperator;
                                        }}
                                    }});
                                }}
                                updateOperator('{item.id}', '{item.placa + "-" + item.operador}');";

                    await page.EvaluateExpressionAsync(script);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error al generar el PDF para {item.cierre} con id {item.id}: {ex.Message}");
                    //tasksError.Add(cellData2,cellData1);
                    var docAdd = new UrlDocumentsPositivaDTO
                    {
                        cierre = item.cierre,
                        id = item.id,
                        operador = item.operador,
                        placa = item.placa
                    };
                    tasksErrors.Add(docAdd);
                }
                finally
                {
                    semaphore.Release();
                }
            }
        }
    }
}
