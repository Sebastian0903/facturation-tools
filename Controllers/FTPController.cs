using ManagerPdf.Data.Dtos;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Hosting;
using Renci.SshNet;
using System.Net;
using System.Net.Security;
using WinSCP;
using WinSCPSessionOptions = WinSCP.SessionOptions;

// For more information on enabling Web API for empty projects, visit https://go.microsoft.com/fwlink/?LinkID=397860

namespace ManagerPdf.Controllers
{
    [Authorize]
    [Route("api/[controller]")]
    [ApiController]
    public class FTPController : ControllerBase
    {
        //[HttpPost("getConnectionFtp")]
        //public IActionResult GetConnectionFtp([FromBody] LoginFtpDTO model)
        //{
        //    string ftpUrl ="";
        //    string ftpUser = model.ftpUser;
        //    string ftpPassword = model.ftpPassword;
        //    if (model.ftpUrl == "jota")
        //    {
        //        ftpUrl = "10.128.50.169";
        //    }

        //    if (model.ftpUrl == "sebastian")
        //    {
        //        ftpUrl = "10.128.50.68";
        //    }

        //    try
        //    {
        //        // Asegúrate de que la URL tenga el prefijo correcto y un puerto, si es necesario
        //        string urlFTP = ftpUrl.StartsWith("ftp://") ? ftpUrl : $"ftp://{ftpUrl}:21";

        //        // Configura la validación del certificado SSL
        //        ServicePointManager.ServerCertificateValidationCallback = delegate { return true; };

        //        FtpWebRequest request = (FtpWebRequest)WebRequest.Create(urlFTP);
        //        request.UsePassive = true;  // Cambia según sea necesario (true para modo pasivo)
        //        request.UseBinary = true;   // Asegura transferencia correcta de datos binarios

        //        // Intenta sin SSL primero
        //        request.EnableSsl = true;

        //        // Define el método a usar, por ejemplo, listar directorios o comprobar la conexión
        //        request.Method = WebRequestMethods.Ftp.PrintWorkingDirectory;
        //        //request.Method = WebRequestMethods.Ftp.ListDirectory;


        //        // Configura las credenciales
        //        request.Credentials = new NetworkCredential(ftpUser, ftpPassword);

        //        using (FtpWebResponse response = (FtpWebResponse)request.GetResponse())
        //        {
        //            if (response.StatusCode == FtpStatusCode.PathnameCreated ||
        //            response.StatusCode == FtpStatusCode.CommandOK ||
        //            response.StatusCode == FtpStatusCode.FileActionOK)
        //            {
        //                return Ok(new { Success = true, Message = "Conexión FTP exitosa." });
        //            }
        //            else
        //            {
        //                return StatusCode((int)response.StatusCode, $"Conexión fallida: {response.StatusDescription}");
        //            }
        //            //using (StreamReader reader = new StreamReader(response.GetResponseStream()))
        //            //{
        //            //    string directoryList = reader.ReadToEnd(); // Leer el listado de directorios
        //            //    return Ok(new { Success = true, Message = "Conexión FTP exitosa.", Directories = directoryList.Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries) });
        //            //}
        //        }
        //    }
        //    catch (WebException ex)
        //    {
        //        FtpWebResponse response = (FtpWebResponse)ex.Response;
        //        return BadRequest(new { Success = false, Message = $"Errors: {ex.Message}, Status: {response?.StatusDescription}" });
        //    }
        //    catch (Exception ex)
        //    {
        //        return BadRequest(new { Success = false, Message = $"Error: {ex.Message}" });
        //    }
        //}

        [HttpPost("connect")]
        public async Task<IActionResult> ConnectToSftp([FromBody] LoginFtpDTO loginFtpDTO)
        {
            string sftpHost = "";
            string sftpUser = loginFtpDTO.ftpUser;
            string sftpPassword = loginFtpDTO.ftpPassword;
            int port = 22;
            if (loginFtpDTO.ftpUrl == "capital")
            {
                sftpHost = "190.144.238.211";
            }

            if (loginFtpDTO.ftpUrl == "nueva")
            {
                sftpHost = "ftp.nuevaeps.com.co";
            }

            if (loginFtpDTO == null || string.IsNullOrEmpty(sftpHost) ||
                string.IsNullOrEmpty(loginFtpDTO.ftpUser) || string.IsNullOrEmpty(loginFtpDTO.ftpPassword))
            {
                return BadRequest(new { Success = false, Message = "Los datos del SFTP son requeridos." });
            }

            var connectionInfo = new Renci.SshNet.ConnectionInfo(sftpHost, port, sftpUser,
            new PasswordAuthenticationMethod(sftpUser, sftpPassword))
            {
                Timeout = TimeSpan.FromMinutes(2) // Cambia el tiempo según lo necesites (en este caso, 5 minutos)
            };

            using (var client = new SftpClient(connectionInfo))
            {
                try
                {
                    client.Connect();
                    Console.WriteLine("Conexión SFTP exitosa.");

                    var files = client.ListDirectory("/");

                    var directoriesAndFiles = files.Select(file => new
                    {
                        file.Name,
                        file.FullName,
                        file.Length,
                        file.LastWriteTime
                    }).ToList();

                    client.Disconnect();

                    return Ok(new { Success = true, Message = "Conexión SFTP exitosa.", Data = directoriesAndFiles });
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error al conectar al SFTP: {ex.Message}");
                    return BadRequest(new { Success = false, Message = $"Error de conexión: {ex.Message}" });
                }
            }
        }

        [HttpPost("uploadFilesToSftp")]
        public async Task<IActionResult> UploadFilesToSftp(IFormFileCollection files, string sftpHost, string sftpUser, string sftpPassword)
        {
            int port = 22;
            if (sftpHost == "capital")
            {
                sftpHost = "190.144.238.211";
            }

            if (sftpHost == "nueva")
            {
                sftpHost = "ftp.nuevaeps.com.co";
            }

            try
            {
                if (files == null || files.Count == 0)
                {
                    return BadRequest(new { Success = false, Message = "No se ha proporcionado ningún archivo para cargar." });
                }

                List<string> resultados = new List<string>();

                using (var client = new SftpClient(sftpHost, port, sftpUser, sftpPassword))
                {
                    try
                    {
                        client.Connect();
                        Console.WriteLine("Conexión SFTP exitosa."); // Mensaje de éxito

                        foreach (var file in files)
                        {
                            try
                            {
                                // Crea un flujo de memoria para el archivo
                                using (var stream = file.OpenReadStream())
                                {
                                    // Utiliza UploadFile para cargar el archivo desde el flujo
                                    client.UploadFile(stream, file.FileName); // Carga el archivo desde el flujo
                                    resultados.Add($"Archivo {file.FileName} cargado exitosamente.");
                                }
                            }
                            catch (Exception ex)
                            {
                                resultados.Add($"Error al cargar el archivo {file.FileName}: {ex.Message}");
                            }
                        }

                        client.Disconnect();
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error al conectar al SFTP: {ex.Message}"); // Mensaje de error
                        return BadRequest(new { Success = false, Message = $"Error de conexión: {ex.Message}" });
                    }
                }

                return Ok(new { Success = true, Message = "Proceso completado.", Resultados = resultados });
            }
            catch (Exception ex)
            {
                return BadRequest(new { Success = false, Message = $"Error general: {ex.Message}" });
            }
        }

        [HttpPost("connectff")]
        public async Task<IActionResult> ConnectToFtp([FromBody] LoginFtpDTO loginFtpDTO)
        {
            string ftpUrl = loginFtpDTO.ftpUrl;
            int port = 21; // Por defecto FTP (puerto para Capital Salud)
            bool isSftp = false; // Bandera para SFTP

            if (ftpUrl == "capital")
            {
                ftpUrl = "190.144.238.211";
                port = 22; // SFTP para Capital Salud
                isSftp = true; // Cambiar a SFTP
            }
            else if (ftpUrl == "nueva")
            {
                // Ajuste según la información provista para Nueva EPS (FTPS)
                ftpUrl = "ftp.nuevaeps.com.co";
                port = 990; // Puerto para FTPS explícito
            }

            string ftpUser = loginFtpDTO.ftpUser;
            string ftpPassword = loginFtpDTO.ftpPassword;

            if (loginFtpDTO == null || string.IsNullOrEmpty(ftpUrl) ||
                string.IsNullOrEmpty(ftpUser) || string.IsNullOrEmpty(ftpPassword))
            {
                return BadRequest(new { Success = false, Message = "Los datos del FTP son requeridos." });
            }

            // Configuración de la sesión
            WinSCPSessionOptions sessionOptions = new WinSCPSessionOptions
            {
                Protocol = isSftp ? Protocol.Sftp : Protocol.Ftp,
                HostName = ftpUrl,
                UserName = ftpUser,
                Password = ftpPassword,
                PortNumber = port,
                FtpSecure = isSftp ? FtpSecure.Explicit : FtpSecure.Implicit // Ajuste de encriptación
            };
            string winscpExecutablePath = Path.Combine("FTPexe", "WinSCP.exe");
            using (var session = new WinSCP.Session())
            {
                try
                {
                    session.ExecutablePath = winscpExecutablePath;
                    session.Open(sessionOptions);
                    Console.WriteLine("Conexión FTP exitosa.");

                    // Listar archivos y directorios en el directorio raíz
                    RemoteDirectoryInfo transferResult = session.ListDirectory("/");
                    var directoriesAndFiles = transferResult.Files.Select(item => new
                    {
                        Name = item.Name,
                        FullName = item.FullName,
                        Length = item.Length,               // Cambiado a Length
                        LastWriteTime = item.LastWriteTime,    // Cambiado a LastWriteTime
                        IsDirectory = item.IsDirectory
                    }).ToList();

                    return Ok(new { Success = true, Message = "Conexión FTP exitosa.", Data = directoriesAndFiles });
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error al conectar al FTP: {ex.Message}");
                    return BadRequest(new { Success = false, Message = $"Error de conexión: {ex.Message}" });
                }
            }
        }

        [HttpPost("uploadFilesToFtp")]
        public IActionResult UploadFilesToFtp(IFormFileCollection files, string ftpUrl, string ftpUser, string ftpPassword)
        {
            string rutaFtp = "";
            if (ftpUrl == "capital")
            {
                rutaFtp = "190.144.238.211";
            }
            else if (ftpUrl == "nueva")
            {
                rutaFtp = "ftp.nuevaeps.com.co";
            }
            // Configuración de la sesión
            WinSCPSessionOptions sessionOptions = new WinSCPSessionOptions
            {
                Protocol = Protocol.Ftp,
                HostName = rutaFtp,
                UserName = ftpUser,
                Password = ftpPassword,
                FtpSecure = FtpSecure.Implicit,
                PortNumber = 990 
            };

            if (ftpUrl == "capital")
            {
                sessionOptions.Protocol = Protocol.Sftp; // Cambiar a SFTP
                sessionOptions.PortNumber = 22; // SFTP para Capital Salud
            }
            else if (ftpUrl == "nueva")
            {
                sessionOptions.PortNumber = 990; // Puerto para FTPS explícito
            }

            try
            {
                if (files == null || files.Count == 0)
                {
                    return BadRequest(new { Success = false, Message = "No se ha proporcionado ningún archivo para cargar." });
                }
                string winscpExecutablePath = Path.Combine("FTPexe", "WinSCP.exe");
                List<string> resultados = new List<string>();

                using (var session = new WinSCP.Session())
                {
                    // Conectar a la sesión
                    session.Open(sessionOptions);

                    foreach (var file in files)
                    {
                        try
                        {
                            // Verificar si el archivo tiene contenido
                            if (file.Length == 0)
                            {
                                resultados.Add($"El archivo {file.FileName} está vacío.");
                                continue;
                            }

                            // Guarda el archivo temporalmente
                            var tempFilePath = Path.GetTempPath() + file.FileName;

                            using (var stream = file.OpenReadStream())
                            {
                                // Guarda el archivo temporalmente en el disco
                                using (var fileStream = new FileStream(tempFilePath, FileMode.Create, FileAccess.Write))
                                {
                                    stream.CopyTo(fileStream);
                                }
                            }

                            // Sube el archivo
                            TransferOptions transferOptions = new TransferOptions
                            {
                                TransferMode = TransferMode.Binary // O TransferMode.Text si es un archivo de texto
                            };

                            TransferOperationResult transferResult;
                            transferResult = session.PutFiles(tempFilePath, "/" + file.FileName, false, transferOptions);

                            // Comprobar resultados
                            transferResult.Check();
                            foreach (TransferEventArgs transfer in transferResult.Transfers)
                            {
                                resultados.Add($"Archivo {file.FileName} cargado exitosamente.");
                            }

                            // Elimina el archivo temporal
                            System.IO.File.Delete(tempFilePath);
                        }
                        catch (Exception ex)
                        {
                            resultados.Add($"Error al cargar el archivo {file.FileName}: {ex.Message}");
                        }
                    }
                }

                return Ok(new { Success = true, Message = "Proceso completado.", Resultados = resultados });
            }
            catch (Exception ex)
            {
                return BadRequest(new { Success = false, Message = $"Error general: {ex.Message}" });
            }
        }
    }
}
