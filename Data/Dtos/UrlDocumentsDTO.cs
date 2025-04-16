namespace ManagerPdf.Data.Dtos
{
    public class UrlDocumentsDTO
    {
        public string client {  get; set; }
        public string solicitud { get; set; }
    }

    public class LoginFtpDTO
    {
        public string ftpUrl { get; set; }
        public string ftpUser { get; set; }
        public string ftpPassword { get; set; }
    }

    public class FtpConnectionDTO
    {
        public string Host { get; set; } // Dirección del servidor FTP
        public string Username { get; set; } // Usuario de FTP
        public string Password { get; set; } // Contraseña de FTP
        public int Port { get; set; } // Puerto (21 para FTP, 990 para FTPS implícito o personalizado)
        public bool UseFTPS { get; set; } // Especificar si es FTPS o FTP
        public bool ImplicitFTPS { get; set; } // Especificar si es FTPS implícito
    }

    public class UrlDocumentsPositivaDTO
    {
        public string cierre { get; set; }
        public string id { get; set; }
        public string operador { get; set; }
        public string placa { get; set;}
    }
}
