namespace ProcesarFacturas.Services
{
    public class LoggerService
    {
        private readonly string _rutaLog;

        public LoggerService()
        {
            string carpetaLogs = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Logs");
            if (!Directory.Exists(carpetaLogs))
            {
                Directory.CreateDirectory(carpetaLogs);
            }

            string nombreArchivo = $"Auditoria_{DateTime.Now:yyyy_MM_dd}.log";
            _rutaLog = Path.Combine(carpetaLogs, nombreArchivo);
        }

        public void RegistrarInfo(string mensaje)
        {
            Escribir("INFO", mensaje);
        }

        public void RegistrarAdvertencia(string mensaje)
        {
            Escribir("WARN", mensaje);
        }

        public void RegistrarError(string mensaje, Exception? ex = null)
        {
            string detalle = ex != null ? $"{mensaje} | Excepción: {ex.Message}" : mensaje;
            Escribir("ERROR", detalle);
        }

        private void Escribir(string nivel, string mensaje)
        {
            string linea = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] [{nivel}] {mensaje}";
            Console.WriteLine(linea); // Muestra en consola
            File.AppendAllText(_rutaLog, linea + Environment.NewLine); // Guarda en el .log
        }
    }
}