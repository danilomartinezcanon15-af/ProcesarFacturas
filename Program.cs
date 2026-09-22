using ProcesarFacturas.Models;
using ProcesarFacturas.Services;

namespace ProcesarFacturas
{
    internal class Program
    {
        static void Main(string[] args)
        {
            var logger = new LoggerService();
            logger.RegistrarInfo("=== INICIANDO PROCESO DE FACTURAS ===");

            try
            {
                string rutaBase = AppDomain.CurrentDomain.BaseDirectory;
                string rutaTemp = Path.Combine(rutaBase, "Temp");
                string rutaExcelDestino = Path.Combine(rutaBase, "Formato_Para_Cargues_P&G.xlsx");

                var outlookReader = new OutlookReader();
                var excelReader = new ExcelReader();
                var excelWriter = new ExcelWriter(rutaExcelDestino);

                logger.RegistrarInfo("Descargando adjuntos de Outlook...");
                var mensajes = outlookReader.DescargarMensajesConAdjuntos(rutaTemp, logger, incluirProcesados: true);

                var mapaEntryIdRegistros = new Dictionary<string, List<RegistroFactura>>();

                foreach (var mensaje in mensajes)
                {
                    var registrosDelCorreo = new List<RegistroFactura>();

                    foreach (var archivo in mensaje.ArchivosAdjuntos)
                    {
                        logger.RegistrarInfo($"Evaluando archivo: {Path.GetFileName(archivo)}");
                        var registros = excelReader.Leer(archivo);

                        if (registros.Count > 0)
                        {
                            logger.RegistrarInfo($"Registros válidos extraídos: {registros.Count}");
                            registrosDelCorreo.AddRange(registros);
                        }
                        else
                        {
                            logger.RegistrarAdvertencia($"El archivo {Path.GetFileName(archivo)} no contiene la estructura requerida. Omitiendo...");
                        }
                    }

                    if (registrosDelCorreo.Count > 0)
                    {
                        mapaEntryIdRegistros[mensaje.EntryID] = registrosDelCorreo;
                    }
                }

                int totalProcesados = 0;
                var entryIdsParaMover = new List<string>();

                foreach (var kvp in mapaEntryIdRegistros)
                {
                    string entryId = kvp.Key;
                    var registros = kvp.Value;

                    int agregados = excelWriter.AgregarEnHojaDbShipto("Db_Shipto", registros);

                    if (agregados > 0)
                    {
                        totalProcesados += agregados;
                        entryIdsParaMover.Add(entryId);
                    }
                }

                logger.RegistrarInfo($"Total de registros nuevos efectivamente agregados: {totalProcesados}");

                if (entryIdsParaMover.Count > 0)
                {
                    logger.RegistrarInfo($"Moviendo {entryIdsParaMover.Count} correos válidos a 'Facturas_Procesadas'...");
                    outlookReader.MoverCorreosPorEntryId(entryIdsParaMover, logger);
                }
                else
                {
                    logger.RegistrarInfo("No hubo correos con registros nuevos inéditos para mover.");
                }

                // Limpieza segura de la carpeta Temp
                LimpiarCarpetaTemporal(rutaTemp, logger);

                logger.RegistrarInfo("=== PROCESO FINALIZADO CORRECTAMENTE ===");
            }
            catch (Exception ex)
            {
                logger.RegistrarError("Error crítico durante la ejecución", ex);
            }
        }

        private static void LimpiarCarpetaTemporal(string rutaTemp, LoggerService logger)
        {
            if (!Directory.Exists(rutaTemp)) return;

            logger.RegistrarInfo("Limpiando archivos temporales...");

            // Forzar liberación de handles de archivos retenidos en memoria
            GC.Collect();
            GC.WaitForPendingFinalizers();

            foreach (var file in Directory.GetFiles(rutaTemp))
            {
                try
                {
                    File.Delete(file);
                }
                catch
                {
                    // Si OneDrive o un proceso retiene temporalmente un archivo, se ignora
                }
            }
        }
    }
}