using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ProcesarFacturas.Models;
using ProcesarFacturas.Services;

namespace ProcesarFacturas
{
    internal class Program
    {
        static void Main(string[] args)
        {
            var logger = new LoggerService();
            logger.RegistrarInfo("=== INICIANDO PROCESO DE FACTURAS (MODO OPTIMIZADO SQL SERVER) ===");

            try
            {
                string rutaBase = AppDomain.CurrentDomain.BaseDirectory;
                string rutaTemp = Path.Combine(rutaBase, "Temp");

                var outlookReader = new OutlookReader();
                var excelReader = new ExcelReader();
                var shiptoService = new ShiptoService();

                // 1. Carga de datos existentes a memoria local
                logger.RegistrarInfo("Cargando catálogo existente desde SQL Server...");
                HashSet<string> codigosExistentes = shiptoService.ObtenerCodigosExistentes();
                logger.RegistrarInfo($"Se cargaron {codigosExistentes.Count} registros existentes en memoria.");

                // 2. Descarga y procesamiento de correos
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

                // 3. Filtrado en memoria e inserción masiva (bulk insert)
                int totalProcesados = 0;
                var entryIdsParaMover = new List<string>();

                foreach (var kvp in mapaEntryIdRegistros)
                {
                    string entryId = kvp.Key;
                    var registrosDelCorreo = kvp.Value;

                    // Filtrar solo los registros que realmente son NUEVOS
                    var registrosNuevos = registrosDelCorreo
                        .Where(r => !string.IsNullOrWhiteSpace(r.ShipToParty) && !codigosExistentes.Contains(r.ShipToParty))
                        .GroupBy(r => r.ShipToParty, StringComparer.OrdinalIgnoreCase)
                        .Select(g => g.First())
                        .ToList();

                    if (registrosNuevos.Count > 0)
                    {
                        // Inserción en lote masivo a SQL Server
                        int insertados = shiptoService.GuardarNuevosShiptosMasivo(registrosNuevos);

                        if (insertados > 0)
                        {
                            totalProcesados += insertados;
                            entryIdsParaMover.Add(entryId);

                            // Actualizar catálogo local en memoria
                            foreach (var reg in registrosNuevos)
                            {
                                codigosExistentes.Add(reg.ShipToParty);
                            }
                        }
                    }
                }

                logger.RegistrarInfo($"Total de registros nuevos efectivamente agregados a SQL Server: {totalProcesados}");

                // 4. Mover correos a la carpeta procesada
                if (entryIdsParaMover.Count > 0)
                {
                    logger.RegistrarInfo($"Moviendo {entryIdsParaMover.Count} correos válidos a 'Facturas_Procesadas'...");
                    outlookReader.MoverCorreosPorEntryId(entryIdsParaMover, logger);
                }
                else
                {
                    logger.RegistrarInfo("No hubo correos con registros nuevos inéditos para mover.");
                }

                // 5. Limpieza de temporales
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
                    // Manejo de bloqueos por procesos externos
                }
            }
        }
    }
}