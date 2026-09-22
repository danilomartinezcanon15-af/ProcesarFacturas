using NetOffice.OutlookApi;
using NetOffice.OutlookApi.Enums;
using OutlookApp = NetOffice.OutlookApi.Application;
using System.Text.RegularExpressions;

namespace ProcesarFacturas.Services
{
    public class CorreoConAdjuntos
    {
        public string EntryID { get; set; } = string.Empty;
        public string Asunto { get; set; } = string.Empty;
        public string Categorias { get; set; } = string.Empty;
        public List<string> ArchivosAdjuntos { get; set; } = new List<string>();
    }

    public class OutlookReader
    {
        private static readonly List<string> PalabrasClaveAsunto = new()
        {
            "factura", "facturas", "informe", "layout", "cargue", "despacho", "despachadas", "fiy", "fwd", "rv"
        };

        public List<CorreoConAdjuntos> DescargarMensajesConAdjuntos(string rutaDestino, LoggerService logger, bool incluirProcesados = false)
        {
            var resultado = new List<CorreoConAdjuntos>();

            if (!Directory.Exists(rutaDestino))
            {
                Directory.CreateDirectory(rutaDestino);
            }

            using (var app = new OutlookApp())
            {
                var ns = app.GetNamespace("MAPI");
                var inbox = ns.GetDefaultFolder(OlDefaultFolders.olFolderInbox);

                ProcessFolder(inbox, ns, rutaDestino, resultado, logger, esBandejaEntrada: true);

                if (incluirProcesados)
                {
                    MAPIFolder? folderProcesados = null;
                    try
                    {
                        folderProcesados = inbox.Folders["Facturas_Procesadas"];
                    }
                    catch
                    {
                        logger.RegistrarAdvertencia("No se encontró la carpeta 'Facturas_Procesadas' en Outlook.");
                    }

                    if (folderProcesados != null)
                    {
                        logger.RegistrarInfo("Evaluando correos históricos en 'Facturas_Procesadas'...");
                        ProcessFolder(folderProcesados, ns, rutaDestino, resultado, logger, esBandejaEntrada: false);
                    }
                }
            }

            return resultado;
        }

        private void ProcessFolder(MAPIFolder folder, _NameSpace ns, string rutaDestino, List<CorreoConAdjuntos> listaMensajes, LoggerService logger, bool esBandejaEntrada)
        {
            var items = folder.Items;
            logger.RegistrarInfo($"Evaluando {items.Count} elementos en carpeta '{folder.Name}'...");

            foreach (var item in items)
            {
                if (item is MailItem mail)
                {
                    try
                    {
                        string asunto = mail.Subject ?? "";
                        string categorias = mail.Categories ?? "";

                        // Si estamos en Bandeja de Entrada y ya tiene la categoría [Procesado P&G], se omite
                        if (esBandejaEntrada && categorias.Contains("[Procesado P&G]"))
                        {
                            continue;
                        }

                        // Filtro por palabras clave en el Asunto
                        if (!EsAsuntoValido(asunto))
                        {
                            continue;
                        }

                        if (mail.Attachments.Count > 0)
                        {
                            var adjuntosDescargados = new List<string>();

                            // Extraer archivos directos o inspeccionar correos adjuntos (.msg / Elementos de Outlook)
                            ExtraerAdjuntosDeMail(mail, ns, rutaDestino, adjuntosDescargados, logger);

                            if (adjuntosDescargados.Count > 0)
                            {
                                listaMensajes.Add(new CorreoConAdjuntos
                                {
                                    EntryID = mail.EntryID,
                                    Asunto = asunto,
                                    Categorias = categorias,
                                    ArchivosAdjuntos = adjuntosDescargados
                                });
                            }
                        }
                    }
                    catch (System.Exception ex)
                    {
                        logger.RegistrarAdvertencia($"Error leyendo mensaje individual: {ex.Message}");
                    }
                }
            }
        }

        private void ExtraerAdjuntosDeMail(MailItem mail, _NameSpace ns, string rutaDestino, List<string> adjuntosDescargados, LoggerService logger)
        {
            foreach (Attachment attachment in mail.Attachments)
            {
                try
                {
                    string? fileName = attachment.FileName;

                    // CASO 1: Es un correo adjunto (Elemento de Outlook / Embedded Item)
                    if (attachment.Type == OlAttachmentType.olEmbeddeditem || (fileName != null && fileName.EndsWith(".msg", StringComparison.OrdinalIgnoreCase)))
                    {
                        string rutaMsgTemp = Path.Combine(rutaDestino, $"{Guid.NewGuid()}_submail.msg");
                        attachment.SaveAsFile(rutaMsgTemp);

                        // Abrir el correo adjunto en memoria
                        if (ns.OpenSharedItem(rutaMsgTemp) is MailItem subMail)
                        {
                            logger.RegistrarInfo($"Inspeccionando correo adjunto interno: '{subMail.Subject}'");
                            // Recursión: Extraer los Excels de este correo interno
                            ExtraerAdjuntosDeMail(subMail, ns, rutaDestino, adjuntosDescargados, logger);
                        }
                    }
                    // CASO 2: Es un archivo de hoja de cálculo (.xlsx, .xlsm, .xls)
                    else if (!string.IsNullOrEmpty(fileName))
                    {
                        string ext = Path.GetExtension(fileName).ToLower();

                        if (ext == ".xlsx" || ext == ".xlsm" || ext == ".xls")
                        {
                            string rutaArchivo = Path.Combine(rutaDestino, $"{Guid.NewGuid()}_{fileName}");
                            attachment.SaveAsFile(rutaArchivo);
                            adjuntosDescargados.Add(rutaArchivo);
                        }
                    }
                }
                catch (System.Exception ex)
                {
                    logger.RegistrarAdvertencia($"Ocurrió un detalle al extraer un adjunto: {ex.Message}");
                }
            }
        }

        private static bool EsAsuntoValido(string asunto)
        {
            if (string.IsNullOrWhiteSpace(asunto)) return false;
            string asuntoLimpio = Regex.Replace(asunto.ToLower(), @"[^a-z0-9]", " ");
            return PalabrasClaveAsunto.Any(p => asuntoLimpio.Contains(p));
        }

        public void MoverCorreosPorEntryId(List<string> entryIds, LoggerService logger)
        {
            if (entryIds == null || entryIds.Count == 0) return;

            using (var app = new OutlookApp())
            {
                var ns = app.GetNamespace("MAPI");

                try
                {
                    ns.Categories.Add("[Procesado P&G]", OlCategoryColor.olCategoryColorGreen, OlCategoryShortcutKey.olCategoryShortcutKeyNone);
                }
                catch
                {
                    // Si la categoría ya existe en Outlook, continúa sin lanzar excepción
                }

                var inbox = ns.GetDefaultFolder(OlDefaultFolders.olFolderInbox);

                MAPIFolder? destinoProcesados = null;
                try
                {
                    destinoProcesados = inbox.Folders["Facturas_Procesadas"];
                }
                catch
                {
                    logger.RegistrarAdvertencia("No se encontró la carpeta 'Facturas_Procesadas' para mover los correos.");
                    return;
                }

                foreach (var entryId in entryIds)
                {
                    try
                    {
                        var mail = ns.GetItemFromID(entryId) as MailItem;
                        if (mail != null)
                        {
                            string asunto = mail.Subject ?? "Sin asunto";
                            mail.Categories = "[Procesado P&G]";
                            mail.Save();
                            mail.Move(destinoProcesados);
                            logger.RegistrarInfo($"Correo '{asunto}' categorizado como [Procesado P&G] y movido.");
                        }
                    }
                    catch (System.Exception ex)
                    {
                        logger.RegistrarAdvertencia($"No se pudo mover el correo con EntryID {entryId}: {ex.Message}");
                    }
                }
            }
        }
    }
}