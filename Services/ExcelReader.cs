using ExcelDataReader;
using ProcesarFacturas.Models;
using System.Data;
using System.Text;
using System.Text.RegularExpressions;

namespace ProcesarFacturas.Services
{
    public class ExcelReader
    {
        // Registro del proveedor de codificación requerido por ExcelDataReader para .xls
        static ExcelReader()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        }

        // DICCIONARIO / GLOSARIO DE PALABRAS CLAVE Y SUS SINÓNIMOS
        private static readonly List<string> SinonimosShipTo = new() 
        { 
            "shipto", "shiptoparty", "solicitante", "destinatario", "codcliente", "codigo", "shipto", "ship_to", "cliente_cod" 
        };

        private static readonly List<string> SinonimosName = new() 
        { 
            "name1", "name", "nombre", "cliente", "razonsocial", "nombrecliente", "nom_cliente", "destinatariog" 
        };

        private static readonly List<string> SinonimosStreet = new() 
        { 
            "street", "direccion", "dir", "direccionentrega", "domicilio", "calle" 
        };

        private static readonly List<string> SinonimosCity = new() 
        { 
            "city", "ciudad", "poblacion", "municipio", "distrito" 
        };

        public List<RegistroFactura> Leer(string rutaArchivo)
        {
            var registros = new List<RegistroFactura>();

            if (!File.Exists(rutaArchivo)) return registros;

            string extension = Path.GetExtension(rutaArchivo).ToLower();
            if (extension != ".xlsx" && extension != ".xlsm" && extension != ".xls") return registros;

            try
            {
                using (var stream = File.Open(rutaArchivo, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        var result = reader.AsDataSet(new ExcelDataSetConfiguration()
                        {
                            ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                            {
                                UseHeaderRow = false // Leemos filas crudas para buscar encabezados dinámicos
                            }
                        });

                        if (result.Tables.Count == 0) return registros;

                        DataTable table = result.Tables[0];
                        int ultimaFila = table.Rows.Count;
                        int ultimaColumna = table.Columns.Count;

                        if (ultimaFila <= 1) return registros;

                        int filaEncabezado = -1; // 0-indexed
                        int colShipTo = -1, colName = -1, colStreet = -1, colCity = -1;

                        // Escanear las primeras 15 filas usando el Glosario de Sinónimos
                        int maxFilaEscanear = Math.Min(15, ultimaFila);
                        for (int f = 0; f < maxFilaEscanear; f++)
                        {
                            int cShip = -1, cName = -1, cStreet = -1, cCity = -1;

                            for (int col = 0; col < ultimaColumna; col++)
                            {
                                string cellVal = table.Rows[f][col]?.ToString() ?? "";
                                string headerLimpio = NormalizarTexto(cellVal);

                                if (cShip == -1 && SinonimosShipTo.Any(s => headerLimpio.Contains(s))) cShip = col;
                                if (cName == -1 && SinonimosName.Any(s => headerLimpio.Contains(s))) cName = col;
                                if (cStreet == -1 && SinonimosStreet.Any(s => headerLimpio.Contains(s))) cStreet = col;
                                if (cCity == -1 && SinonimosCity.Any(s => headerLimpio.Contains(s))) cCity = col;
                            }

                            if (cShip != -1)
                            {
                                filaEncabezado = f;
                                colShipTo = cShip;
                                colName = cName;
                                colStreet = cStreet;
                                colCity = cCity;
                                break;
                            }
                        }

                        if (filaEncabezado == -1 || colShipTo == -1) return registros;

                        // Extraer registros a partir de la fila siguiente al encabezado
                        for (int fila = filaEncabezado + 1; fila < ultimaFila; fila++)
                        {
                            string shipTo = table.Rows[fila][colShipTo]?.ToString()?.Trim() ?? "";
                            string name = colName != -1 ? table.Rows[fila][colName]?.ToString()?.Trim() ?? "" : "";
                            string street = colStreet != -1 ? table.Rows[fila][colStreet]?.ToString()?.Trim() ?? "" : "";
                            string city = colCity != -1 ? table.Rows[fila][colCity]?.ToString()?.Trim() ?? "" : "";

                            // FILTRO DE VALIDACIÓN: El Ship-to debe ser un código válido
                            if (EsShipToValido(shipTo))
                            {
                                var reg = new RegistroFactura
                                {
                                    ShipToParty = shipTo,
                                    Cliente = name,
                                    Direccion = street,
                                    Ciudad = city
                                };

                                reg.Datos["Ship-to"] = shipTo;
                                reg.Datos["Name 1"] = name;
                                reg.Datos["Street"] = street;
                                reg.Datos["City"] = city;

                                registros.Add(reg);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[WARN] No se pudo leer el archivo {Path.GetFileName(rutaArchivo)}. Detalle: {ex.Message}");
            }

            return registros;
        }

        private static string NormalizarTexto(string texto)
        {
            if (string.IsNullOrWhiteSpace(texto)) return "";
            return Regex.Replace(texto.ToLower(), @"[^a-z0-9]", "");
        }

        private static bool EsShipToValido(string shipTo)
        {
            if (string.IsNullOrWhiteSpace(shipTo)) return false;
            if (shipTo.Length < 4) return false;

            string shipToLimpio = NormalizarTexto(shipTo);

            // Palabras clave no permitidas (firmas/pies de página)
            string[] palabrasInvalidas = { "cantidad", "lasalida", "autoriza", "concepto", "quiensoi", "observacion", "total", "subtotal", "firma" };
            if (palabrasInvalidas.Any(p => shipToLimpio.Contains(p))) return false;

            // Debe contener al menos 3 dígitos numéricos para ser un Ship-To legítimo
            int digitosCount = shipTo.Count(char.IsDigit);
            return digitosCount >= 3;
        }
    }
}