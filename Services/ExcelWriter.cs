using ClosedXML.Excel;
using ProcesarFacturas.Models;

namespace ProcesarFacturas.Services
{
    public class ExcelWriter
    {
        private readonly string _rutaExcelDestino;

        public ExcelWriter(string rutaExcelDestino)
        {
            _rutaExcelDestino = rutaExcelDestino;
        }

        public bool ExisteHoja(string nombreHoja)
        {
            if (!File.Exists(_rutaExcelDestino)) return false;

            using (var workbook = new XLWorkbook(_rutaExcelDestino))
            {
                return workbook.Worksheets.Contains(nombreHoja);
            }
        }

        public int AgregarEnHojaDbShipto(string nombreHoja, List<RegistroFactura> nuevosRegistros)
        {
            if (nuevosRegistros == null || nuevosRegistros.Count == 0) return 0;

            using (var workbook = new XLWorkbook(_rutaExcelDestino))
            {
                IXLWorksheet? sheet;
                if (!workbook.Worksheets.TryGetWorksheet(nombreHoja, out sheet) || sheet == null)
                {
                    sheet = workbook.Worksheets.Add(nombreHoja);
                    sheet.Cell(1, 1).Value = "Ship-to";
                    sheet.Cell(1, 2).Value = "Name 1";
                    sheet.Cell(1, 3).Value = "Street";
                    sheet.Cell(1, 4).Value = "City";
                }

                var shipTosExistentes = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                var lastRowUsed = sheet.LastRowUsed();
                int ultimaFilaUsada = lastRowUsed != null ? lastRowUsed.RowNumber() : 0;

                if (ultimaFilaUsada > 1)
                {
                    for (int r = 2; r <= ultimaFilaUsada; r++)
                    {
                        var cellValue = sheet.Cell(r, 1).Value;
                        string? val = cellValue.IsBlank ? null : cellValue.ToString()?.Trim();
                        if (!string.IsNullOrEmpty(val))
                        {
                            shipTosExistentes.Add(val);
                        }
                    }
                }

                int filaActual = ultimaFilaUsada <= 1 ? 2 : ultimaFilaUsada + 1;
                int agregadosCount = 0;

                foreach (var reg in nuevosRegistros)
                {
                    string shipTo = reg.ObtenerValor("Ship-to");
                    if (string.IsNullOrWhiteSpace(shipTo)) continue;

                    if (!shipTosExistentes.Contains(shipTo.Trim()))
                    {
                        sheet.Cell(filaActual, 1).Value = shipTo;
                        sheet.Cell(filaActual, 2).Value = reg.ObtenerValor("Name 1");
                        sheet.Cell(filaActual, 3).Value = reg.ObtenerValor("Street");
                        sheet.Cell(filaActual, 4).Value = reg.ObtenerValor("City");

                        shipTosExistentes.Add(shipTo.Trim());
                        filaActual++;
                        agregadosCount++;
                    }
                }

                workbook.Save();
                return agregadosCount;
            }
        }
    }
}