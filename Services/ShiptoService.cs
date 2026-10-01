using System;
using System.Collections.Generic;
using System.Data;
using Microsoft.Data.SqlClient;
using ProcesarFacturas.Models;

namespace ProcesarFacturas.Services
{
    public class ShiptoService
    {
        private const string ConnectionString =
            "Server=Icoldata01n;Database=dbtraductorQ;Integrated Security=True;TrustServerCertificate=True;";

        // 1. OBTENER SHIPTOS EXISTENTES Y RETORNARLOS EN UN HASHSET
        public HashSet<string> ObtenerCodigosExistentes()
        {
            var codigos = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            try
            {
                using var conexion = new SqlConnection(ConnectionString);
                conexion.Open();

                const string query = "SELECT ShipTo FROM [dbo].[Shipto_Procter] WITH (NOLOCK)";

                using var comando = new SqlCommand(query, conexion);
                using var reader = comando.ExecuteReader();

                while (reader.Read())
                {
                    // Se usa ?? string.Empty para evitar la advertencia CS8600
                    string shipTo = reader["ShipTo"]?.ToString()?.Trim() ?? string.Empty;
                    if (!string.IsNullOrEmpty(shipTo))
                    {
                        codigos.Add(shipTo);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[ERROR SQL] Error al cargar catálogo de Shiptos: {ex.Message}");
            }

            return codigos;
        }

        // 2. INSERCIÓN MASIVA EN LOTE (BULK INSERT)
        public int GuardarNuevosShiptosMasivo(List<RegistroFactura> nuevosRegistros)
        {
            if (nuevosRegistros == null || nuevosRegistros.Count == 0) return 0;

            try
            {
                var tablaDatos = new DataTable();
                tablaDatos.Columns.Add("ShipTo", typeof(string));
                tablaDatos.Columns.Add("Name1", typeof(string));
                tablaDatos.Columns.Add("Street", typeof(string));
                tablaDatos.Columns.Add("City", typeof(string));

                foreach (var reg in nuevosRegistros)
                {
                    tablaDatos.Rows.Add(
                        reg.ShipToParty ?? (object)DBNull.Value,
                        reg.Cliente ?? (object)DBNull.Value,
                        reg.Direccion ?? (object)DBNull.Value,
                        reg.Ciudad ?? (object)DBNull.Value
                    );
                }

                using var conexion = new SqlConnection(ConnectionString);
                conexion.Open();

                using var bulkCopy = new SqlBulkCopy(conexion)
                {
                    DestinationTableName = "[dbo].[Shipto_Procter]",
                    BatchSize = 5000,
                    BulkCopyTimeout = 60
                };

                bulkCopy.ColumnMappings.Add("ShipTo", "ShipTo");
                bulkCopy.ColumnMappings.Add("Name1", "Name1");
                bulkCopy.ColumnMappings.Add("Street", "Street");
                bulkCopy.ColumnMappings.Add("City", "City");

                bulkCopy.WriteToServer(tablaDatos);
                return nuevosRegistros.Count;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[ERROR SQL BULK] Fallo en inserción masiva: {ex.Message}");
                return 0;
            }
        }
    }
}