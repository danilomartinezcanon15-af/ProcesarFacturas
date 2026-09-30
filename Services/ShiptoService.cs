using System;
using System.Collections.Generic;
using Microsoft.Data.SqlClient;
using ProcesarFacturas.Models;

namespace ProcesarFacturas.Services
{
    public class ShiptoService
    {
        // Cadena de conexión corporativa a la base de datos
        private const string ConnectionString =
            "Server=Icoldata01n;Database=dbtraductorQ;Integrated Security=True;TrustServerCertificate=True;";

        // 1. OBTENER SHIPTOS EXISTENTES EN SQL SERVER
        public List<RegistroFactura> ObtenerShiptos()
        {
            var lista = new List<RegistroFactura>();

            try
            {
                using var conexion = new SqlConnection(ConnectionString);
                conexion.Open();

                const string query = @"
                    SELECT ShipTo, Name1, Street, City
                    FROM [dbo].[Shipto_Procter]";

                using var comando = new SqlCommand(query, conexion);
                using var reader = comando.ExecuteReader();

                while (reader.Read())
                {
                    lista.Add(new RegistroFactura
                    {
                        ShipToParty = reader["ShipTo"]?.ToString()?.Trim() ?? "",
                        Cliente     = reader["Name1"]?.ToString()?.Trim() ?? "",
                        Direccion   = reader["Street"]?.ToString()?.Trim() ?? "",
                        Ciudad      = reader["City"]?.ToString()?.Trim()?.ToUpper() ?? ""
                    });
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[ERROR SQL] Error al consultar Shipto_Procter: {ex.Message}");
            }

            return lista;
        }

        // 2. INSERTAR UN NUEVO SHIPTO (Solo si no existe)
        public bool GuardarNuevoShipto(RegistroFactura registro)
        {
            if (string.IsNullOrWhiteSpace(registro.ShipToParty)) return false;

            try
            {
                using var conexion = new SqlConnection(ConnectionString);
                conexion.Open();

                // Validar duplicado antes de insertar
                const string checkQuery = @"
                    SELECT COUNT(1) 
                    FROM [dbo].[Shipto_Procter] 
                    WHERE ShipTo = @ShipTo";

                using (var checkCmd = new SqlCommand(checkQuery, conexion))
                {
                    checkCmd.Parameters.AddWithValue("@ShipTo", registro.ShipToParty);
                    if (Convert.ToInt32(checkCmd.ExecuteScalar()) > 0) return false;
                }

                // Inserción segura y parametrizada
                const string insertQuery = @"
                    INSERT INTO [dbo].[Shipto_Procter] (ShipTo, Name1, Street, City)
                    VALUES (@ShipTo, @Name1, @Street, @City)";

                using var insertCmd = new SqlCommand(insertQuery, conexion);
                insertCmd.Parameters.AddWithValue("@ShipTo", registro.ShipToParty ?? (object)DBNull.Value);
                insertCmd.Parameters.AddWithValue("@Name1", registro.Cliente ?? (object)DBNull.Value);
                insertCmd.Parameters.AddWithValue("@Street", registro.Direccion ?? (object)DBNull.Value);
                insertCmd.Parameters.AddWithValue("@City", registro.Ciudad ?? (object)DBNull.Value);

                return insertCmd.ExecuteNonQuery() > 0;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[ERROR SQL] No se pudo guardar el ShipTo '{registro.ShipToParty}': {ex.Message}");
                return false;
            }
        }
    }
}