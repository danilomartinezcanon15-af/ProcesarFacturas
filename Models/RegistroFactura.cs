namespace ProcesarFacturas.Models
{
    public class RegistroFactura
    {
        public string ShipToParty { get; set; } = string.Empty;
        public string Cliente { get; set; } = string.Empty;
        public string Direccion { get; set; } = string.Empty;
        public string Ciudad { get; set; } = string.Empty;

        public Dictionary<string, string> Datos { get; set; } = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        public string ObtenerValor(string clave)
        {
            if (Datos != null && Datos.TryGetValue(clave, out string? valor) && !string.IsNullOrEmpty(valor))
            {
                return valor;
            }

            return clave.ToLower() switch
            {
                "ship-to" or "shiptoparty" => ShipToParty,
                "name 1" or "cliente" => Cliente,
                "street" or "direccion" => Direccion,
                "city" or "ciudad" => Ciudad,
                _ => string.Empty
            };
        }
    }
}