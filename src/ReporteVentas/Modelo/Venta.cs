namespace ReporteVentas.Modelo;

public record Venta(
    DateTime Fecha,
    string Sucursal,
    string Codigo,
    string Producto,
    string Categoria,
    int Cantidad,
    decimal PrecioUnitario,
    decimal Importe,
    string Vendedor);
