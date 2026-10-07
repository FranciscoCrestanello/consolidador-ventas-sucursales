using ReporteVentas.Modelo;

namespace ReporteVentas.Lectores;

/// <summary>Arma la <see cref="Venta"/> unificada: completa con el catálogo y calcula precio/importe faltantes.</summary>
public static class ConstructorVenta
{
    public const string SinCategoria = "Sin catálogo";

    public static void Agregar(LecturaArchivo lectura, Catalogo catalogo, string sucursal, DateTime fecha,
        string? codigo, string producto, string? categoria, int cantidad,
        decimal? precioUnitario, decimal? importe, string vendedor)
    {
        var enCatalogo = (string.IsNullOrWhiteSpace(codigo) ? null : catalogo.BuscarPorCodigo(codigo))
                         ?? catalogo.BuscarPorNombre(producto);
        if (enCatalogo is null) lectura.SinCatalogo.Add(producto.Trim());

        // Último recurso si la fila no trae ni precio ni importe: el precio de lista del catálogo.
        precioUnitario ??= importe is null ? enCatalogo?.PrecioLista : importe / cantidad;
        importe ??= cantidad * precioUnitario;
        if (importe is null || precioUnitario is null)
            throw new InvalidDataException($"No se pudo determinar el importe de '{producto}' ({fecha:dd/MM/yyyy}).");

        lectura.Ventas.Add(new Venta(
            fecha, sucursal,
            enCatalogo?.Codigo ?? codigo?.Trim() ?? "",
            enCatalogo?.Nombre ?? producto.Trim(),
            enCatalogo?.Categoria ?? (string.IsNullOrWhiteSpace(categoria) ? SinCategoria : categoria.Trim()),
            cantidad, precioUnitario.Value, importe.Value, vendedor.Trim()));
    }
}
