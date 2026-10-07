using ReporteVentas.Modelo;

namespace ReporteVentas.Lectores;

public interface ILectorSucursal
{
    string Sucursal { get; }
    bool PuedeLeer(string rutaArchivo);
    /// <summary>Devuelve una lectura por archivo (o por hoja, si el archivo trae una hoja por mes).</summary>
    IReadOnlyList<LecturaArchivo> Leer(string rutaArchivo, Catalogo catalogo);
}
