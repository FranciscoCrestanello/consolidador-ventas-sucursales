namespace ReporteVentas.Modelo;

/// <summary>Resultado de leer un archivo (o una hoja), con los datos para la hoja Control.</summary>
public class LecturaArchivo
{
    public required string Archivo { get; init; }
    public required string Sucursal { get; init; }
    public List<Venta> Ventas { get; } = [];
    public List<string> SinCatalogo { get; } = [];
    /// <summary>Filas de datos procesadas, antes de quitar duplicados.</summary>
    public int FilasLeidas { get; set; }
    /// <summary>Títulos, filas vacías y filas de total que no se suman.</summary>
    public int FilasSalteadas { get; set; }
    public int DuplicadosQuitados { get; set; }
    /// <summary>Suma de las filas TOTAL del archivo, si las trae (para cruzar contra lo calculado).</summary>
    public decimal? TotalDeclarado { get; set; }
}
