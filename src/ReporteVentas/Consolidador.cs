using ReporteVentas.Lectores;
using ReporteVentas.Modelo;

namespace ReporteVentas;

public class Consolidado
{
    public required List<Venta> Ventas { get; init; }
    public required List<LecturaArchivo> Lecturas { get; init; }
    public required List<string> NoReconocidos { get; init; }
    public required List<string> SinCatalogo { get; init; }
}

public static class Consolidador
{
    /// <summary>Sumar una sucursal nueva es agregar un lector acá.</summary>
    public static readonly IReadOnlyList<ILectorSucursal> Lectores =
        [new LectorCentro(), new LectorNorte(), new LectorRosario(), new LectorCordoba()];

    public const string ArchivoCatalogo = "Catalogo_Productos.xlsx";

    public static Consolidado Ejecutar(string carpetaEntrada, Action<string>? avance = null)
    {
        var catalogo = Catalogo.Cargar(Path.Combine(carpetaEntrada, ArchivoCatalogo));
        var lecturas = new List<LecturaArchivo>();
        var noReconocidos = new List<string>();

        var archivos = Directory.EnumerateFiles(carpetaEntrada)
            .Where(f => !Path.GetFileName(f).StartsWith("~$")) // temporales de Excel
            .Where(f => !Path.GetFileName(f).Equals(ArchivoCatalogo, StringComparison.OrdinalIgnoreCase))
            .Order(StringComparer.OrdinalIgnoreCase);

        foreach (var archivo in archivos)
        {
            var lector = Lectores.FirstOrDefault(l => l.PuedeLeer(archivo));
            if (lector is null) { noReconocidos.Add(Path.GetFileName(archivo)); continue; }

            var leidas = lector.Leer(archivo, catalogo);
            lecturas.AddRange(leidas);
            avance?.Invoke($"{lector.Sucursal,-8} {Path.GetFileName(archivo)}  ({leidas.Sum(l => l.Ventas.Count):N0} ventas)");
        }

        return new Consolidado
        {
            Ventas = lecturas.SelectMany(l => l.Ventas).OrderBy(v => v.Fecha).ThenBy(v => v.Sucursal).ToList(),
            Lecturas = lecturas,
            NoReconocidos = noReconocidos,
            SinCatalogo = lecturas.SelectMany(l => l.SinCatalogo).Distinct().Order().ToList(),
        };
    }
}
