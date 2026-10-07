using ClosedXML.Excel;

namespace ReporteVentas;

public record ProductoCatalogo(string Codigo, string Nombre, string Categoria, decimal PrecioLista);

public class Catalogo
{
    private readonly Dictionary<string, ProductoCatalogo> _porNombre;
    private readonly Dictionary<string, ProductoCatalogo> _porCodigo;

    private Catalogo(IReadOnlyList<ProductoCatalogo> productos)
    {
        _porNombre = productos.ToDictionary(p => Normalizar(p.Nombre));
        _porCodigo = productos.ToDictionary(p => p.Codigo.Trim().ToUpperInvariant());
    }

    /// <summary>Quita espacios de más y pasa a minúsculas; conserva las tildes.</summary>
    public static string Normalizar(string texto) =>
        string.Join(' ', texto.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries)).ToLowerInvariant();

    public static Catalogo Cargar(string ruta)
    {
        using var wb = new XLWorkbook(ruta);
        var productos = wb.Worksheets.First().RowsUsed().Skip(1)
            .Select(r => new ProductoCatalogo(
                r.Cell(1).GetString().Trim(),
                r.Cell(2).GetString().Trim(),
                r.Cell(3).GetString().Trim(),
                r.Cell(4).GetValue<decimal>()))
            .ToList();
        return new Catalogo(productos);
    }

    public ProductoCatalogo? BuscarPorNombre(string nombre) =>
        _porNombre.GetValueOrDefault(Normalizar(nombre));

    public ProductoCatalogo? BuscarPorCodigo(string codigo) =>
        _porCodigo.GetValueOrDefault(codigo.Trim().ToUpperInvariant());
}
