using OfficeOpenXml;
using ReporteVentas;
using ReporteVentas.Reporte;

namespace ReporteVentas.Tests;

public class ConsolidacionTests
{
    private static readonly string Entrada = Path.Combine(
        Rutas.BuscarCarpetaBase() ?? throw new DirectoryNotFoundException("No se encontró la carpeta 'entrada'."), "entrada");

    // Se calcula una sola vez: leer los archivos es lo más lento.
    private static readonly Lazy<Consolidado> Datos = new(() => Consolidador.Ejecutar(Entrada));

    [Theory]
    [InlineData("Centro", 691, 5_762_270)]
    [InlineData("Norte", 435, 3_482_100)]
    [InlineData("Rosario", 501, 3_949_680)]
    [InlineData("Córdoba", 356, 2_970_130)]
    public void Cada_sucursal_coincide_con_el_resultado_esperado(string sucursal, int operaciones, decimal ventas)
    {
        var deLaSucursal = Datos.Value.Ventas.Where(v => v.Sucursal == sucursal).ToList();
        Assert.Equal(operaciones, deLaSucursal.Count);
        Assert.Equal(ventas, deLaSucursal.Sum(v => v.Importe));
    }

    [Fact]
    public void El_total_coincide_con_el_resultado_esperado()
    {
        Assert.Equal(1_983, Datos.Value.Ventas.Count);
        Assert.Equal(16_164_180m, Datos.Value.Ventas.Sum(v => v.Importe));
    }

    [Fact]
    public void Solo_se_quitan_los_5_duplicados_de_Cordoba()
    {
        var cordoba = Datos.Value.Lecturas.Single(l => l.Sucursal == "Córdoba");
        Assert.Equal(361, cordoba.FilasLeidas);
        Assert.Equal(5, cordoba.DuplicadosQuitados);
        // Centro, Norte y Rosario tienen una fila repetida legítima cada una: no se toca.
        Assert.All(Datos.Value.Lecturas.Where(l => l.Sucursal != "Córdoba"), l => Assert.Equal(0, l.DuplicadosQuitados));
    }

    [Fact]
    public void Todas_las_ventas_quedan_completas_y_con_producto_del_catalogo()
    {
        Assert.Empty(Datos.Value.SinCatalogo);
        Assert.Empty(Datos.Value.NoReconocidos);
        Assert.All(Datos.Value.Ventas, v =>
        {
            Assert.False(string.IsNullOrWhiteSpace(v.Codigo));
            Assert.False(string.IsNullOrWhiteSpace(v.Categoria));
            Assert.Equal(v.Producto, v.Producto.Trim());
            Assert.Equal(v.Cantidad * v.PrecioUnitario, v.Importe);
        });
    }

    [Fact]
    public void Los_totales_declarados_por_Rosario_coinciden_con_lo_calculado()
    {
        var rosario = Datos.Value.Lecturas.Where(l => l.Sucursal == "Rosario").ToList();
        Assert.Equal(3, rosario.Count); // una hoja por mes
        Assert.All(rosario, l => Assert.Equal(l.TotalDeclarado, l.Ventas.Sum(v => v.Importe)));
    }

    [Fact]
    public void Normalizar_ignora_mayusculas_y_espacios_pero_no_tildes()
    {
        Assert.Equal("yerba mate 1kg", Catalogo.Normalizar("  YERBA   MATE 1KG "));
        Assert.NotEqual(Catalogo.Normalizar("Café"), Catalogo.Normalizar("Cafe"));
    }

    [Fact]
    public void El_reporte_se_genera_con_las_5_hojas_los_2_graficos_y_la_tabla()
    {
        var ruta = Path.Combine(Path.GetTempPath(), $"Reporte_{Guid.NewGuid():N}.xlsx");
        try
        {
            GeneradorReporte.Generar(Datos.Value, ruta);
            using var paquete = new ExcelPackage(new FileInfo(ruta));
            var libro = paquete.Workbook;
            Assert.Equal(
                ["Resumen", "Por sucursal y mes", "Productos", "Datos", "Control"],
                libro.Worksheets.Select(h => h.Name));
            Assert.Single(libro.Worksheets["Por sucursal y mes"].Drawings);
            Assert.Single(libro.Worksheets["Productos"].Drawings);
            Assert.Single(libro.Worksheets["Datos"].Tables);
            Assert.Equal(1_984, libro.Worksheets["Datos"].Dimension.End.Row); // encabezado + 1.983 filas
            Assert.Equal(16_164_180d, Convert.ToDouble(libro.Worksheets["Resumen"].Cells["B5"].Value));
        }
        finally { File.Delete(ruta); }
    }
}
