using System.Data;
using System.Drawing;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table;
using ReporteVentas.Modelo;

namespace ReporteVentas.Reporte;

/// <summary>
/// Genera Reporte_Ventas.xlsx. Se escribe todo con EPPlus 4.5.3 (última versión gratuita), que además de tablas
/// y estilos arma los gráficos nativos; ClosedXML solo se usa para leer. Mezclar ambos para escribir no funciona:
/// EPPlus no entiende el XML con prefijo "x:" que genera ClosedXML y deja la hoja corrupta.
/// </summary>
public static class GeneradorReporte
{
    public const string HResumen = "Resumen", HSucMes = "Por sucursal y mes", HProductos = "Productos",
        HDatos = "Datos", HControl = "Control";

    private static readonly Color Azul = ColorTranslator.FromHtml("#1F4E78"), AzulClaro = ColorTranslator.FromHtml("#DDEBF7"),
        Verde = ColorTranslator.FromHtml("#2E7D32");
    private const string FmtMoneda = "$ #,##0", FmtEntero = "#,##0";
    private static readonly string[] Meses =
        ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"];

    public static string NombreMes(DateTime f) => char.ToUpper(Meses[f.Month - 1][0]) + Meses[f.Month - 1][1..];

    public static void Generar(Consolidado datos, string rutaSalida)
    {
        var ventas = datos.Ventas;
        if (File.Exists(rutaSalida)) File.Delete(rutaSalida); // falla con IOException si está abierto en Excel

        using var paquete = new ExcelPackage();
        HojaResumen(paquete.Workbook.Worksheets.Add(HResumen), ventas);
        HojaSucursalMes(paquete.Workbook.Worksheets.Add(HSucMes), ventas);
        HojaProductos(paquete.Workbook.Worksheets.Add(HProductos), ventas);
        HojaDatos(paquete.Workbook.Worksheets.Add(HDatos), ventas);
        HojaControl(paquete.Workbook.Worksheets.Add(HControl), datos);

        paquete.Workbook.Calculate(); // guarda el valor de las fórmulas, para que cualquier visor las muestre
        paquete.SaveAs(new FileInfo(rutaSalida));
    }

    // ---------- Resumen ----------

    private static void HojaResumen(ExcelWorksheet ws, List<Venta> ventas)
    {
        ws.View.ShowGridLines = false;
        Titulo(ws, "Distribuidora Del Sur — Reporte de ventas consolidado");

        var desde = ventas.Min(v => v.Fecha); var hasta = ventas.Max(v => v.Fecha);
        ws.Cells["A2"].Value = $"Período: {desde:dd/MM/yyyy} al {hasta:dd/MM/yyyy}";
        ws.Cells["A2"].Style.Font.Bold = true; ws.Cells["A2"].Style.Font.Size = 12;
        ws.Cells["A3"].Value = $"{ventas.Select(v => v.Sucursal).Distinct().Count()} sucursales · generado el {DateTime.Now:dd/MM/yyyy HH:mm}";
        ws.Cells["A3"].Style.Font.Color.SetColor(Color.Gray);

        var total = ventas.Sum(v => v.Importe);
        var mejorSuc = ventas.GroupBy(v => v.Sucursal).Select(g => (g.Key, Total: g.Sum(v => v.Importe))).MaxBy(x => x.Total);
        var mejorVend = ventas.GroupBy(v => (v.Vendedor, v.Sucursal)).Select(g => (g.Key, Total: g.Sum(v => v.Importe))).MaxBy(x => x.Total);
        var topProd = ventas.GroupBy(v => v.Producto).Select(g => (Nombre: g.Key, Unidades: g.Sum(v => v.Cantidad))).MaxBy(x => x.Unidades);

        var filas = new (string Etiqueta, object Valor, string? Formato, string Detalle)[]
        {
            ("Ventas totales", total, FmtMoneda, ""),
            ("Cantidad de operaciones", ventas.Count, FmtEntero, ""),
            ("Ticket promedio", total / ventas.Count, FmtMoneda, "ventas totales / operaciones"),
            ("Mejor sucursal", mejorSuc.Key, null, $"$ {mejorSuc.Total:N0}"),
            ("Mejor vendedor", mejorVend.Key.Vendedor, null, $"{mejorVend.Key.Sucursal} · $ {mejorVend.Total:N0}"),
            ("Producto más vendido", topProd.Nombre, null, $"{topProd.Unidades:N0} unidades"),
        };
        int fila = 5;
        foreach (var (etiqueta, valor, formato, detalle) in filas)
        {
            var e = ws.Cells[fila, 1]; e.Value = etiqueta; e.Style.Font.Bold = true; Rellenar(e, AzulClaro);
            var v = ws.Cells[fila, 2]; v.Value = valor; v.Style.Font.Bold = true; v.Style.Font.Size = 13;
            v.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
            if (formato is not null) v.Style.Numberformat.Format = formato;
            var d = ws.Cells[fila, 3]; d.Value = detalle; d.Style.Font.Color.SetColor(Color.Gray);
            ws.Cells[fila, 1, fila, 2].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            ws.Row(fila).Height = 22;
            ws.Cells[fila, 1, fila, 3].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            fila++;
        }
        ws.Column(1).Width = 28; ws.Column(2).Width = 26; ws.Column(3).Width = 36;
    }

    // ---------- Por sucursal y mes ----------

    private static void HojaSucursalMes(ExcelWorksheet ws, List<Venta> ventas)
    {
        ws.View.ShowGridLines = false;
        Titulo(ws, "Ventas por sucursal y mes (AR$)");

        var meses = ventas.Select(v => new DateTime(v.Fecha.Year, v.Fecha.Month, 1)).Distinct().Order().ToList();
        var sucursales = ventas.GroupBy(v => v.Sucursal).OrderByDescending(g => g.Sum(v => v.Importe)).Select(g => g.Key).ToList();

        const int enc = 3;
        ws.Cells[enc, 1].Value = "Sucursal";
        for (int m = 0; m < meses.Count; m++) ws.Cells[enc, 2 + m].Value = NombreMes(meses[m]);
        int colTotal = 2 + meses.Count;
        ws.Cells[enc, colTotal].Value = "Total";
        Encabezado(ws.Cells[enc, 1, enc, colTotal]);

        int fila = enc + 1;
        foreach (var suc in sucursales)
        {
            ws.Cells[fila, 1].Value = suc;
            for (int m = 0; m < meses.Count; m++)
                ws.Cells[fila, 2 + m].Value = ventas
                    .Where(v => v.Sucursal == suc && v.Fecha.Year == meses[m].Year && v.Fecha.Month == meses[m].Month).Sum(v => v.Importe);
            ws.Cells[fila, colTotal].Formula = $"SUM({ws.Cells[fila, 2].Address}:{ws.Cells[fila, colTotal - 1].Address})";
            fila++;
        }
        int ult = fila - 1;
        ws.Cells[fila, 1].Value = "Total";
        for (int c = 2; c <= colTotal; c++)
            ws.Cells[fila, c].Formula = $"SUM({ws.Cells[enc + 1, c].Address}:{ws.Cells[ult, c].Address})";
        var totales = ws.Cells[fila, 1, fila, colTotal];
        totales.Style.Font.Bold = true; Rellenar(totales, AzulClaro);
        totales.Style.Border.Top.Style = ExcelBorderStyle.Thin;

        ws.Cells[enc + 1, 2, fila, colTotal].Style.Numberformat.Format = FmtMoneda;
        ws.Cells[enc + 1, colTotal, ult, colTotal].Style.Font.Bold = true;
        for (int c = 1; c <= colTotal; c++) ws.Column(c).Width = 18;

        // Columnas agrupadas: una serie por mes, una categoría por sucursal.
        var grafico = (ExcelBarChart)ws.Drawings.AddChart("GraficoSucursalMes", eChartType.ColumnClustered);
        grafico.Title.Text = "Ventas por sucursal y mes (AR$)";
        var categorias = ws.Cells[enc + 1, 1, ult, 1];
        for (int c = 2; c < colTotal; c++)
        {
            var serie = grafico.Series.Add(ws.Cells[enc + 1, c, ult, c], categorias);
            serie.HeaderAddress = ws.Cells[enc, c];
        }
        grafico.YAxis.Format = "$ #,##0";
        grafico.Legend.Position = eLegendPosition.Bottom;
        grafico.SetPosition(fila + 1, 0, 0, 0);
        grafico.SetSize(720, 380);
    }

    // ---------- Productos ----------

    private static void HojaProductos(ExcelWorksheet ws, List<Venta> ventas)
    {
        ws.View.ShowGridLines = false;
        Titulo(ws, "Productos y categorías");

        // Top 10 por importe (A:E)
        ws.Cells["A2"].Value = "Top 10 productos por importe"; ws.Cells["A2"].Style.Font.Bold = true;
        string[] cab = ["#", "Producto", "Categoría", "Unidades", "Importe"];
        for (int i = 0; i < cab.Length; i++) ws.Cells[3, 1 + i].Value = cab[i];
        Encabezado(ws.Cells[3, 1, 3, 5]);
        var top = ventas.GroupBy(v => (v.Producto, v.Categoria))
            .Select(g => (g.Key.Producto, g.Key.Categoria, Unidades: g.Sum(v => v.Cantidad), Importe: g.Sum(v => v.Importe)))
            .OrderByDescending(x => x.Importe).Take(10).ToList();
        for (int i = 0; i < top.Count; i++)
        {
            int f = 4 + i;
            ws.Cells[f, 1].Value = i + 1; ws.Cells[f, 2].Value = top[i].Producto; ws.Cells[f, 3].Value = top[i].Categoria;
            ws.Cells[f, 4].Value = top[i].Unidades; ws.Cells[f, 5].Value = top[i].Importe;
        }
        ws.Cells[4, 4, 3 + top.Count, 4].Style.Numberformat.Format = FmtEntero;
        ws.Cells[4, 5, 3 + top.Count, 5].Style.Numberformat.Format = FmtMoneda;

        // Ventas por categoría (G:I)
        ws.Cells["G2"].Value = "Ventas por categoría"; ws.Cells["G2"].Style.Font.Bold = true;
        string[] cab2 = ["Categoría", "Unidades", "Importe"];
        for (int i = 0; i < cab2.Length; i++) ws.Cells[3, 7 + i].Value = cab2[i];
        Encabezado(ws.Cells[3, 7, 3, 9]);
        var cats = ventas.GroupBy(v => v.Categoria)
            .Select(g => (Categoria: g.Key, Unidades: g.Sum(v => v.Cantidad), Importe: g.Sum(v => v.Importe)))
            .OrderByDescending(x => x.Importe).ToList();
        for (int i = 0; i < cats.Count; i++)
        {
            int f = 4 + i;
            ws.Cells[f, 7].Value = cats[i].Categoria; ws.Cells[f, 8].Value = cats[i].Unidades; ws.Cells[f, 9].Value = cats[i].Importe;
        }
        int ultCat = 3 + cats.Count;
        ws.Cells[4, 8, ultCat, 8].Style.Numberformat.Format = FmtEntero;
        ws.Cells[4, 9, ultCat, 9].Style.Numberformat.Format = FmtMoneda;

        ws.Column(1).Width = 5; ws.Column(2).Width = 28; ws.Column(3).Width = 14; ws.Column(4).Width = 11; ws.Column(5).Width = 16;
        ws.Column(6).Width = 4; ws.Column(7).Width = 16; ws.Column(8).Width = 11; ws.Column(9).Width = 16;

        var grafico = (ExcelBarChart)ws.Drawings.AddChart("GraficoCategorias", eChartType.ColumnClustered);
        grafico.Title.Text = "Ventas por categoría (AR$)";
        var serie = grafico.Series.Add(ws.Cells[4, 9, ultCat, 9], ws.Cells[4, 7, ultCat, 7]);
        serie.HeaderAddress = ws.Cells[3, 9];
        grafico.YAxis.Format = "$ #,##0";
        grafico.Legend.Remove();
        grafico.SetPosition(Math.Max(ultCat, 13) + 1, 0, 0, 0); // debajo de las dos tablas
        grafico.SetSize(620, 340);
    }

    // ---------- Datos ----------

    private static void HojaDatos(ExcelWorksheet ws, List<Venta> ventas)
    {
        string[] cab = ["Fecha", "Sucursal", "Código", "Producto", "Categoría", "Cantidad", "Precio unitario", "Importe", "Vendedor"];
        for (int i = 0; i < cab.Length; i++) ws.Cells[1, 1 + i].Value = cab[i];
        for (int i = 0; i < ventas.Count; i++)
        {
            var v = ventas[i]; int f = 2 + i;
            ws.Cells[f, 1].Value = v.Fecha; ws.Cells[f, 2].Value = v.Sucursal; ws.Cells[f, 3].Value = v.Codigo;
            ws.Cells[f, 4].Value = v.Producto; ws.Cells[f, 5].Value = v.Categoria; ws.Cells[f, 6].Value = v.Cantidad;
            ws.Cells[f, 7].Value = v.PrecioUnitario; ws.Cells[f, 8].Value = v.Importe; ws.Cells[f, 9].Value = v.Vendedor;
        }
        int ult = ventas.Count + 1;
        ws.Cells[2, 1, ult, 1].Style.Numberformat.Format = "dd/mm/yyyy";
        ws.Cells[2, 7, ult, 8].Style.Numberformat.Format = FmtMoneda;

        var tabla = ws.Tables.Add(ws.Cells[1, 1, ult, cab.Length], "TablaVentas"); // formato de tabla + filtros
        tabla.TableStyle = TableStyles.Medium2;
        ws.View.FreezePanes(2, 1);
        double[] anchos = [12, 11, 9, 24, 12, 10, 16, 14, 18];
        for (int i = 0; i < anchos.Length; i++) ws.Column(1 + i).Width = anchos[i];
    }

    // ---------- Control ----------

    private static void HojaControl(ExcelWorksheet ws, Consolidado datos)
    {
        ws.View.ShowGridLines = false;
        Titulo(ws, "Control: qué encontró el programa");

        // 1) Archivos leídos
        ws.Cells["A3"].Value = "Archivos leídos"; ws.Cells["A3"].Style.Font.Bold = true;
        string[] cab = ["Archivo", "Sucursal", "Filas leídas", "Filas salteadas*", "Duplicados quitados", "Operaciones", "Importe (AR$)"];
        for (int i = 0; i < cab.Length; i++) ws.Cells[4, 1 + i].Value = cab[i];
        Encabezado(ws.Cells[4, 1, 4, cab.Length]);
        int f = 5;
        foreach (var l in datos.Lecturas)
        {
            ws.Cells[f, 1].Value = l.Archivo; ws.Cells[f, 2].Value = l.Sucursal; ws.Cells[f, 3].Value = l.FilasLeidas;
            ws.Cells[f, 4].Value = l.FilasSalteadas; ws.Cells[f, 5].Value = l.DuplicadosQuitados;
            ws.Cells[f, 6].Value = l.Ventas.Count; ws.Cells[f, 7].Value = l.Ventas.Sum(v => v.Importe);
            f++;
        }
        ws.Cells[f, 1].Value = "Total";
        for (int c = 3; c <= 7; c++) ws.Cells[f, c].Formula = $"SUM({ws.Cells[5, c].Address}:{ws.Cells[f - 1, c].Address})";
        var tot = ws.Cells[f, 1, f, 7]; tot.Style.Font.Bold = true; Rellenar(tot, AzulClaro);
        ws.Cells[5, 3, f, 6].Style.Numberformat.Format = FmtEntero;
        ws.Cells[5, 7, f, 7].Style.Numberformat.Format = FmtMoneda;
        f++;
        ws.Cells[f, 1].Value = "* Títulos, encabezados, filas vacías y filas TOTAL que no se suman.";
        ws.Cells[f, 1].Style.Font.Color.SetColor(Color.Gray); ws.Cells[f, 1].Style.Font.Italic = true;
        f += 2;

        // 2) Verificación contra los TOTAL que traen los archivos
        var conTotal = datos.Lecturas.Where(l => l.TotalDeclarado is not null).ToList();
        if (conTotal.Count > 0)
        {
            ws.Cells[f, 1].Value = "Verificación contra los TOTAL que traen los archivos"; ws.Cells[f, 1].Style.Font.Bold = true; f++;
            string[] cab2 = ["Archivo", "Total declarado", "Total calculado", "Resultado"];
            for (int i = 0; i < cab2.Length; i++) ws.Cells[f, 1 + i].Value = cab2[i];
            Encabezado(ws.Cells[f, 1, f, cab2.Length]); f++;
            foreach (var l in conTotal)
            {
                var calculado = l.Ventas.Sum(v => v.Importe);
                ws.Cells[f, 1].Value = l.Archivo; ws.Cells[f, 2].Value = l.TotalDeclarado!.Value; ws.Cells[f, 3].Value = calculado;
                ws.Cells[f, 2, f, 3].Style.Numberformat.Format = FmtMoneda;
                var ok = l.TotalDeclarado == calculado;
                ws.Cells[f, 4].Value = ok ? "✓ Coincide" : "✗ No coincide";
                ws.Cells[f, 4].Style.Font.Color.SetColor(ok ? Verde : Color.Red); ws.Cells[f, 4].Style.Font.Bold = true;
                f++;
            }
            f++;
        }

        // 3) Catálogo
        ws.Cells[f, 1].Value = "Productos que no se encontraron en el catálogo"; ws.Cells[f, 1].Style.Font.Bold = true; f++;
        if (datos.SinCatalogo.Count == 0)
        {
            ws.Cells[f, 1].Value = "✓ Todos los productos se encontraron en el catálogo (0 sin catálogo)";
            ws.Cells[f, 1].Style.Font.Color.SetColor(Verde); ws.Cells[f, 1].Style.Font.Bold = true; f++;
        }
        else
            foreach (var p in datos.SinCatalogo) { ws.Cells[f, 1].Value = p; ws.Cells[f, 1].Style.Font.Color.SetColor(Color.Red); f++; }
        f++;

        // 4) Archivos sin lector
        ws.Cells[f, 1].Value = "Archivos de la carpeta que no se reconocieron"; ws.Cells[f, 1].Style.Font.Bold = true; f++;
        if (datos.NoReconocidos.Count == 0)
        {
            ws.Cells[f, 1].Value = "✓ Todos los archivos se reconocieron";
            ws.Cells[f, 1].Style.Font.Color.SetColor(Verde); ws.Cells[f, 1].Style.Font.Bold = true;
        }
        else
            foreach (var a in datos.NoReconocidos) { ws.Cells[f, 1].Value = a; ws.Cells[f, 1].Style.Font.Color.SetColor(Color.Red); f++; }

        ws.Column(1).Width = 52; ws.Column(2).Width = 16;
        for (int c = 3; c <= 7; c++) ws.Column(c).Width = 18;
    }

    // ---------- Estilos ----------

    private static void Titulo(ExcelWorksheet ws, string texto)
    {
        ws.Cells["A1"].Value = texto;
        ws.Cells["A1"].Style.Font.Bold = true; ws.Cells["A1"].Style.Font.Size = 16; ws.Cells["A1"].Style.Font.Color.SetColor(Azul);
    }

    private static void Rellenar(ExcelRange rango, Color color)
    {
        rango.Style.Fill.PatternType = ExcelFillStyle.Solid;
        rango.Style.Fill.BackgroundColor.SetColor(color);
    }

    private static void Encabezado(ExcelRange rango)
    {
        rango.Style.Font.Bold = true; rango.Style.Font.Color.SetColor(Color.White); Rellenar(rango, Azul);
        rango.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
    }
}
