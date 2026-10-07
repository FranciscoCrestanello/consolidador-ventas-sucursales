using System.Diagnostics;
using ReporteVentas;
using ReporteVentas.Reporte;

Console.OutputEncoding = System.Text.Encoding.UTF8;
var abrir = !args.Contains("--no-abrir");
var reloj = Stopwatch.StartNew();

Console.WriteLine("=== Distribuidora Del Sur · Reporte de ventas consolidado ===\n");
try
{
    var carpetaBase = args.FirstOrDefault(a => !a.StartsWith("--")) ?? Rutas.BuscarCarpetaBase()
        ?? throw new DirectoryNotFoundException("No encontré la carpeta 'entrada' junto al programa.");
    var entrada = Path.Combine(carpetaBase, "entrada");
    var salida = Path.Combine(carpetaBase, "salida");
    Directory.CreateDirectory(salida);
    var rutaReporte = Path.Combine(salida, "Reporte_Ventas.xlsx");

    Console.WriteLine($"Leyendo archivos de: {entrada}\n");
    var datos = Consolidador.Ejecutar(entrada, linea => Console.WriteLine("  " + linea));

    Console.WriteLine("\nGenerando el reporte...");
    GeneradorReporte.Generar(datos, rutaReporte);

    var duplicados = datos.Lecturas.Sum(l => l.DuplicadosQuitados);
    Console.WriteLine($"\n✓ {datos.Ventas.Count:N0} operaciones · $ {datos.Ventas.Sum(v => v.Importe):N0} · {duplicados} duplicadas quitadas · {datos.SinCatalogo.Count} productos sin catálogo");
    foreach (var a in datos.NoReconocidos) Console.WriteLine($"  ! Archivo no reconocido (se ignoró): {a}");
    Console.WriteLine($"✓ Listo en {reloj.Elapsed.TotalSeconds:N1} segundos: {rutaReporte}");

    if (abrir)
        Process.Start(new ProcessStartInfo(rutaReporte) { UseShellExecute = true });
    return 0;
}
catch (IOException ex) when ((ex.HResult & 0xFFFF) is 32 or 33) // archivo en uso por otro proceso (Excel)
{
    Console.WriteLine("\n✗ No pude guardar Reporte_Ventas.xlsx: está abierto en Excel. Cerralo y volvé a ejecutar.");
    return Salir(1, abrir);
}
catch (Exception ex)
{
    Console.WriteLine($"\n✗ Error: {ex.Message}");
    return Salir(1, abrir);
}

static int Salir(int codigo, bool interactivo)
{
    if (interactivo && !Console.IsInputRedirected)
    {
        Console.WriteLine("\nPresioná una tecla para cerrar...");
        Console.ReadKey(true);
    }
    return codigo;
}
