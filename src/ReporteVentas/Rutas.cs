namespace ReporteVentas;

public static class Rutas
{
    /// <summary>
    /// Busca la carpeta que contiene "entrada": primero junto al .exe y hacia arriba (modo desarrollo),
    /// después desde la carpeta de trabajo.
    /// </summary>
    public static string? BuscarCarpetaBase()
    {
        foreach (var inicio in new[] { AppContext.BaseDirectory, Directory.GetCurrentDirectory() })
            for (var dir = new DirectoryInfo(inicio); dir is not null; dir = dir.Parent)
                if (Directory.Exists(Path.Combine(dir.FullName, "entrada")))
                    return dir.FullName;
        return null;
    }
}
