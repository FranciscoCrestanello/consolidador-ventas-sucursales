using System.Globalization;
using System.Text;
using ClosedXML.Excel;

namespace ReporteVentas.Lectores;

public static class ExcelUtil
{
    private static readonly string[] FormatosFecha = ["dd/MM/yyyy", "d/M/yyyy", "yyyy-MM-dd"];

    /// <summary>Clave de encabezado: minúsculas, sin tildes ni signos ("Cód." → "cod").</summary>
    public static string Clave(string texto)
    {
        var sb = new StringBuilder();
        foreach (var c in texto.Normalize(NormalizationForm.FormD))
        {
            if (CharUnicodeInfo.GetUnicodeCategory(c) == UnicodeCategory.NonSpacingMark) continue;
            if (char.IsLetterOrDigit(c)) sb.Append(char.ToLowerInvariant(c));
        }
        return sb.ToString();
    }

    public static Dictionary<string, int> Encabezados(IXLRow fila)
    {
        var mapa = new Dictionary<string, int>();
        foreach (var celda in fila.CellsUsed())
            mapa.TryAdd(Clave(celda.GetString()), celda.Address.ColumnNumber);
        return mapa;
    }

    /// <summary>Busca una columna por cualquiera de sus nombres posibles; 0 si no está y es opcional.</summary>
    public static int Col(Dictionary<string, int> encabezados, bool opcional, params string[] alias)
    {
        foreach (var a in alias)
            if (encabezados.TryGetValue(a, out var idx)) return idx;
        return opcional
            ? 0
            : throw new InvalidDataException($"No se encontró la columna '{alias[0]}'. Encabezados: {string.Join(", ", encabezados.Keys)}");
    }

    public static DateTime Fecha(IXLCell celda)
    {
        if (celda.DataType == XLDataType.DateTime) return celda.GetDateTime().Date;
        return ParsearFecha(celda.GetString());
    }

    public static DateTime ParsearFecha(string texto) =>
        DateTime.ParseExact(texto.Trim(), FormatosFecha, CultureInfo.InvariantCulture, DateTimeStyles.None);

    public static decimal Numero(IXLCell celda) =>
        celda.DataType == XLDataType.Number
            ? celda.GetValue<decimal>()
            : decimal.Parse(celda.GetString().Trim().Replace(".", "").Replace(',', '.'), CultureInfo.InvariantCulture);

    public static bool TieneTexto(IXLCell celda, string texto) =>
        string.Equals(celda.GetString().Trim(), texto, StringComparison.OrdinalIgnoreCase);
}
