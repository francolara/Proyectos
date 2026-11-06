using System.Text.Json;
using System.Text.RegularExpressions;
using Microsoft.SqlServer.TransactSql.ScriptDom;

// Valida archivos y sintaxis offline. No contiene SqlConnection ni ejecuta SQL.
var root = Path.GetFullPath(args.Length == 1 ? args[0] : "Basededatos/Dbsishyt");
using var manifest = JsonDocument.Parse(File.ReadAllText(Path.Combine(root, "manifest.json")));
var paths = manifest.RootElement.GetProperty("scripts").EnumerateArray().Select(x => x.GetString()!).ToArray();
var parser = new TSql160Parser(true);
var failures = new List<string>();
bool ValidSignatures(string sql)
{
    var dates=Regex.Matches(sql,@"Create date:\s*(\d{2}/\d{2}/\d{4})").Select(match=>match.Groups[1].Value).ToArray();
    return dates.Length>0&&dates.Distinct().Count()==dates.Length&&dates.All(date=>DateOnly.TryParseExact(date,"dd/MM/yyyy",System.Globalization.CultureInfo.InvariantCulture,System.Globalization.DateTimeStyles.None,out _));
}

var tables = new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);
var candidateKeys = new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);
var foreignKeys = new List<(string Owner, string Target, string[] Local, string[] Remote)>();
var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
var allText = new Dictionary<string, string>();
foreach (var relative in paths)
{
    if (!seen.Add(relative)) failures.Add($"Script duplicado: {relative}");
    var path = Path.GetFullPath(Path.Combine(root, relative));
    if (!path.StartsWith(root + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase))
        throw new InvalidDataException("Ruta fuera del paquete.");
    var bytes = File.ReadAllBytes(path);
    var sql = new System.Text.UTF8Encoding(false, true).GetString(bytes);
    allText[relative] = sql;
    if (sql.Replace("\r\n", "").Contains('\n')) failures.Add($"Finales de linea no CRLF: {relative}");
    if (!ValidSignatures(sql)) failures.Add($"Firma incorrecta: {relative}");
    parser.Parse(new StringReader(sql), out var errors);
    failures.AddRange(errors.Select(e => $"{relative}:{e.Line}:{e.Column}: {e.Message}"));
    if (relative.StartsWith("StoreProcedure/"))
    {
        if (!sql.Contains("CREATE OR ALTER PROCEDURE") || !sql.Contains("SET NOCOUNT ON") || !sql.Contains("BEGIN TRY") || !sql.Contains("BEGIN CATCH"))
            failures.Add($"Contrato SP incompleto: {relative}");
        if (Regex.IsMatch(sql, @"SELECT\s+\*", RegexOptions.IgnoreCase)) failures.Add($"SELECT *: {relative}");
    }
    var match = Regex.Match(sql, @"CREATE TABLE dbo\.(\w+)\s*\((.*?)^    \);", RegexOptions.Singleline | RegexOptions.Multiline | RegexOptions.IgnoreCase);
    if (match.Success)
    {
        var name = match.Groups[1].Value;
        var definition = match.Groups[2].Value;
        var columns = Regex.Matches(definition, @"^        (\w+)\s+(?:N?VARCHAR|CHAR|INT|BIGINT|TINYINT|BIT|DATE|DATETIME2|DATETIMEOFFSET|DECIMAL|ROWVERSION|UNIQUEIDENTIFIER)\b", RegexOptions.Multiline | RegexOptions.IgnoreCase)
            .Select(m => m.Groups[1].Value).ToHashSet(StringComparer.OrdinalIgnoreCase);
        if (!tables.TryAdd(name, columns)) failures.Add($"Tabla duplicada: {name}");
        var keys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (Match key in Regex.Matches(definition, @"(?:PRIMARY KEY|UNIQUE)\s*\(([^)]+)\)", RegexOptions.IgnoreCase))
            keys.Add(string.Join(",", key.Groups[1].Value.Split(',').Select(x=>x.Trim())));
        foreach (Match key in Regex.Matches(definition, @"^        (\w+) [^\r\n]+?(?:PRIMARY KEY|UNIQUE)\s*[,]?\s*$", RegexOptions.Multiline | RegexOptions.IgnoreCase))
            keys.Add(key.Groups[1].Value);
        candidateKeys[name] = keys;

        foreach (Match fk in Regex.Matches(definition, @"FOREIGN KEY\s*\(([^)]+)\)\s*REFERENCES dbo\.(\w+)\s*\(([^)]+)\)", RegexOptions.IgnoreCase))
        {
            var target = fk.Groups[2].Value;
            if (!tables.ContainsKey(target)) failures.Add($"Dependencia fuera de orden: {name} -> {target}");
            foreignKeys.Add((name, target, fk.Groups[1].Value.Split(',').Select(x=>x.Trim()).ToArray(), fk.Groups[3].Value.Split(',').Select(x=>x.Trim()).ToArray()));
        }
    }
    foreach (Match addition in Regex.Matches(sql, @"ALTER TABLE dbo\.(\w+) ADD (\w+)\s", RegexOptions.IgnoreCase))
        if (tables.TryGetValue(addition.Groups[1].Value, out var existingColumns)) existingColumns.Add(addition.Groups[2].Value);
}
// Comprueba columnas de alias que se pueden resolver sin motor SQL. Los alias
// de OPENJSON, variables y temporales requieren comprobación manual de ejecución.
foreach (var entry in allText.Where(entry => entry.Key.StartsWith("StoreProcedure/")))
{
    var fragment = parser.Parse(new StringReader(entry.Value), out _);
    var references = new SqlReferences(); fragment.Accept(references);
    var aliases = new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);
    foreach (var reference in references.Tables)
    {
        if (!tables.TryGetValue(reference.SchemaObject.BaseIdentifier.Value, out var columns)) continue;
        var alias = reference.Alias?.Value ?? reference.SchemaObject.BaseIdentifier.Value;
        if (!aliases.TryGetValue(alias, out var possible)) aliases[alias] = possible = new(StringComparer.OrdinalIgnoreCase);
        possible.UnionWith(columns);
    }
    foreach (var column in references.Columns)
    {
        var identifiers = column.MultiPartIdentifier?.Identifiers;
        if (identifiers?.Count != 2) continue;
        if (aliases.TryGetValue(identifiers[0].Value, out var available) && !available.Contains(identifiers[1].Value))
            failures.Add($"Columna referida ausente: {entry.Key}: {identifiers[0].Value}.{identifiers[1].Value}");
    }
}
foreach (var fk in foreignKeys)
{
    if (!tables.TryGetValue(fk.Target, out var target)) { failures.Add($"Tabla ausente: {fk.Target}"); continue; }
    if (!candidateKeys[fk.Target].Contains(string.Join(",", fk.Remote))) failures.Add($"FK sin clave candidata: {fk.Owner} -> {fk.Target}");
    if (fk.Local.Length != fk.Remote.Length) failures.Add($"FK de dimensiones distintas: {fk.Owner}");
    foreach (var column in fk.Local) if (!tables[fk.Owner].Contains(column)) failures.Add($"Columna FK local ausente: {fk.Owner}.{column}");
    foreach (var column in fk.Remote) if (!target.Contains(column)) failures.Add($"Columna FK referida ausente: {fk.Target}.{column}");
}
foreach (var file in Directory.EnumerateFiles(root, "*.sql", SearchOption.AllDirectories))
{
    var rel = Path.GetRelativePath(root, file).Replace('\\', '/');
    if (!seen.Contains(rel)) failures.Add($"SQL sin registrar: {rel}");
}
void Require(string file, string fragment) { if (!allText[file].Contains(fragment)) failures.Add($"Invariante ausente: {file}: {fragment}"); }
Require("Tablas/Presupuesto.sql", "TiendaId BIGINT NOT NULL");
Require("Tablas/Presupuesto.sql", "FOREIGN KEY (ClienteId, TiendaId)");
Require("Tablas/GuiaConformidad.sql", "UQ_Guia_Presupuesto UNIQUE");
Require("Tablas/OrdenCompraCliente.sql", "UQ_OC_Presupuesto UNIQUE");
Require("Tablas/OrdenCompraCliente.sql", "NumeroOCCliente NVARCHAR");
Require("Tablas/ReferenciaComprobanteExterno.sql", "FOREIGN KEY (OrdenCompraId, ConfirmacionId)");
Require("Tablas/TarifarioPartida.sql", "PrecioUSD DECIMAL(19,6) NULL");
Require("Tablas/PresupuestoRevision.sql", "TasaIGVAplicada DECIMAL");
Require("StoreProcedure/spu_ListarClientes.sql", "FETCH NEXT 20 ROWS ONLY");
Require("StoreProcedure/spu_RegistrarTasaIGV.sql", "EXEC dbo.spu_RegistrarAuditoria");
Require("Script/008_PermisosAplicacion.sql", "DENY INSERT, UPDATE, DELETE ON SCHEMA::dbo");
Require("Script/007_ProteccionAuditoria.sql", "INSTEAD OF UPDATE, DELETE");
var packageDirectory = Path.GetFullPath(Path.Combine(root, "../../.NET2026/SistemaWebHYT/entregables"));
foreach (var package in new[] { (File: "ETAPA_1_HYT.sql", Hashes: "SHA256.json"), (File: "ETAPA_2_HYT.sql", Hashes: "SHA256_ETAPA_2.json"), (File: "ETAPA_3_HYT.sql", Hashes: "SHA256_ETAPA_3.json"), (File: "ETAPA_4_HYT.sql", Hashes: "SHA256_ETAPA_4.json") })
{
var bundlePath = Path.Combine(packageDirectory, package.File);
if (!File.Exists(bundlePath)) continue;
{
    var bundle = File.ReadAllText(bundlePath);
    parser.Parse(new StringReader(bundle), out var bundleErrors);
    failures.AddRange(bundleErrors.Select(e => $"Paquete:{e.Line}:{e.Column}: {e.Message}"));
    if (!ValidSignatures(bundle)) failures.Add("El paquete debe tener una sola firma.");
    using var hashes = JsonDocument.Parse(File.ReadAllText(Path.Combine(packageDirectory, package.Hashes)));
    foreach (var entry in hashes.RootElement.EnumerateArray())
    {
        var relative = entry.GetProperty("path").GetString()!;
        var file = relative == package.File ? bundlePath : Path.Combine(root, relative);
        var hash = Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(File.ReadAllBytes(file))).ToLowerInvariant();
        if (hash != entry.GetProperty("sha256").GetString()) failures.Add($"Huella no coincide: {relative}");
    }
}
}
if (failures.Count > 0) { foreach (var failure in failures) Console.Error.WriteLine(failure); return 1; }
Console.WriteLine($"VALIDADO OFFLINE: {paths.Length} scripts; {tables.Count} tablas; {foreignKeys.Count} relaciones. Sin ejecucion SQL.");
return 0;

sealed class SqlReferences : TSqlFragmentVisitor
{
    public List<NamedTableReference> Tables { get; } = [];
    public List<ColumnReferenceExpression> Columns { get; } = [];
    public override void ExplicitVisit(NamedTableReference node) { Tables.Add(node); base.ExplicitVisit(node); }
    public override void ExplicitVisit(ColumnReferenceExpression node) { Columns.Add(node); base.ExplicitVisit(node); }
}
