using GestionServicios.Application.Security;
using GestionServicios.Infrastructure.Security;
using System.Data;
using System.Text.Json;
using System.Text.RegularExpressions;
using Microsoft.Data.SqlClient;

// Firma: FRANCO LARA - 03/11/2026 | Pruebas SQL aisladas de presupuestos, impresiones, documentos, archivos, estados, regularización, historial y reversión por fallo de auditoría.

if (!args.Contains("--run-local-sql"))
{
    Console.WriteLine("Usa --run-local-sql para pruebas aisladas en localhost\\SQLEXPRESS.");
    return;
}
var directory = new DirectoryInfo(AppContext.BaseDirectory);
while (directory is not null && !File.Exists(Path.Combine(directory.FullName, "SistemaWebHYT.slnx"))) directory = directory.Parent;
var project = directory!.FullName;
var sources = Path.GetFullPath(Path.Combine(project, "../../Basededatos/Dbsishyt"));
var testDatabase = "SistemaWebHYT_Test_" + Guid.NewGuid().ToString("N");
const string localConnection = "Server=localhost\\SQLEXPRESS;Integrated Security=True;Encrypt=True;TrustServerCertificate=True;Pooling=False;Connect Timeout=10;";
const string actor = "sql-test-editor";
var checks = 0;
void Check(bool condition, string description)
{
    if (!condition) throw new InvalidOperationException(description);
    checks++; Console.WriteLine("OK SQL: " + description);
}
async Task Execute(SqlConnection connection, string text)
{
    await using var command = new SqlCommand(text, connection) { CommandTimeout = 60 };
    await command.ExecuteNonQueryAsync();
}
async Task<object?> Scalar(SqlConnection connection, string text)
{
    await using var command = new SqlCommand(text, connection);
    return await command.ExecuteScalarAsync();
}
async Task ExpectFailure(Func<Task> action, string message, string description)
{
    try { await action(); }
    catch (SqlException error) when (error.Message.Contains(message, StringComparison.Ordinal)) { Check(true, description); return; }
    throw new InvalidOperationException("No se rechazó: " + description);
}
async Task ExpectDomainFailure(Func<Task> action,string message,string description)
{
    try{await action();}
    catch(GestionServicios.Application.Maintenance.MaintenanceException error)when(error.Message.Contains(message,StringComparison.Ordinal)){Check(true,description);return;}
    throw new InvalidOperationException("No se rechazó: "+description);
}
async Task<BudgetVersion> Save(SqlConnection connection, object payload, BudgetVersion? prior = null, bool newRevision = false)
{
    await using var command = new SqlCommand("dbo.spu_Presupuesto_Guardar", connection) { CommandType = CommandType.StoredProcedure };
    command.Parameters.Add("@UsuarioId", SqlDbType.NVarChar, 128).Value = actor;
    command.Parameters.Add("@CorrelacionId", SqlDbType.UniqueIdentifier).Value = Guid.NewGuid();
    command.Parameters.Add("@Datos", SqlDbType.NVarChar, -1).Value = JsonSerializer.Serialize(payload);
    command.Parameters.Add("@PresupuestoId", SqlDbType.BigInt).Value = (object?)prior?.BudgetId ?? DBNull.Value;
    command.Parameters.Add("@RevisionId", SqlDbType.BigInt).Value = (object?)prior?.RevisionId ?? DBNull.Value;
    command.Parameters.Add("@VersionOriginal", SqlDbType.Binary, 8).Value = (object?)prior?.RevisionVersion ?? DBNull.Value;
    command.Parameters.Add("@VersionPresupuesto", SqlDbType.Binary, 8).Value = (object?)prior?.HeaderVersion ?? DBNull.Value;
    command.Parameters.Add("@NuevaRevision", SqlDbType.Bit).Value = newRevision;
    await using var reader = await command.ExecuteReaderAsync(); await reader.ReadAsync();
    return new(reader.GetInt64(0), reader.GetInt64(1), reader.GetString(2), reader.GetInt32(3), (byte[])reader.GetValue(4), (byte[])reader.GetValue(5));
}
async Task<BudgetVersion> State(SqlConnection connection, BudgetVersion prior, string state,string? reason="Prueba aislada")
{
    await using var command = new SqlCommand("dbo.spu_Presupuesto_CambiarEstado", connection) { CommandType = CommandType.StoredProcedure };
    command.Parameters.Add("@UsuarioId", SqlDbType.NVarChar, 128).Value = actor;
    command.Parameters.Add("@CorrelacionId", SqlDbType.UniqueIdentifier).Value = Guid.NewGuid();
    command.Parameters.Add("@PresupuestoId", SqlDbType.BigInt).Value = prior.BudgetId;
    command.Parameters.Add("@RevisionId", SqlDbType.BigInt).Value = prior.RevisionId;
    command.Parameters.Add("@VersionOriginal", SqlDbType.Binary, 8).Value = prior.RevisionVersion;
    command.Parameters.Add("@VersionPresupuesto", SqlDbType.Binary, 8).Value = prior.HeaderVersion;
    command.Parameters.Add("@NuevoEstado", SqlDbType.NVarChar, 20).Value = state;
    command.Parameters.Add("@FechaEvento", SqlDbType.Date).Value = new DateTime(2026,11,3);
    command.Parameters.Add("@Responsable", SqlDbType.NVarChar, 150).Value = "Representante de prueba";
    command.Parameters.Add("@Motivo", SqlDbType.NVarChar, 2000).Value = (object?)reason ?? DBNull.Value;
    await using var reader = await command.ExecuteReaderAsync(); await reader.ReadAsync();
    return prior with { HeaderVersion = (byte[])reader.GetValue(2), RevisionVersion = (byte[])reader.GetValue(3) };
}
async Task<long> SaveDocument(SqlConnection connection, string code, long parent, byte[] version, object payload, Guid operation)
{
    if (code is not ("GUIA" or "OC" or "CONFIRMACION" or "EJECUCION")) throw new ArgumentException("Documento no válido.");
    await using var command = new SqlCommand("dbo.spu_Seguimiento" + code + "_Guardar", connection) { CommandType = CommandType.StoredProcedure };
    command.Parameters.Add("@UsuarioId",SqlDbType.NVarChar,128).Value=actor;
    command.Parameters.Add("@CorrelacionId",SqlDbType.UniqueIdentifier).Value=operation;
    command.Parameters.Add("@PadreId",SqlDbType.BigInt).Value=parent;
    command.Parameters.Add("@VersionPadre",SqlDbType.Binary,8).Value=version;
    command.Parameters.Add("@Datos",SqlDbType.NVarChar,-1).Value=JsonSerializer.Serialize(payload);
    return Convert.ToInt64(await command.ExecuteScalarAsync());
}
object Payload(object[] lines, string currency = "PEN", string tax = "SIN_IGV", bool revalue = false, string description = "Trabajo de prueba", long store = 1, string attention = "EMERGENCIA")
    => new { ClienteId = 1, TiendaId = store, TipoAtencion = attention, Descripcion = description, RevalorizarMoneda = revalue,
        Revision = new { FechaEmision = "2026-11-03", CodigoMoneda = currency, CodigoTipoIGV = tax, NumeroCotizacionCliente = "COT-0001A", Consideraciones = "Prueba", Firmante = "Firmante de prueba" }, Partidas = lines };
object Manual(decimal? price = 100, string currency = "PEN", decimal quantity = 2)
    => new { OrigenPrecio = "MANUAL", Descripcion = "Servicio", Unidad = "UND", Cantidad = quantity, PrecioManual = price, CodigoMonedaOrigen = currency, CodigoTipoIGVOrigen = "SIN_IGV", MotivoPrecioManual = "Precio acordado" };
try
{
    await using (var setup = new SqlConnection(localConnection + "Initial Catalog=master;"))
    {
        await setup.OpenAsync();
        await Execute(setup, "SET QUOTED_IDENTIFIER ON; SET ANSI_NULLS ON;");
        using var manifest = JsonDocument.Parse(await File.ReadAllTextAsync(Path.Combine(sources,"manifest.json")));
        var installation=manifest.RootElement.GetProperty("scripts").EnumerateArray().Select(item=>Path.Combine(sources,item.GetString()!)).ToList();
        if(args.Contains("--use-bundles"))
        {
            var bundled=new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach(var stage in new[]{5,6,7,8})
            {
                using var stageManifest=JsonDocument.Parse(await File.ReadAllTextAsync(Path.Combine(sources,$"manifest-etapa{stage}.json")));
                foreach(var item in stageManifest.RootElement.GetProperty("scripts").EnumerateArray())bundled.Add(Path.Combine(sources,item.GetString()!));
            }
            installation.RemoveAll(source => bundled.Contains(source) ||
                Path.GetFileName(source) is "spu_Auditoria_Listar.sql" or "spu_Auditoria_Obtener.sql" or "032_PermisosAdministracion.sql");
            foreach(var stage in new[]{5,6,7,8})installation.Add(Path.Combine(project,$"entregables/ETAPA_{stage}_HYT.sql"));
            installation.Add(Path.Combine(project,"entregables/ADMINISTRACION_HYT.sql"));
            Console.WriteLine("Instalación aislada desde paquetes incrementales de etapas 5/6/7/8 y administración.");
        }
        foreach (var name in installation)
        {
            var text = (await File.ReadAllTextAsync(name)).Replace("SistemaWebHYT", testDatabase, StringComparison.Ordinal);
            foreach (var batch in Regex.Split(text, @"^[ \t]*GO[ \t]*(?:--[^\r\n]*)?\r?$", RegexOptions.Multiline | RegexOptions.IgnoreCase))
                if (!string.IsNullOrWhiteSpace(batch))
                    try { await Execute(setup, batch); } catch (Exception error) { throw new InvalidOperationException("Fuente: " + name, error); }
        }
    }
    await using var connection = new SqlConnection(localConnection + "Initial Catalog=" + testDatabase + ";");
    await connection.OpenAsync();
    await Execute(connection, """
        SET QUOTED_IDENTIFIER ON;
        INSERT INTO dbo.AspNetUsers (Id,Nombres,Activo) VALUES (N'sql-test-editor',N'Editor temporal',1);
        INSERT INTO dbo.AspNetUserRoles (UserId,RoleId) SELECT N'sql-test-editor',r.Id FROM dbo.AspNetRoles AS r WHERE r.NormalizedName=N'ADMINISTRADOR';
        INSERT INTO dbo.ConfiguracionHYT (Id,Ruc,RazonSocial,UsuarioRegistro) VALUES (1,N'00000000000',N'HYT PRUEBA AISLADA',N'sql-test-editor');
        INSERT INTO dbo.Cliente (Codigo,TipoEntidad,RazonSocial,UsuarioRegistro) VALUES (N'CLI1',N'JURIDICA',N'Cliente original',N'sql-test-editor'),(N'CLI2',N'JURIDICA',N'Otro cliente',N'sql-test-editor');
        INSERT INTO dbo.Tienda (ClienteId,Codigo,Nombre,CentroSAPFICO,UsuarioRegistro) VALUES (1,N'T1',N'Tienda original',N'F001',N'sql-test-editor'),(2,N'T2',N'Tienda ajena',N'F002',N'sql-test-editor');
        INSERT INTO dbo.UnidadMedida (Codigo,Descripcion,UsuarioRegistro) VALUES (N'UND',N'Unidad',N'sql-test-editor');
        INSERT INTO dbo.ServicioMaterial (CodigoInterno,Descripcion,UnidadMedidaId,UsuarioRegistro) VALUES (N'0001-A',N'Material original',1,N'sql-test-editor');
        INSERT INTO dbo.Tarifario (Codigo,Nombre,UsuarioRegistro) VALUES (N'TAR1',N'Tarifa de prueba',N'sql-test-editor');
        INSERT INTO dbo.TarifarioVersion (TarifarioId,NumeroVersion,CodigoMonedaPredeterminada,CodigoTipoIGVPredeterminado,Estado,UsuarioRegistro) VALUES (1,1,N'PEN',N'SIN_IGV',N'VIGENTE',N'sql-test-editor');
        INSERT INTO dbo.TarifarioPartida (VersionId,MaterialId,CodigoMoneda,CodigoTipoIGV,PrecioPEN,PrecioUSD,UsuarioRegistro) VALUES (1,1,N'PEN',N'SIN_IGV',100,NULL,N'sql-test-editor');
        """);
    var budget = await Save(connection, Payload([Manual()]));
    Check(Regex.IsMatch(budget.Number,@"^\d{12}$") && budget.RevisionNumber == 1 && Convert.ToString(await Scalar(connection,"SELECT NumeroInterno FROM dbo.Presupuesto WHERE PresupuestoId=1"))=="PPTO-1", "Numeración interna y primera revisión");
    Check(Convert.ToString(await Scalar(connection,$"SELECT TiendaCodigo FROM dbo.PresupuestoRevision WHERE RevisionId={budget.RevisionId}"))=="T1"&&Convert.ToString(await Scalar(connection,$"SELECT CentroSAPFICO FROM dbo.PresupuestoPartida WHERE RevisionId={budget.RevisionId}"))=="F001","Código de tienda y centro FICO se recuperan de datos reales");
    Check(Convert.ToDecimal(await Scalar(connection,$"SELECT Total FROM dbo.PresupuestoRevision WHERE RevisionId={budget.RevisionId}")) == 236m, "Totales reales calculados en SQL a dos decimales");
    await ExpectFailure(() => Save(connection,Payload([Manual(quantity:1.001m)])), "dos decimales", "Rechaza cantidades con tres decimales");
    await ExpectFailure(() => Save(connection,Payload([Manual()],store:2)), "cliente y tienda", "Rechaza tienda ajena al cliente");
    var previous = budget;
    budget = await Save(connection,Payload([Manual(150)],description:"Descripción actualizada"),budget);
    await ExpectFailure(() => Save(connection,Payload([Manual()]),previous), "cambió", "Rechaza guardado con versiones anteriores");
    Check(Convert.ToDecimal(await Scalar(connection,$"SELECT Total FROM dbo.PresupuestoRevision WHERE RevisionId={budget.RevisionId}")) == 354m, "Guardado concurrente fallido no modifica importes");
    budget = await State(connection,budget,"ENVIADO");
    await ExpectFailure(() => Save(connection,Payload([Manual()]),budget), "Edita solo", "Revisión enviada no permite editar contenido");
    var copy = await Save(connection,Payload([Manual(200)],description:"Nueva revisión"),budget,true);
    Check(copy.RevisionNumber == 2 && copy.RevisionId != budget.RevisionId, "Nueva revisión mantiene la anterior");
    Check(Convert.ToString(await Scalar(connection,$"SELECT DescripcionTrabajo FROM dbo.PresupuestoRevision WHERE RevisionId={budget.RevisionId}")) == "Descripción actualizada", "Descripción histórica conservada por revisión");
    Check(Convert.ToDecimal(await Scalar(connection,$"SELECT Total FROM dbo.PresupuestoRevision WHERE RevisionId={budget.RevisionId}")) == 354m, "Nueva revisión no recalcula la enviada");
    copy = await State(connection,copy,"ENVIADO");
    copy = await State(connection,copy,"APROBADO");
    Check(await Scalar(connection,$"SELECT FechaAprobacion FROM dbo.PresupuestoRevision WHERE RevisionId={copy.RevisionId}") is DateTime approval && approval.Date == new DateTime(2026,11,3), "Aprobación conserva fecha del evento");
    var source = new { OrigenPrecio = "TARIFARIO", PartidaTarifarioId = 1, Cantidad = 2, PrecioManual = 999, CodigoMonedaOrigen = "USD" };
    var tariffBudget = await Save(connection,Payload([source]));
    Check(Convert.ToDecimal(await Scalar(connection,$"SELECT Total FROM dbo.PresupuestoRevision WHERE RevisionId={tariffBudget.RevisionId}")) == 236m, "Tarifa usa precio de servidor e ignora precio y moneda manipulados");
    var dollar = await Save(connection,Payload([source],"USD"));
    Check(await Scalar(connection,$"SELECT Total FROM dbo.PresupuestoRevision WHERE RevisionId={dollar.RevisionId}") is DBNull, "USD sin precio permanece NULL");
    await ExpectFailure(() => State(connection,dollar,"ENVIADO"), "Completa las partidas", "No envía presupuesto con precios pendientes");
    await ExpectFailure(() => Save(connection,Payload([Manual(currency:"USD")],"USD"),tariffBudget), "revalorización", "Cambio de moneda exige confirmación explícita");
    var dollarsRevalued = await Save(connection,Payload([Manual(50,"USD")],"USD",revalue:true),tariffBudget);
    Check(Convert.ToDecimal(await Scalar(connection,$"SELECT Total FROM dbo.PresupuestoRevision WHERE RevisionId={dollarsRevalued.RevisionId}")) == 118m, "Revalorización manual USD recalcula sin conversión cambiaria");
    var gross = await Save(connection,Payload([Manual(118)],tax:"CON_IGV"));
    Check(Convert.ToDecimal(await Scalar(connection,$"SELECT Total FROM dbo.PresupuestoRevision WHERE RevisionId={gross.RevisionId}")) == 278.48m, "Conversión SIN IGV a CON IGV usa unitario de dos decimales");
    Check(Convert.ToInt64(await Scalar(connection,"SELECT COUNT_BIG(1) FROM dbo.EventoAuditoria WHERE Entidad IN (N'Presupuesto',N'PresupuestoRevision')")) >= 15, "Guardados y transiciones tienen auditoría persistida");
    var repository = new GestionServicios.Infrastructure.Budgets.SqlBudgetRepository(new GestionServicios.Infrastructure.SqlConnectionFactory(localConnection+"Initial Catalog="+testDatabase+";"));
    var listed = await repository.ListAsync(actor,DateTimeOffset.UtcNow.AddDays(-1),DateTimeOffset.UtcNow.AddDays(1),null,1,null,CancellationToken.None);
    Check(listed.Total == 4 && listed.Rows.Count == 4 && listed.Rows.Any(row=>row.Total is null), "Repositorio ADO.NET mapea listado y totales incompletos reales");
    var recovered = await repository.GetAsync(actor,budget.BudgetId,budget.RevisionId,1,CancellationToken.None);
    Check(recovered is not null && recovered.Revision.GetProperty("Total").GetDecimal()==354m && recovered.Audit.Count>0, "Repositorio recupera snapshot histórico y eventos reales");
    var choices = await repository.TariffsAsync(actor,1,new DateOnly(2026,11,3),"USD",null,1,CancellationToken.None);
    Check(choices.Total==1 && choices.Rows[0].Price is null && choices.Rows[0].InternalCode=="0001-A", "Ayuda ADO.NET conserva ceros y precio USD ausente");
    var defaults=await repository.DefaultsAsync(actor,new DateOnly(2026,11,3),1,1,CancellationToken.None);
    Check(defaults.GetProperty("Tasa").GetProperty("Valor").GetDecimal()==0.18m&&defaults.GetProperty("Tienda").GetProperty("Nombre").GetString()=="Tienda original", "Datos iniciales reales devuelven tasa y tienda del cliente");
    var companyRepository=new GestionServicios.Infrastructure.Maintenance.SqlCompanySettingsRepository(new GestionServicios.Infrastructure.SqlConnectionFactory(localConnection+"Initial Catalog="+testDatabase+";"));
    var company=await companyRepository.GetAsync(actor,1,CancellationToken.None);
    var companyPayload=JsonSerializer.Serialize(new {Ruc="00000000000",RazonSocial="HYT DATOS NUEVOS",Direccion="Dirección nueva",NombreFirmante="Firmante",CargoFirmante="Comercial"});
    await companyRepository.SaveAsync(actor,companyPayload,Convert.FromBase64String(company!.Text("VersionFila")),CancellationToken.None);
    var companyUpdated=await companyRepository.GetAsync(actor,1,CancellationToken.None);
    Check(companyUpdated?.Text("RazonSocial")=="HYT DATOS NUEVOS"&&companyUpdated.AuditCount==1,"Datos HYT se actualizan con auditoría real");
    await ExpectDomainFailure(()=>companyRepository.SaveAsync(actor,companyPayload,Convert.FromBase64String(company.Text("VersionFila")),CancellationToken.None),"cambió","Configuración rechaza versión anterior");
    Check(Convert.ToString(await Scalar(connection,$"SELECT JSON_VALUE(EmpresaSnapshotJson,'$.RazonSocial') FROM dbo.PresupuestoRevision WHERE RevisionId={budget.RevisionId}"))=="HYT PRUEBA AISLADA","Cambiar datos HYT no reescribe la empresa de revisiones históricas");
    var fromRepository = await repository.SaveAsync(actor,JsonSerializer.Serialize(Payload([Manual(10)])),null,false,CancellationToken.None);
    await repository.ChangeStateAsync(actor,fromRepository,"ENVIADO",new DateOnly(2026,11,3),"Responsable de prueba",null,CancellationToken.None);
    var sent = await repository.GetAsync(actor,fromRepository.BudgetId,null,1,CancellationToken.None);
    Check(sent?.Revision.GetProperty("Estado").GetString()=="ENVIADO" && sent.Audit.Any(item=>item.EventDate==new DateOnly(2026,11,3)), "Guardado y transición ADO.NET con fecha de evento sin desfase horario");
    var printRepository=new GestionServicios.Infrastructure.Budgets.SqlBudgetPrintRepository(new GestionServicios.Infrastructure.SqlConnectionFactory(localConnection+"Initial Catalog="+testDatabase+";"));
    var printHtml="<!DOCTYPE html><html lang=\"es\"><head><meta charset=\"utf-8\"></head><body>Presupuesto histórico de prueba. Importe 23.60. Moneda PEN. Texto con ñ y tildes.</body></html>";
    var printVersion=Convert.FromBase64String(GestionServicios.Application.Budgets.BudgetRecord.Text(sent!.Revision,"VersionFila"));
    var newPrint=await printRepository.SaveAsync(actor,fromRepository.RevisionId,printVersion,"NUEVO","HYT-2026.1",printHtml,CancellationToken.None);
    await printRepository.SaveAsync(actor,fromRepository.RevisionId,printVersion,"ANTERIOR","HYT-2026.1",printHtml,CancellationToken.None);
    var savedPrint=await printRepository.GetAsync(actor,newPrint,CancellationToken.None);
    Check(savedPrint?.Html==printHtml&&savedPrint.Hash.Length==64,"HTML Unicode y huella se conservan y verifican contra SQL");
    var printHistory=await printRepository.ListAsync(actor,fromRepository.RevisionId,1,CancellationToken.None);
    Check(printHistory.Total==2&&printHistory.Rows.Select(row=>row.Format).Distinct().Count()==2,"Historial SQL conserva ambos formatos con versión y usuario");
    await ExpectDomainFailure(()=>printRepository.SaveAsync(actor,fromRepository.RevisionId,fromRepository.RevisionVersion,"NUEVO","HYT-2026.1",printHtml,CancellationToken.None),"cambió","Generación rechaza revisión cambiada durante el renderizado");
    Check(Convert.ToInt64(await Scalar(connection,$"SELECT COUNT_BIG(1) FROM dbo.EventoAuditoria WHERE Entidad=N'PresupuestoRevision' AND RegistroId=N'{fromRepository.RevisionId}' AND Accion=N'GENERAR_IMPRESION'"))==2,"Cada generación queda auditada en la revisión");
    await Execute(connection,$"UPDATE dbo.ImpresionPresupuesto SET ContenidoHTML=ContenidoHTML+N'alterado' WHERE ImpresionId={newPrint};");
    await ExpectDomainFailure(()=>printRepository.GetAsync(actor,newPrint,CancellationToken.None),"huella","Contenido alterado no se entrega como impresión íntegra");
    await Execute(connection,"""
        INSERT INTO dbo.AspNetUsers (Id,Nombres,Activo) VALUES (N'sql-test-budget-only',N'Solo presupuesto',1);
        INSERT INTO dbo.AspNetRoles (Id,Name,NormalizedName) VALUES (N'sql-test-budget-role',N'Solo consulta presupuesto',N'SOLO_CONSULTA_PRESUPUESTO');
        INSERT INTO dbo.AspNetUserRoles (UserId,RoleId) VALUES (N'sql-test-budget-only',N'sql-test-budget-role');
        INSERT INTO dbo.RolPermiso (RoleId,CodigoPermiso) VALUES (N'sql-test-budget-role',N'PRESUPUESTOS.CONSULTAR');
        """);
    await ExpectDomainFailure(()=>printRepository.GetAsync("sql-test-budget-only",newPrint,CancellationToken.None),"permiso","SQL impide consultar impresiones sin permiso de reportes");
    var historicalSource = await Save(connection,Payload([source]));
    var originalLineId = Convert.ToInt64(await Scalar(connection,$"SELECT PartidaId FROM dbo.PresupuestoPartida WHERE RevisionId={historicalSource.RevisionId}"));
    await Execute(connection,"UPDATE dbo.ServicioMaterial SET Activo=0,Descripcion=N'Descripción posterior' WHERE MaterialId=1;");
    historicalSource = await Save(connection,Payload([new { PartidaId=originalLineId,OrigenPrecio="TARIFARIO",PartidaTarifarioId=1,Cantidad=3 }]),historicalSource);
    Check(Convert.ToString(await Scalar(connection,$"SELECT Descripcion FROM dbo.PresupuestoPartida WHERE RevisionId={historicalSource.RevisionId}"))=="Material original" &&
        Convert.ToDecimal(await Scalar(connection,$"SELECT Total FROM dbo.PresupuestoRevision WHERE RevisionId={historicalSource.RevisionId}"))==354m, "Edición conserva referencia y descripción histórica de material desactivado");
    await Execute(connection,"DELETE FROM dbo.ConfiguracionHYT WHERE Id=1;");
    var withoutCompany = await Save(connection,Payload([Manual(10)]));
    await ExpectFailure(()=>State(connection,withoutCompany,"ENVIADO"),"Configura los datos de HYT","Permite preparar borrador sin empresa configurada y bloquea su envío");
    await companyRepository.SaveAsync(actor,companyPayload,null,CancellationToken.None);
    Check((await companyRepository.GetAsync(actor,1,CancellationToken.None))?.Text("RazonSocial")=="HYT DATOS NUEVOS","Configuración corporativa permite el registro inicial único");
    await Execute(connection,"""
        INSERT INTO dbo.Tarifario (Codigo,Nombre,ClienteId,UsuarioRegistro) VALUES (N'TAR2',N'Tarifa cliente ajeno',2,N'sql-test-editor');
        INSERT INTO dbo.TarifarioVersion (TarifarioId,NumeroVersion,CodigoMonedaPredeterminada,CodigoTipoIGVPredeterminado,Estado,UsuarioRegistro) VALUES (2,1,N'PEN',N'SIN_IGV',N'VIGENTE',N'sql-test-editor');
        INSERT INTO dbo.TarifarioPartida (VersionId,MaterialId,CodigoMoneda,CodigoTipoIGV,PrecioPEN,UsuarioRegistro) VALUES (2,1,N'PEN',N'SIN_IGV',10,N'sql-test-editor');
        """);
    await ExpectFailure(()=>Save(connection,Payload([new {OrigenPrecio="MANUAL",PartidaTarifarioId=2,Descripcion="Servicio",Unidad="UND",Cantidad=1,PrecioManual=10,CodigoMonedaOrigen="PEN",CodigoTipoIGVOrigen="SIN_IGV",MotivoPrecioManual="Acordado"}])),"no corresponde al material o cliente","Precio manual no puede referenciar tarifario privado de otro cliente");
    var emergency = await Save(connection,Payload([]));
    var guideOperation=Guid.NewGuid();
    var guidePayload=new {FechaActa="2026-11-03",NumeroActa="GUIA-0007B",Observacion="Emergencia sin OC"};
    var guideId=await SaveDocument(connection,"GUIA",emergency.BudgetId,emergency.HeaderVersion,guidePayload,guideOperation);
    Check(guideId>0,"Emergencia en borrador permite guía sin OC ni partidas");
    Check(await SaveDocument(connection,"GUIA",emergency.BudgetId,emergency.HeaderVersion,guidePayload,guideOperation)==guideId,"Reintento de guía es idempotente con versión original");
    Check(Convert.ToString(await Scalar(connection,$"SELECT NumeroActa FROM dbo.GuiaConformidad WHERE GuiaId={guideId}"))=="GUIA-0007B","Guía conserva número alfanumérico y ceros");
    Check(Convert.ToInt64(await Scalar(connection,$"SELECT COUNT_BIG(1) FROM dbo.EventoAuditoria WHERE Entidad=N'GuiaConformidad' AND RegistroId=N'{guideId}'"))==1,"Reintento no duplica auditoría de creación");
    var emergencyVersion=(byte[])(await Scalar(connection,$"SELECT VersionFila FROM dbo.Presupuesto WHERE PresupuestoId={emergency.BudgetId}"))!;
    try { await SaveDocument(connection,"GUIA",emergency.BudgetId,emergencyVersion,guidePayload,Guid.NewGuid()); throw new InvalidOperationException("Guía duplicada aceptada."); }
    catch(SqlException) { Check(true,"Segunda guía del mismo presupuesto es rechazada"); }
    var invitation=await Save(connection,Payload([Manual()],attention:"INVITACION"));
    var ocPayload=new {RevisionId=invitation.RevisionId,NumeroOCCliente="OC-000A/01",FechaEmisionOC="2026-11-03",FechaRecepcionOC="2026-11-03",Importe=236m};
    invitation=await State(connection,invitation,"ENVIADO");
    var ocId=await SaveDocument(connection,"OC",invitation.BudgetId,invitation.HeaderVersion,ocPayload,Guid.NewGuid());
    Check(ocId>0&&Convert.ToString(await Scalar(connection,$"SELECT NumeroOCCliente FROM dbo.OrdenCompraCliente WHERE OrdenCompraId={ocId}"))=="OC-000A/01","OC conserva número externo alfanumérico después de aprobar invitación");
    Check(Convert.ToString(await Scalar(connection,$"SELECT Estado FROM dbo.PresupuestoRevision WHERE RevisionId={invitation.RevisionId}"))=="APROBADO","Registrar OC en invitación enviada acepta automáticamente el presupuesto");
    Check(Convert.ToInt64(await Scalar(connection,$"SELECT COUNT_BIG(1) FROM dbo.EventoAuditoria WHERE Entidad=N'PresupuestoRevision' AND RegistroId=N'{invitation.RevisionId}' AND EstadoNuevo=N'APROBADO' AND Motivo LIKE N'%automática%OC%'"))==1,"OC registra una aceptación auditada con origen documental");
    var parentVersion=(byte[])(await Scalar(connection,$"SELECT VersionFila FROM dbo.Presupuesto WHERE PresupuestoId={invitation.BudgetId}"))!;
    await ExpectFailure(()=>SaveDocument(connection,"OC",invitation.BudgetId,parentVersion,ocPayload,Guid.NewGuid()),"Ya existe","Un presupuesto no admite segunda OC");
    var ocVersion=(byte[])(await Scalar(connection,$"SELECT VersionFila FROM dbo.OrdenCompraCliente WHERE OrdenCompraId={ocId}"))!;
    var confirmationPayload=new {FechaConfirmacion="2026-11-03",NumeroConfirmacion="ACEP-0002B",Observacion="Aceptación del cliente"};
    var confirmationOperation=Guid.NewGuid();
    var confirmationId=await SaveDocument(connection,"CONFIRMACION",ocId,ocVersion,confirmationPayload,confirmationOperation);
    Check(confirmationId>0,"Confirmación del cliente se registra en su OC");
    Check(Convert.ToString(await Scalar(connection,$"SELECT Estado FROM dbo.OrdenCompraCliente WHERE OrdenCompraId={ocId}"))=="TRABAJO_COMPLETADO","Confirmación completa OC automáticamente sin ejecución registrada");
    Check(await SaveDocument(connection,"CONFIRMACION",ocId,ocVersion,confirmationPayload,confirmationOperation)==confirmationId&&Convert.ToInt64(await Scalar(connection,$"SELECT COUNT_BIG(1) FROM dbo.EventoAuditoria WHERE Entidad=N'OrdenCompraCliente' AND RegistroId=N'{ocId}' AND Accion=N'CAMBIAR_ESTADO_AUTOMATICO' AND EstadoAnterior=N'REGISTRADA' AND EstadoNuevo=N'TRABAJO_COMPLETADO' AND FechaEvento='20261103'"))==1,"Reintento de confirmación no duplica cambio automático auditado con fecha del cliente");
    ocVersion=(byte[])(await Scalar(connection,$"SELECT VersionFila FROM dbo.OrdenCompraCliente WHERE OrdenCompraId={ocId}"))!;
    await ExpectFailure(()=>SaveDocument(connection,"CONFIRMACION",ocId,ocVersion,confirmationPayload,Guid.NewGuid()),"Ya existe","OC impide dos confirmaciones vigentes");
    foreach(var entry in new[]{(Code:"GUIA",Id:guideId),(Code:"OC",Id:ocId),(Code:"CONFIRMACION",Id:confirmationId)})
    {
        await using var getCommand=new SqlCommand("dbo.spu_Seguimiento"+entry.Code+"_Obtener",connection){CommandType=CommandType.StoredProcedure};
        getCommand.Parameters.AddWithValue("@UsuarioId",actor);getCommand.Parameters.AddWithValue("@Id",entry.Id);
        var json=Convert.ToString(await getCommand.ExecuteScalarAsync());
        using var document=JsonDocument.Parse(json!);
        Check(document.RootElement.GetProperty("Documento").ValueKind==JsonValueKind.Object,"Consulta SQL recupera contexto y documento "+entry.Code);
        await using var listCommand=new SqlCommand("dbo.spu_Seguimiento"+entry.Code+"_Listar",connection){CommandType=CommandType.StoredProcedure};
        listCommand.Parameters.AddWithValue("@UsuarioId",actor);listCommand.Parameters.AddWithValue("@Desde",DateTime.UtcNow.AddDays(-1));listCommand.Parameters.AddWithValue("@HastaExclusiva",DateTime.UtcNow.AddDays(1));
        using var rows=JsonDocument.Parse(Convert.ToString(await listCommand.ExecuteScalarAsync())!);
        Check(rows.RootElement.GetArrayLength()>0,"Listado SQL devuelve documento "+entry.Code);
    }
    var documents=new GestionServicios.Infrastructure.Followup.SqlDocumentRepository(new GestionServicios.Infrastructure.SqlConnectionFactory(localConnection+"Initial Catalog="+testDatabase+";"));
    var ocKind=GestionServicios.Application.Followup.DocumentKind.Resolve("oc-cliente");
    var actualOc=await documents.GetAsync(ocKind,actor,ocId,null,1,CancellationToken.None);
    Check(actualOc is not null&&actualOc.Audit.Count==2&&actualOc.Data.GetProperty("Documento").GetProperty("NumeroOCCliente").GetString()=="OC-000A/01","Repositorio documental mapea ficha, contexto y auditoría reales");
    var actualPage=await documents.ListAsync(ocKind,actor,DateTimeOffset.UtcNow.AddDays(-1),DateTimeOffset.UtcNow.AddDays(1),"OC-000A/01",1,null,true,CancellationToken.None);
    Check(actualPage.Total==1&&actualPage.Rows.Count==1,"Repositorio documental mapea listado y filtro de confirmación");
    var noConfirmation=await documents.ListAsync(ocKind,actor,DateTimeOffset.UtcNow.AddDays(-1),DateTimeOffset.UtcNow.AddDays(1),null,1,null,false,CancellationToken.None);
    Check(noConfirmation.Total==0,"Filtro sin confirmar excluye OC con confirmación vigente");
    var executionKind=GestionServicios.Application.Followup.DocumentKind.Resolve("ejecucion");
    var execution=await documents.SaveAsync(executionKind,actor,emergency.BudgetId,emergencyVersion,JsonSerializer.Serialize(new{Responsable="Operaciones",FechaProgramada="2026-11-04"}),Guid.NewGuid(),null,null,CancellationToken.None);
    var executionRecord=await documents.GetAsync(executionKind,actor,execution.Id,null,1,CancellationToken.None);
    Check(executionRecord?.Data.GetProperty("Documento").GetProperty("EstadoDocumento").GetString()=="PENDIENTE","Repositorio guarda ejecución sin cambiar su estado automáticamente");
    await ExpectDomainFailure(()=>documents.GetAsync(ocKind,"sql-test-budget-only",ocId,null,1,CancellationToken.None),"permiso","Repositorio protege consulta documental con permisos SQL");
    var guideKind=GestionServicios.Application.Followup.DocumentKind.Resolve("guias");
    var attachmentOperation=Guid.NewGuid();
    var guideParent=(byte[])(await Scalar(connection,$"SELECT VersionFila FROM dbo.Presupuesto WHERE PresupuestoId={emergency.BudgetId}"))!;
    var guideVersion=(byte[])(await Scalar(connection,$"SELECT VersionFila FROM dbo.GuiaConformidad WHERE GuiaId={guideId}"))!;
    var attachmentPayload=JsonSerializer.Serialize(new{FechaActa="2026-11-03",NumeroActa="GUIA-0007B",MotivoCambio="Adjuntar escaneado",Adjunto=new{NombreOriginal="Guía escaneada.pdf",ClaveAlmacenamiento="Guia/GUIA-0007B/"+attachmentOperation.ToString("N")+".pdf",TipoContenido="application/pdf",TamanoBytes=15,SHA256=new string('A',64)}});
    var attached=await documents.SaveAsync(guideKind,actor,emergency.BudgetId,guideParent,attachmentPayload,attachmentOperation,guideId,guideVersion,CancellationToken.None);
    var filesRepository=new GestionServicios.Infrastructure.Followup.SqlDocumentFileRepository(new GestionServicios.Infrastructure.SqlConnectionFactory(localConnection+"Initial Catalog="+testDatabase+";"));
    var file=await filesRepository.GetAsync(guideKind,actor,attached.FileId!.Value,CancellationToken.None);
    Check(file?.Metadata.NombreOriginal=="Guía escaneada.pdf","Adjunto SQL conserva nombre Unicode y metadatos vinculados");
    Check(await filesRepository.GetAsync(ocKind,actor,attached.FileId.Value,CancellationToken.None) is null,"Archivo de guía no se expone bajo ruta OC");
    await ExpectDomainFailure(()=>filesRepository.GetAsync(guideKind,"sql-test-budget-only",attached.FileId.Value,CancellationToken.None),"permiso","Adjunto requiere permiso específico del documento");
    var replayAttachment=await documents.SaveAsync(guideKind,actor,emergency.BudgetId,guideParent,attachmentPayload,attachmentOperation,guideId,guideVersion,CancellationToken.None);
    Check(replayAttachment.FileId==attached.FileId&&Convert.ToInt64(await Scalar(connection,$"SELECT COUNT_BIG(1) FROM dbo.EventoAuditoria WHERE Entidad=N'ArchivoAdjunto' AND RegistroId=N'{attached.FileId}' AND Accion=N'CREAR'"))==1,"Reintento con archivo no duplica metadatos ni auditoría");
    guideParent=(byte[])(await Scalar(connection,$"SELECT VersionFila FROM dbo.Presupuesto WHERE PresupuestoId={emergency.BudgetId}"))!;
    guideVersion=(byte[])(await Scalar(connection,$"SELECT VersionFila FROM dbo.GuiaConformidad WHERE GuiaId={guideId}"))!;
    await documents.SaveAsync(guideKind,actor,emergency.BudgetId,guideParent,JsonSerializer.Serialize(new{FechaActa="2026-11-03",NumeroActa="GUIA-0007B",MotivoCambio="Escaneado reemplazado",RetirarAdjuntos=new[]{attached.FileId.Value}}),Guid.NewGuid(),guideId,guideVersion,CancellationToken.None);
    Check(await filesRepository.GetAsync(guideKind,actor,attached.FileId.Value,CancellationToken.None) is null,"Archivo retirado deja de estar disponible para visualización");
    Check(Convert.ToInt64(await Scalar(connection,$"SELECT COUNT_BIG(1) FROM dbo.EventoAuditoria WHERE Entidad=N'ArchivoAdjunto' AND RegistroId=N'{attached.FileId}' AND Accion=N'RETIRAR'"))==1,"Retiro de archivo conserva auditoría y motivo");
    async Task DocumentState(string key,long id,long parent,string next,string date="2026-11-03",Guid? operation=null,GestionServicios.Application.Followup.DocumentRecord? original=null)
    {
        var kind=GestionServicios.Application.Followup.DocumentKind.Resolve(key);
        original??=await documents.GetAsync(kind,actor,id,parent,1,CancellationToken.None);
        await using var command=new SqlCommand("dbo.spu_Seguimiento"+kind.Code+"_CambiarEstado",connection){CommandType=CommandType.StoredProcedure};
        command.Parameters.AddWithValue("@UsuarioId",actor);command.Parameters.AddWithValue("@CorrelacionId",operation??Guid.NewGuid());
        command.Parameters.AddWithValue("@Id",id);command.Parameters.AddWithValue("@PadreId",parent);
        command.Parameters.Add("@VersionPadre",SqlDbType.Binary,8).Value=Convert.FromBase64String(original!.Data.GetProperty("VersionPadre").GetString()!);
        command.Parameters.Add("@VersionOriginal",SqlDbType.Binary,8).Value=Convert.FromBase64String(original.Data.GetProperty("Documento").GetProperty("VersionFila").GetString()!);
        command.Parameters.AddWithValue("@NuevoEstado",next);command.Parameters.Add("@FechaEvento",SqlDbType.Date).Value=DateOnly.Parse(date).ToDateTime(TimeOnly.MinValue);
        command.Parameters.AddWithValue("@Responsable","Representante de prueba");command.Parameters.AddWithValue("@Motivo","Cambio documentado en prueba aislada");
        await command.ExecuteNonQueryAsync();
    }
    await DocumentState("ejecucion",execution.Id,emergency.BudgetId,"PROGRAMADA");
    await DocumentState("ejecucion",execution.Id,emergency.BudgetId,"EN_EJECUCION");
    await ExpectFailure(()=>DocumentState("ejecucion",execution.Id,emergency.BudgetId,"TERMINADA","2026-11-02"),"fecha de término","Trabajo no puede terminar antes de su inicio real");
    await DocumentState("ejecucion",execution.Id,emergency.BudgetId,"TERMINADA","2026-11-04");
    Check(Convert.ToString(await Scalar(connection,$"SELECT Estado FROM dbo.Ejecucion WHERE EjecucionId={execution.Id}"))=="TERMINADA","Emergencia termina ejecución sin OC");
    await DocumentState("oc-cliente",ocId,invitation.BudgetId,"REGISTRADA");
    await DocumentState("oc-cliente",ocId,invitation.BudgetId,"TRABAJO_COMPLETADO");
    Check(Convert.ToString(await Scalar(connection,$"SELECT Estado FROM dbo.OrdenCompraCliente WHERE OrdenCompraId={ocId}"))=="TRABAJO_COMPLETADO","Cambio manual admite confirmación vigente sin exigir ejecución terminada");
    await ExpectFailure(()=>DocumentState("oc-cliente",ocId,invitation.BudgetId,"CERRADA_OPERATIVAMENTE"),"ejecución terminada","Cierre operativo conserva requisito de ejecución terminada");
    var invitationParent=(byte[])(await Scalar(connection,$"SELECT VersionFila FROM dbo.Presupuesto WHERE PresupuestoId={invitation.BudgetId}"))!;
    var inviteExecution=await documents.SaveAsync(executionKind,actor,invitation.BudgetId,invitationParent,JsonSerializer.Serialize(new{Responsable="Técnico",FechaProgramada="2026-11-03"}),Guid.NewGuid(),null,null,CancellationToken.None);
    await ExpectFailure(()=>DocumentState("ejecucion",inviteExecution.Id,invitation.BudgetId,"TERMINADA"),"Transición","Ejecución pendiente no salta directamente a terminada");
    await DocumentState("ejecucion",inviteExecution.Id,invitation.BudgetId,"PROGRAMADA");
    await DocumentState("ejecucion",inviteExecution.Id,invitation.BudgetId,"EN_EJECUCION");
    await DocumentState("ejecucion",inviteExecution.Id,invitation.BudgetId,"TERMINADA","2026-11-04");
    await ExpectFailure(()=>DocumentState("oc-cliente",ocId,invitation.BudgetId,"CERRADA_OPERATIVAMENTE"),"guía registrada","Cierre OC requiere guía registrada");
    invitationParent=(byte[])(await Scalar(connection,$"SELECT VersionFila FROM dbo.Presupuesto WHERE PresupuestoId={invitation.BudgetId}"))!;
    var inviteGuide=await documents.SaveAsync(guideKind,actor,invitation.BudgetId,invitationParent,JsonSerializer.Serialize(new{FechaActa="2026-11-04",NumeroActa="INV-0008C"}),Guid.NewGuid(),null,null,CancellationToken.None);
    await DocumentState("oc-cliente",ocId,invitation.BudgetId,"CERRADA_OPERATIVAMENTE");
    Check(Convert.ToString(await Scalar(connection,$"SELECT Estado FROM dbo.OrdenCompraCliente WHERE OrdenCompraId={ocId}"))=="CERRADA_OPERATIVAMENTE","OC cierra después de ejecución terminada y guía registrada");
    var closedConfirmation=(await documents.GetAsync(GestionServicios.Application.Followup.DocumentKind.Resolve("confirmaciones"),actor,confirmationId,null,1,CancellationToken.None))!;
    await documents.SaveAsync(GestionServicios.Application.Followup.DocumentKind.Resolve("confirmaciones"),actor,ocId,Convert.FromBase64String(closedConfirmation.Data.GetProperty("VersionPadre").GetString()!),JsonSerializer.Serialize(new{FechaConfirmacion="2026-11-03",NumeroConfirmacion="ACEP-0002B",MotivoCambio="Corregir observación sin reabrir OC"}),Guid.NewGuid(),confirmationId,Convert.FromBase64String(closedConfirmation.Data.GetProperty("Documento").GetProperty("VersionFila").GetString()!),CancellationToken.None);
    Check(Convert.ToString(await Scalar(connection,$"SELECT Estado FROM dbo.OrdenCompraCliente WHERE OrdenCompraId={ocId}"))=="CERRADA_OPERATIVAMENTE","Editar confirmación no degrada una OC cerrada operativamente");
    await DocumentState("oc-cliente",ocId,invitation.BudgetId,"TRABAJO_COMPLETADO");
    await ExpectFailure(()=>DocumentState("ejecucion",inviteExecution.Id,invitation.BudgetId,"EN_EJECUCION"),"Reabre la OC","Trabajo no se reabre mientras OC figura trabajo completado");
    await DocumentState("oc-cliente",ocId,invitation.BudgetId,"REGISTRADA");
    await DocumentState("ejecucion",inviteExecution.Id,invitation.BudgetId,"EN_EJECUCION");
    Check(await Scalar(connection,$"SELECT FechaTerminoReal FROM dbo.Ejecucion WHERE EjecucionId={inviteExecution.Id}") is DBNull,"Reapertura de trabajo limpia término actual y conserva historial");
    await DocumentState("confirmaciones",confirmationId,ocId,"RETIRADA");
    ocVersion=(byte[])(await Scalar(connection,$"SELECT VersionFila FROM dbo.OrdenCompraCliente WHERE OrdenCompraId={ocId}"))!;
    var replacementConfirmation=await documents.SaveAsync(GestionServicios.Application.Followup.DocumentKind.Resolve("confirmaciones"),actor,ocId,ocVersion,JsonSerializer.Serialize(new{FechaConfirmacion="2026-11-04",NumeroConfirmacion="ACEP-REEMPLAZO"}),Guid.NewGuid(),null,null,CancellationToken.None);
    Check(replacementConfirmation.Id!=confirmationId&&Convert.ToInt64(await Scalar(connection,$"SELECT COUNT_BIG(1) FROM dbo.ConfirmacionFacturacion WHERE OrdenCompraId={ocId}"))==2&&Convert.ToInt64(await Scalar(connection,$"SELECT COUNT_BIG(1) FROM dbo.ConfirmacionFacturacion WHERE OrdenCompraId={ocId} AND Vigente=1"))==1,"Retiro de confirmación conserva historia y permite solo una nueva vigente");
    var executionForRepository=await documents.GetAsync(executionKind,actor,inviteExecution.Id,null,1,CancellationToken.None);
    await documents.ChangeStateAsync(executionKind,actor,inviteExecution.Id,invitation.BudgetId,Convert.FromBase64String(executionForRepository!.Data.GetProperty("VersionPadre").GetString()!),Convert.FromBase64String(executionForRepository.Data.GetProperty("Documento").GetProperty("VersionFila").GetString()!),Guid.NewGuid(),"TERMINADA",new DateOnly(2026,11,5),"Técnico responsable","Término real",CancellationToken.None);
    Check(await Scalar(connection,$"SELECT FechaTerminoReal FROM dbo.Ejecucion WHERE EjecucionId={inviteExecution.Id}") is DateTime repositoryDate&&repositoryDate.Date==new DateTime(2026,11,5),"Repositorio de estados conserva fecha real sin desfase horario");
    await ExpectDomainFailure(()=>documents.ChangeStateAsync(executionKind,"sql-test-budget-only",inviteExecution.Id,invitation.BudgetId,new byte[8],new byte[8],Guid.NewGuid(),"CANCELADA",new DateOnly(2026,11,5),"Responsable","Prueba",CancellationToken.None),"permiso","Cambio de estado documental exige permiso de edición SQL");
    var guideInvitation=await Save(connection,Payload([Manual()],attention:"INVITACION"));
    guideInvitation=await State(connection,guideInvitation,"ENVIADO");
    var guideAcceptanceOperation=Guid.NewGuid();
    var guideAcceptancePayload=new {FechaActa="2026-11-06",NumeroActa="AUTO-INVITACION"};
    await Execute(connection,"CREATE OR ALTER TRIGGER dbo.trg_TestAutoAcceptanceFailure ON dbo.EventoAuditoria AFTER INSERT AS BEGIN SET NOCOUNT ON; IF EXISTS(SELECT 1 FROM inserted AS e WHERE e.Entidad=N'PresupuestoRevision' AND e.EstadoNuevo=N'APROBADO') THROW 51999,N'Fallo aceptación automática.',1; END;");
    try
    {
        await ExpectFailure(()=>SaveDocument(connection,"GUIA",guideInvitation.BudgetId,guideInvitation.HeaderVersion,guideAcceptancePayload,guideAcceptanceOperation),"Fallo aceptación automática","Error al auditar aceptación revierte todo el registro de guía");
        Check(Convert.ToString(await Scalar(connection,$"SELECT Estado FROM dbo.PresupuestoRevision WHERE RevisionId={guideInvitation.RevisionId}"))=="ENVIADO"&&Convert.ToInt64(await Scalar(connection,$"SELECT COUNT_BIG(1) FROM dbo.GuiaConformidad WHERE PresupuestoId={guideInvitation.BudgetId}"))==0,"Fallo no deja presupuesto aceptado ni guía parcial");
    }
    finally {await Execute(connection,"DROP TRIGGER dbo.trg_TestAutoAcceptanceFailure;");}
    var autoGuideId=await SaveDocument(connection,"GUIA",guideInvitation.BudgetId,guideInvitation.HeaderVersion,guideAcceptancePayload,guideAcceptanceOperation);
    Check(autoGuideId>0&&Convert.ToString(await Scalar(connection,$"SELECT Estado FROM dbo.PresupuestoRevision WHERE RevisionId={guideInvitation.RevisionId}"))=="APROBADO","Guía de invitación enviada acepta presupuesto sin OC previa");
    Check(await SaveDocument(connection,"GUIA",guideInvitation.BudgetId,guideInvitation.HeaderVersion,guideAcceptancePayload,guideAcceptanceOperation)==autoGuideId&&Convert.ToInt64(await Scalar(connection,$"SELECT COUNT_BIG(1) FROM dbo.EventoAuditoria WHERE Entidad=N'PresupuestoRevision' AND RegistroId=N'{guideInvitation.RevisionId}' AND EstadoNuevo=N'APROBADO'"))==1,"Reintento de guía no duplica aceptación ni auditoría");
    var regularization=await Save(connection,Payload([Manual(25)]));
    var regularizationOc=await documents.SaveAsync(ocKind,actor,regularization.BudgetId,regularization.HeaderVersion,JsonSerializer.Serialize(new{RevisionId=regularization.RevisionId,NumeroOCCliente="EMERG-0009A",FechaEmisionOC="2026-11-03",FechaRecepcionOC="2026-11-03"}),Guid.NewGuid(),null,null,CancellationToken.None);
    regularization=regularization with{HeaderVersion=(byte[])(await Scalar(connection,$"SELECT VersionFila FROM dbo.Presupuesto WHERE PresupuestoId={regularization.BudgetId}"))!,RevisionVersion=(byte[])(await Scalar(connection,$"SELECT VersionFila FROM dbo.PresupuestoRevision WHERE RevisionId={regularization.RevisionId}"))!};
    await ExpectFailure(()=>Save(connection,Payload([Manual(99)]),regularization),"no admite cambios","Emergencia con OC mantiene bloqueados contenido e importes");
    Check(Convert.ToString(await Scalar(connection,$"SELECT Estado FROM dbo.PresupuestoRevision WHERE RevisionId={regularization.RevisionId}"))=="APROBADO","OC de emergencia acepta automáticamente sin regularización manual");
    Check(Convert.ToDecimal(await Scalar(connection,$"SELECT Total FROM dbo.PresupuestoRevision WHERE RevisionId={regularization.RevisionId}"))==59m,"Emergencia se regulariza después de OC sin recalcular importes");
    await ExpectFailure(()=>State(connection,regularization,"CANCELADO"),"OC no permite","Presupuesto aprobado con OC no se cancela desde regularización");
    await ExpectFailure(()=>DocumentState("guias",guideId,emergency.BudgetId,"ACEPTADA"),"no utiliza estados","Guía rechaza cambios de estado en SQL");
    async Task EditGuideForTest(string observation)
    {
        var original=await documents.GetAsync(guideKind,actor,guideId,null,1,CancellationToken.None);
        await documents.SaveAsync(guideKind,actor,emergency.BudgetId,Convert.FromBase64String(original!.Data.GetProperty("VersionPadre").GetString()!),JsonSerializer.Serialize(new{FechaActa="2026-11-03",NumeroActa="GUIA-0007B",Observacion=observation,MotivoCambio="Corrección de datos de guía"}),Guid.NewGuid(),guideId,Convert.FromBase64String(original.Data.GetProperty("Documento").GetProperty("VersionFila").GetString()!),CancellationToken.None);
    }
    var priorGuideObservation=Convert.ToString(await Scalar(connection,$"SELECT Observacion FROM dbo.GuiaConformidad WHERE GuiaId={guideId}"));
    await Execute(connection,"""
        CREATE OR ALTER TRIGGER dbo.trg_TestAuditFailure ON dbo.EventoAuditoria AFTER INSERT AS
        BEGIN
            SET NOCOUNT ON;
            IF EXISTS(SELECT 1 FROM inserted AS evento WHERE evento.Entidad=N'GuiaConformidad' AND evento.Accion=N'ACTUALIZAR')
                THROW 51999,N'Fallo de auditoría provocado para prueba aislada.',1;
        END;
        """);
    try
    {
        await ExpectFailure(()=>EditGuideForTest("Prueba rollback"),"Fallo de auditoría","Fallo de auditoría revierte edición de guía");
        Check(Convert.ToString(await Scalar(connection,$"SELECT Observacion FROM dbo.GuiaConformidad WHERE GuiaId={guideId}"))==priorGuideObservation,"Fallo de edición conserva datos previos de guía");
    }
    finally{await Execute(connection,"DROP TRIGGER dbo.trg_TestAuditFailure;");}
    for(var eventIndex=0;eventIndex<25;eventIndex++)await EditGuideForTest("Cambio "+eventIndex);
    var fullHistory=await repository.GetAsync(actor,emergency.BudgetId,null,1,CancellationToken.None);
    var nextHistory=await repository.GetAsync(actor,emergency.BudgetId,null,2,CancellationToken.None)??throw new InvalidOperationException("Falta segunda página del expediente.");
    Check(fullHistory!.Audit.Count==20&&fullHistory.AuditTotal>20&&nextHistory!.Audit.Count>0&&!fullHistory.Audit.Select(row=>row.Id).Intersect(nextHistory.Audit.Select(row=>row.Id)).Any(),"Expediente consolida eventos repetidos con páginas SQL sin duplicación");
    var limitedHistory=await repository.GetAsync("sql-test-budget-only",emergency.BudgetId,null,1,CancellationToken.None);
    Check(limitedHistory!.Audit.All(row=>row.Entity is "Presupuesto" or "PresupuestoRevision")&&limitedHistory.AuditTotal<fullHistory.AuditTotal,"Historial del presupuesto no revela documentos sin permiso específico");
    var oldestHistory=await repository.GetAsync(actor,emergency.BudgetId,null,3,CancellationToken.None);
    Check(fullHistory.Audit.Concat(nextHistory.Audit).Concat(oldestHistory!.Audit).Any(row=>row.Entity=="ArchivoAdjunto")&&fullHistory.Audit.Concat(nextHistory.Audit).Concat(oldestHistory.Audit).Any(row=>row.Entity=="Ejecucion"),"Historial consolidado incluye archivos y ejecución del mismo expediente");
    var updatedBudgetPage=await repository.ListAsync(actor,DateTimeOffset.UtcNow.AddDays(-1),DateTimeOffset.UtcNow.AddDays(1),emergency.Number,1,null,CancellationToken.None);
    var latestUpdated=updatedBudgetPage.Rows.Single(row=>row.BudgetId==emergency.BudgetId);
    Check(latestUpdated.UpdatedAt.HasValue&&latestUpdated.UpdatedAt.Value.UtcDateTime==(DateTime)(await Scalar(connection,$"SELECT FechaActualizacion FROM dbo.Presupuesto WHERE PresupuestoId={emergency.BudgetId}"))!,"Listado presupuesto muestra la última actualización documental");
    Check(latestUpdated.ExecutionId==execution.Id&&latestUpdated.GuideId==guideId,"Listado SQL de presupuestos identifica ejecución y guía ya registradas");
    Check(await repository.GetAsync(actor,999999,null,1,CancellationToken.None) is null,"Consulta SQL de presupuesto inexistente devuelve vacío sin error");
    var noOrderInvitation=await Save(connection,Payload([Manual()],attention:"INVITACION"));
    var unapprovedExecution=await documents.SaveAsync(executionKind,actor,noOrderInvitation.BudgetId,noOrderInvitation.HeaderVersion,JsonSerializer.Serialize(new{Responsable="Técnico",FechaProgramada="2026-11-03"}),Guid.NewGuid(),null,null,CancellationToken.None);
    await DocumentState("ejecucion",unapprovedExecution.Id,noOrderInvitation.BudgetId,"PROGRAMADA");
    await ExpectFailure(()=>DocumentState("ejecucion",unapprovedExecution.Id,noOrderInvitation.BudgetId,"EN_EJECUCION"),"invitación requiere OC","Invitación no inicia ejecución sin OC aunque tenga programación");
    var immediateEmergency=await Save(connection,Payload([]));
    var immediateExecution=await documents.SaveAsync(executionKind,actor,immediateEmergency.BudgetId,immediateEmergency.HeaderVersion,JsonSerializer.Serialize(new{Responsable="Técnico de emergencia"}),Guid.NewGuid(),null,null,CancellationToken.None);
    await DocumentState("ejecucion",immediateExecution.Id,immediateEmergency.BudgetId,"EN_EJECUCION");
    Check(Convert.ToString(await Scalar(connection,$"SELECT Estado FROM dbo.Ejecucion WHERE EjecucionId={immediateExecution.Id}"))=="EN_EJECUCION","Emergencia puede iniciar directamente sin programación ni OC");
    await ExpectFailure(()=>DocumentState("oc-cliente",ocId,invitation.BudgetId,"CANCELADA"),"Retira la confirmación","Cancelación OC exige resolver confirmación y guía vigentes");
    await DocumentState("confirmaciones",replacementConfirmation.Id,ocId,"RETIRADA");
    await DocumentState("oc-cliente",ocId,invitation.BudgetId,"CANCELADA");
    await DocumentState("oc-cliente",ocId,invitation.BudgetId,"REGISTRADA");
    Check(Convert.ToInt64(await Scalar(connection,$"SELECT COUNT_BIG(1) FROM dbo.OrdenCompraCliente WHERE PresupuestoId={invitation.BudgetId}"))==1,"Cancelar y reactivar OC conserva el mismo registro único");
    var invalidUploadOperation=Guid.NewGuid();
    var currentGuide=await documents.GetAsync(guideKind,actor,inviteGuide.Id,null,1,CancellationToken.None);
    var beforeInvalidFiles=Convert.ToInt64(await Scalar(connection,"SELECT COUNT_BIG(1) FROM dbo.ArchivoAdjunto"));
    await ExpectDomainFailure(()=>documents.SaveAsync(guideKind,actor,invitation.BudgetId,Convert.FromBase64String(currentGuide!.Data.GetProperty("VersionPadre").GetString()!),JsonSerializer.Serialize(new{FechaActa="2026-11-04",NumeroActa="INV-0008C",MotivoCambio="Prueba rechazo de imagen",Adjunto=new{NombreOriginal="imagen.png",ClaveAlmacenamiento=invalidUploadOperation.ToString("N")+".png",TipoContenido="image/png",TamanoBytes=20,SHA256=new string('B',64)}}),invalidUploadOperation,inviteGuide.Id,Convert.FromBase64String(currentGuide!.Data.GetProperty("Documento").GetProperty("VersionFila").GetString()!),CancellationToken.None),"solo admite escaneado PDF","SQL también rechaza imagen en guía aunque se omita la validación web");
    Check(Convert.ToInt64(await Scalar(connection,"SELECT COUNT_BIG(1) FROM dbo.ArchivoAdjunto"))==beforeInvalidFiles,"Adjunto rechazado no deja metadatos parciales");
    async Task<long> MarkInvoice(long order,byte[] version,Guid operation,long? reference=null,byte[]? referenceVersion=null,bool withdraw=false)
    {
        await using var command=new SqlCommand("dbo.spu_Facturado_Guardar",connection){CommandType=CommandType.StoredProcedure};
        command.Parameters.AddWithValue("@UsuarioId",actor);command.Parameters.AddWithValue("@CorrelacionId",operation);command.Parameters.AddWithValue("@OrdenCompraId",order);
        command.Parameters.Add("@VersionOC",SqlDbType.Binary,8).Value=version;command.Parameters.Add("@FechaDocumento",SqlDbType.Date).Value=new DateTime(2026,11,3);
        command.Parameters.AddWithValue("@CodigoTipoDocumento","01");command.Parameters.AddWithValue("@Serie","F001");command.Parameters.AddWithValue("@NumeroDocumento","000012A");
        command.Parameters.Add("@ReferenciaId",SqlDbType.BigInt).Value=(object?)reference??DBNull.Value;command.Parameters.Add("@VersionReferencia",SqlDbType.Binary,8).Value=(object?)referenceVersion??DBNull.Value;
        command.Parameters.AddWithValue("@Retirar",withdraw);command.Parameters.AddWithValue("@Motivo","Prueba aislada de referencia externa");
        return Convert.ToInt64(await command.ExecuteScalarAsync());
    }
    var markOrder=regularizationOc.Id;
    var markVersion=(byte[])(await Scalar(connection,$"SELECT VersionFila FROM dbo.OrdenCompraCliente WHERE OrdenCompraId={markOrder}"))!;
    await ExpectFailure(()=>MarkInvoice(markOrder,markVersion,Guid.NewGuid()),"confirmación del cliente","Facturado exige confirmación del cliente de la OC");
    await documents.SaveAsync(GestionServicios.Application.Followup.DocumentKind.Resolve("confirmaciones"),actor,markOrder,markVersion,JsonSerializer.Serialize(new{FechaConfirmacion="2026-11-03",NumeroConfirmacion="CONF-FACT-01"}),Guid.NewGuid(),null,null,CancellationToken.None);
    markVersion=(byte[])(await Scalar(connection,$"SELECT VersionFila FROM dbo.OrdenCompraCliente WHERE OrdenCompraId={markOrder}"))!;
    var completionParent=(byte[])(await Scalar(connection,$"SELECT VersionFila FROM dbo.Presupuesto WHERE PresupuestoId={regularization.BudgetId}"))!;
    var completeWork=await documents.SaveAsync(executionKind,actor,regularization.BudgetId,completionParent,JsonSerializer.Serialize(new{Responsable="Técnico de cierre"}),Guid.NewGuid(),null,null,CancellationToken.None);
    await DocumentState("ejecucion",completeWork.Id,regularization.BudgetId,"EN_EJECUCION");
    await DocumentState("ejecucion",completeWork.Id,regularization.BudgetId,"TERMINADA");
    completionParent=(byte[])(await Scalar(connection,$"SELECT VersionFila FROM dbo.Presupuesto WHERE PresupuestoId={regularization.BudgetId}"))!;
    var completeGuide=await documents.SaveAsync(guideKind,actor,regularization.BudgetId,completionParent,JsonSerializer.Serialize(new{FechaActa="2026-11-03",NumeroActa="FACT-GUIA-01"}),Guid.NewGuid(),null,null,CancellationToken.None);
    var markOperation=Guid.NewGuid();var markId=await MarkInvoice(markOrder,markVersion,markOperation);
    Check(Convert.ToString(await Scalar(connection,$"SELECT CodigoTipoDocumento+N'|'+Serie+N'|'+NumeroDocumento FROM dbo.ReferenciaComprobanteExterno WHERE ReferenciaId={markId}"))=="01|F001|000012A","Facturado guarda código textual y número externo sin perder ceros");
    Check(await MarkInvoice(markOrder,markVersion,markOperation)==markId&&Convert.ToInt64(await Scalar(connection,$"SELECT COUNT_BIG(1) FROM dbo.ReferenciaComprobanteExterno WHERE OrdenCompraId={markOrder}"))==1,"Reintento Facturado no duplica referencia");
    markVersion=(byte[])(await Scalar(connection,$"SELECT VersionFila FROM dbo.OrdenCompraCliente WHERE OrdenCompraId={markOrder}"))!;
    await ExpectFailure(()=>MarkInvoice(markOrder,markVersion,Guid.NewGuid()),"ya está marcada","OC no admite dos referencias vigentes");
    var markReferenceVersion=(byte[])(await Scalar(connection,$"SELECT VersionFila FROM dbo.ReferenciaComprobanteExterno WHERE ReferenciaId={markId}"))!;
    await MarkInvoice(markOrder,markVersion,Guid.NewGuid(),markId,markReferenceVersion,true);
    Check(Convert.ToInt64(await Scalar(connection,$"SELECT COUNT_BIG(1) FROM dbo.ReferenciaComprobanteExterno WHERE ReferenciaId={markId} AND Vigente=0 AND MotivoRetiro IS NOT NULL"))==1&&Convert.ToInt64(await Scalar(connection,$"SELECT COUNT_BIG(1) FROM dbo.EventoAuditoria WHERE Entidad=N'ReferenciaComprobanteExterno' AND RegistroId=N'{markId}'"))==2,"Retiro Facturado conserva referencia y auditoría con motivo");
    var invoiceRepository=new GestionServicios.Infrastructure.Billing.SqlInvoiceMarkRepository(new GestionServicios.Infrastructure.SqlConnectionFactory(localConnection+"Initial Catalog="+testDatabase+";"));
    var retiredReference=await invoiceRepository.GetAsync(actor,markOrder,1,CancellationToken.None);
    Check(retiredReference!.Value.GetProperty("Referencia").ValueKind==JsonValueKind.Null&&retiredReference.Value.GetProperty("Historial").GetArrayLength()==1,"Repositorio Facturado recupera historial retirado sin marcado vigente");
    var freshOrderVersion=Convert.FromBase64String(retiredReference.Value.GetProperty("VersionFila").GetString()!);
    var newMark=await invoiceRepository.SaveAsync(actor,new(markOrder,freshOrderVersion,Guid.NewGuid(),new DateOnly(2026,11,3),"03","B001","000001B",null,null,false,null),CancellationToken.None);
    var savedMark=await invoiceRepository.GetAsync(actor,markOrder,1,CancellationToken.None);
    Check(newMark>0&&savedMark!.Value.GetProperty("Referencia").GetProperty("CodigoTipoDocumento").GetString()=="03","Repositorio ADO.NET registra boleta con código textual 03");
    var savedReference=savedMark!.Value.GetProperty("Referencia");
    await invoiceRepository.SaveAsync(actor,new(markOrder,Convert.FromBase64String(savedMark.Value.GetProperty("VersionFila").GetString()!),Guid.NewGuid(),new DateOnly(2026,11,4),"03","B001","000002C",newMark,Convert.FromBase64String(savedReference.GetProperty("VersionFila").GetString()!),false,"Corrección de número externo"),CancellationToken.None);
    var correctedReference=await invoiceRepository.GetAsync(actor,markOrder,1,CancellationToken.None);
    Check(correctedReference!.Value.GetProperty("Auditoria").GetArrayLength()==4&&correctedReference.Value.GetProperty("Auditoria")[0].GetProperty("CambiosJson").GetString()!.Contains("000001B"),"Corrección Facturado conserva valor anterior en auditoría consultable");
    var invoicedList=await documents.ListAsync(ocKind,actor,DateTimeOffset.UtcNow.AddDays(-1),DateTimeOffset.UtcNow.AddDays(1),"EMERG-0009A",1,null,null,CancellationToken.None);
    Check(invoicedList.Rows.Single().GetProperty("Facturado").GetProperty("NumeroDocumento").GetString()=="000002C","Listado OC muestra la referencia Facturado vigente tras corregir");
    await Execute(connection,"INSERT INTO dbo.RolPermiso(RoleId,CodigoPermiso) VALUES(N'sql-test-budget-role',N'OC_CLIENTE.CONSULTAR');");
    var restrictedInvoiceList=await documents.ListAsync(ocKind,"sql-test-budget-only",DateTimeOffset.UtcNow.AddDays(-1),DateTimeOffset.UtcNow.AddDays(1),"EMERG-0009A",1,null,null,CancellationToken.None);
    Check(restrictedInvoiceList.Rows.Single().GetProperty("Facturado").ValueKind==JsonValueKind.Null,"Consulta OC sin permiso Facturado no revela datos del comprobante");
    async Task<JsonElement> Dashboard(string user,DateTime from,DateTime until)
    {
        await using var command=new SqlCommand("dbo.spu_Dashboard_Obtener",connection){CommandType=CommandType.StoredProcedure};command.Parameters.AddWithValue("@UsuarioId",user);command.Parameters.Add("@Desde",SqlDbType.DateTime2).Value=from;command.Parameters.Add("@HastaExclusiva",SqlDbType.DateTime2).Value=until;
        return JsonSerializer.Deserialize<JsonElement>(Convert.ToString(await command.ExecuteScalarAsync())!);
    }
    async Task<(JsonElement[] Rows,long Total)> Report(string code,string user=actor)
    {
        await using var command=new SqlCommand("dbo.spu_Reporte_Listar",connection){CommandType=CommandType.StoredProcedure};command.Parameters.AddWithValue("@UsuarioId",user);command.Parameters.AddWithValue("@Reporte",code);command.Parameters.Add("@Desde",SqlDbType.DateTime2).Value=DateTime.UtcNow.AddDays(-1);command.Parameters.Add("@HastaExclusiva",SqlDbType.DateTime2).Value=DateTime.UtcNow.AddDays(1);
        await using var reader=await command.ExecuteReaderAsync();await reader.ReadAsync();var rows=JsonSerializer.Deserialize<JsonElement[]>(reader.GetString(0))!;await reader.NextResultAsync();await reader.ReadAsync();return(rows,reader.GetInt64(0));
    }
    var dashboard=await Dashboard(actor,DateTime.UtcNow.AddDays(-1),DateTime.UtcNow.AddDays(1));
    Check(dashboard.GetProperty("Presupuestos").GetInt64()==Convert.ToInt64(await Scalar(connection,"SELECT COUNT_BIG(1) FROM dbo.Presupuesto")),"Dashboard cuenta un presupuesto por expediente sin duplicar revisiones");
    Check(dashboard.GetProperty("Ordenes").GetInt64()==Convert.ToInt64(await Scalar(connection,"SELECT COUNT_BIG(1) FROM dbo.OrdenCompraCliente")),"Dashboard concilia OC sin multiplicar confirmaciones históricas");
    Check(dashboard.GetProperty("ImportesPresupuesto").EnumerateArray().Select(row=>row.GetProperty("CodigoMoneda").GetString()).Distinct().Count()==2,"Dashboard separa importes PEN y USD");
    foreach(var code in new[]{"GANADOS","OC_COMPLETADAS","OC_CON_ACTA","OC_PENDIENTES_FACTURA"})
    {
        var report=await Report(code);Check(report.Rows.Length<=20&&report.Rows.Select(row=>row.GetProperty("Id").GetInt64()).Distinct().Count()==report.Rows.Length,"Reporte paginado sin documentos duplicados: "+code);
    }
    var pendingInvoice=await Report("OC_PENDIENTES_FACTURA");
    Check(!pendingInvoice.Rows.Any(row=>row.GetProperty("Id").GetInt64()==markOrder),"Reporte pendientes excluye OC marcada con referencia vigente");
    await ExpectFailure(()=>Dashboard("sql-test-budget-only",DateTime.UtcNow.AddDays(-1),DateTime.UtcNow.AddDays(1)),"permiso","Indicadores exigen permiso de reportes");
    var curveBudget=await Save(connection,Payload([Manual()]));
    await Execute(connection,$"UPDATE dbo.Presupuesto SET FechaRegistro='2026-11-04T02:00:00' WHERE PresupuestoId={curveBudget.BudgetId};");
    var peruvianDashboard=await Dashboard(actor,new DateTime(2026,11,3,5,0,0),new DateTime(2026,11,4,5,0,0));
    Check(peruvianDashboard.GetProperty("Movimientos").EnumerateArray().Any(row=>row.GetProperty("Fecha").GetString()=="2026-11-03"),"Movimientos agrupan por fecha Perú aunque el registro UTC sea del día siguiente");
    var reportingRepository=new GestionServicios.Infrastructure.Reporting.SqlReportingRepository(new GestionServicios.Infrastructure.SqlConnectionFactory(localConnection+"Initial Catalog="+testDatabase+";"));
    var reportFromRepository=await reportingRepository.ReportAsync(actor,"OC_PENDIENTES_FACTURA",DateTimeOffset.UtcNow.AddDays(-1),DateTimeOffset.UtcNow.AddDays(1),1,CancellationToken.None);
    Check(reportFromRepository.Rows.Count<=20&&reportFromRepository.Total>=reportFromRepository.Rows.Count&&reportFromRepository.Amounts.All(row=>row.Currency is "PEN" or "USD"),"Repositorio reportes recupera filas, conteo y totales por moneda reales");
    var realDashboard=await reportingRepository.DashboardAsync(actor,DateTimeOffset.UtcNow.AddDays(-1),DateTimeOffset.UtcNow.AddDays(1),CancellationToken.None);
    Check(realDashboard.GetProperty("VerOC").GetBoolean()&&realDashboard.GetProperty("Ordenes").GetInt64()>0,"Repositorio de indicadores recupera las fuentes autorizadas");
    for(var reportIndex=0;reportIndex<25;reportIndex++)
    {
        var reportBudget=await Save(connection,Payload([Manual(1)]));reportBudget=await State(connection,reportBudget,"ENVIADO");await State(connection,reportBudget,"APROBADO");
    }
    var reportFirst=await reportingRepository.ReportAsync(actor,"GANADOS",DateTimeOffset.UtcNow.AddDays(-1),DateTimeOffset.UtcNow.AddDays(1),1,CancellationToken.None);
    var reportSecond=await reportingRepository.ReportAsync(actor,"GANADOS",DateTimeOffset.UtcNow.AddDays(-1),DateTimeOffset.UtcNow.AddDays(1),2,CancellationToken.None);
    Check(reportFirst.Rows.Count==20&&reportSecond.Rows.Count>0&&reportFirst.Total==reportSecond.Total&&!reportFirst.Rows.Select(row=>row.GetProperty("Id").GetInt64()).Intersect(reportSecond.Rows.Select(row=>row.GetProperty("Id").GetInt64())).Any(),"Reporte real pagina más de veinte documentos sin solapamiento");
    Check(reportFirst.Amounts.SequenceEqual(reportSecond.Amounts)&&reportFirst.Amounts.Sum(row=>row.Count)==reportFirst.Total,"Totales por moneda corresponden a toda la consulta y no cambian de página");
    var queueFirst=await reportingRepository.DashboardAsync(actor,DateTimeOffset.UtcNow.AddDays(-1),DateTimeOffset.UtcNow.AddDays(1),CancellationToken.None);
    var queueSecond=await reportingRepository.DashboardAsync(actor,DateTimeOffset.UtcNow.AddDays(-1),DateTimeOffset.UtcNow.AddDays(1),CancellationToken.None,queuePage:2);
    Check(queueFirst.GetProperty("VerPendientes").GetBoolean()&&queueFirst.GetProperty("Pendientes").GetArrayLength()==20&&queueSecond.GetProperty("Pendientes").GetArrayLength()>0&&!queueFirst.GetProperty("Pendientes").EnumerateArray().Select(row=>row.GetProperty("Id").GetInt64()).Intersect(queueSecond.GetProperty("Pendientes").EnumerateArray().Select(row=>row.GetProperty("Id").GetInt64())).Any(),"Bandeja histórica SQL pagina veinte expedientes sin repetirlos");
    Check(queueFirst.GetProperty("TotalPendientes").GetInt64()>queueFirst.GetProperty("PendientesPeriodo").GetInt64(),"Pendientes acumulados incluyen expedientes anteriores al período");
    var absentClient=await reportingRepository.DashboardAsync(actor,DateTimeOffset.UtcNow.AddDays(-1),DateTimeOffset.UtcNow.AddDays(1),CancellationToken.None,client:"CLIENTE_NO_EXISTE_73");
    var absentStore=await reportingRepository.DashboardAsync(actor,DateTimeOffset.UtcNow.AddDays(-1),DateTimeOffset.UtcNow.AddDays(1),CancellationToken.None,store:"TIENDA_NO_EXISTE_73");
    Check(absentClient.GetProperty("Presupuestos").GetInt64()==0&&absentClient.GetProperty("Ordenes").GetInt64()==0&&absentClient.GetProperty("Movimientos").GetArrayLength()==0&&absentClient.GetProperty("TotalPendientes").GetInt64()==0,"Filtro cliente se aplica a cifras, movimientos y pendientes");
    Check(absentStore.GetProperty("Presupuestos").GetInt64()==0&&absentStore.GetProperty("Ordenes").GetInt64()==0&&absentStore.GetProperty("Movimientos").GetArrayLength()==0&&absentStore.GetProperty("TotalPendientes").GetInt64()==0,"Filtro tienda se aplica a cifras, movimientos y pendientes");
    Check(queueFirst.GetProperty("Movimientos").EnumerateArray().Sum(row=>row.GetProperty("Presupuestos").GetInt64())==queueFirst.GetProperty("Presupuestos").GetInt64()&&queueFirst.GetProperty("Movimientos").EnumerateArray().Sum(row=>row.GetProperty("Ordenes").GetInt64())==queueFirst.GetProperty("Ordenes").GetInt64(),"Curvas actuales concilian registros de presupuestos y OC del período");
    Check(queueFirst.GetProperty("Movimientos").EnumerateArray().Sum(row=>row.GetProperty("Confirmaciones").GetInt64())==Convert.ToInt64(await Scalar(connection,"SELECT COUNT_BIG(1) FROM dbo.ConfirmacionFacturacion AS confirmacion INNER JOIN dbo.OrdenCompraCliente AS oc ON oc.OrdenCompraId=confirmacion.OrdenCompraId WHERE confirmacion.Vigente=1 AND oc.Estado<>N'CANCELADA'")),"Curva confirmación cuenta registros vigentes y no primeros eventos históricos");
    Check(queueFirst.GetProperty("Movimientos").EnumerateArray().Sum(row=>row.GetProperty("Marcados").GetInt64())==Convert.ToInt64(await Scalar(connection,"SELECT COUNT_BIG(1) FROM dbo.ReferenciaComprobanteExterno AS referencia INNER JOIN dbo.OrdenCompraCliente AS oc ON oc.OrdenCompraId=referencia.OrdenCompraId WHERE referencia.Vigente=1 AND oc.Estado<>N'CANCELADA'")),"Curva Facturado cuenta referencias vigentes, sin sumar retiradas");
    await Execute(connection,"INSERT INTO dbo.RolPermiso(RoleId,CodigoPermiso) VALUES(N'sql-test-budget-role',N'REPORTES.CONSULTAR');");
    var restrictedQueue=await reportingRepository.DashboardAsync("sql-test-budget-only",DateTimeOffset.UtcNow.AddDays(-1),DateTimeOffset.UtcNow.AddDays(1),CancellationToken.None);
    Check(!restrictedQueue.GetProperty("VerPendientes").GetBoolean()&&restrictedQueue.GetProperty("Pendientes").GetArrayLength()==0,"Bandeja no revela guía sin su permiso específico");
    var absentBudgetList=await repository.ListAsync(actor,DateTimeOffset.UtcNow.AddDays(-1),DateTimeOffset.UtcNow.AddDays(1),null,1,null,CancellationToken.None,client:"CLIENTE_NO_EXISTE_73");
    var absentOrderList=await documents.ListAsync(ocKind,actor,DateTimeOffset.UtcNow.AddDays(-1),DateTimeOffset.UtcNow.AddDays(1),null,1,null,null,CancellationToken.None,store:"TIENDA_NO_EXISTE_73");
    var absentReport=await reportingRepository.ReportAsync(actor,"GANADOS",DateTimeOffset.UtcNow.AddDays(-1),DateTimeOffset.UtcNow.AddDays(1),1,CancellationToken.None,client:"CLIENTE_NO_EXISTE_73");
    Check(absentBudgetList.Total==0&&absentOrderList.Total==0&&absentReport.Total==0&&absentReport.Amounts.Count==0,"Navegación a listados/reportes aplica filtros comerciales antes de conteo y totales");
    var approvedBudgetList=await repository.ListAsync(actor,DateTimeOffset.UtcNow.AddDays(-1),DateTimeOffset.UtcNow.AddDays(1),null,1,"APROBADO",CancellationToken.None);
    Check(approvedBudgetList.Rows.Count>0&&approvedBudgetList.Rows.All(row=>row.State=="APROBADO"),"Acceso desde columnas del dashboard conserva estado en consulta SQL");
    var clientDashboard=await reportingRepository.DashboardAsync(actor,DateTimeOffset.UtcNow.AddDays(-1),DateTimeOffset.UtcNow.AddDays(1),CancellationToken.None,client:listed.Rows[0].ClientName,store:listed.Rows[0].StoreName);
    Check(clientDashboard.GetProperty("Presupuestos").GetInt64()>0&&clientDashboard.GetProperty("Presupuestos").GetInt64()==queueFirst.GetProperty("Presupuestos").GetInt64(),"Cliente y tienda combinados recuperan los expedientes coincidentes del conjunto de prueba");
    foreach(var selection in new[]{(Code:"SIN_CONFIRMAR",Kind:"OC_PENDIENTES_FACTURA",Metric:"SinConfirmar"),(Code:"CONFIRMADAS",Kind:"OC_PENDIENTES_FACTURA",Metric:"ConfirmadasSinMarcar"),(Code:"LISTAS",Kind:"OC_PENDIENTES_FACTURA",Metric:"ListasParaMarcar"),(Code:"FACTURADAS",Kind:"OC_COMPLETADAS",Metric:"Facturadas")})
    {
        var selectedReport=await reportingRepository.ReportAsync(actor,selection.Kind,DateTimeOffset.UtcNow.AddDays(-1),DateTimeOffset.UtcNow.AddDays(1),1,CancellationToken.None,situation:selection.Code);
        Check(selectedReport.Total==queueFirst.GetProperty(selection.Metric).GetInt64(),"Acceso rápido concilia exactamente su indicador: "+selection.Code);
    }
    await ExpectDomainFailure(()=>reportingRepository.ReportAsync(actor,"GANADOS",DateTimeOffset.UtcNow.AddDays(-1),DateTimeOffset.UtcNow.AddDays(1),1,CancellationToken.None,situation:"FACTURADAS"),"seguimiento","SQL rechaza aplicar seguimiento OC a presupuestos");

    var masters=new GestionServicios.Infrastructure.Maintenance.SqlMaintenanceRepository(new GestionServicios.Infrastructure.SqlConnectionFactory(localConnection+"Initial Catalog="+testDatabase+";"));
    var clientDefinition=GestionServicios.Application.Maintenance.MaintenanceDefinitions.Find("clientes")!;
    async Task<long> MasterSave(string key,object payload,long? id=null,byte[]? version=null)=>await masters.SaveAsync(GestionServicios.Application.Maintenance.MaintenanceDefinitions.Find(key)!,actor,id,version,JsonSerializer.Serialize(payload),CancellationToken.None);
    var autoClient=await MasterSave("clientes",new{Codigo="CLIENTE_ELEGIDO",TipoEntidad="JURIDICA",TipoIdentificacion="RUC",NumeroIdentificacion="20123456789",RazonSocial="Cliente automático",CodigoSAP="000SAP-A/01",Activo=true});
    var autoRecord=(await masters.GetAsync(clientDefinition,actor,autoClient))!;var generatedCode=autoRecord.Text("Codigo");
    Check(generatedCode.StartsWith("CLI-")&&generatedCode!="CLIENTE_ELEGIDO","Código de cliente generado en SQL ignora el enviado por el navegador");
    await MasterSave("clientes",new{Codigo="CODIGO_CAMBIADO",TipoEntidad="JURIDICA",TipoIdentificacion="RUC",NumeroIdentificacion="20123456789",RazonSocial="Cliente actualizado",CodigoSAP="000SAP-B/02",Activo=true},autoClient,Convert.FromBase64String(autoRecord.Text("VersionFila")));
    Check((await masters.GetAsync(clientDefinition,actor,autoClient))!.Text("Codigo")==generatedCode,"Código interno existente se conserva aunque el JSON intente cambiarlo");
    var sapClient=(await masters.GetAsync(clientDefinition,actor,autoClient))!;
    Check(autoRecord.Text("CodigoSAP")=="000SAP-A/01"&&sapClient.Text("CodigoSAP")=="000SAP-B/02","Código SAP alfanumérico por cliente conserva ceros y se actualiza con auditoría");
    var sapBudgets=new GestionServicios.Infrastructure.Budgets.SqlBudgetRepository(new GestionServicios.Infrastructure.SqlConnectionFactory(localConnection+"Initial Catalog="+testDatabase+";"));
    var sapDefaults=await sapBudgets.DefaultsAsync(actor,new DateOnly(2026,11,5),autoClient,null,CancellationToken.None);
    Check(sapDefaults.GetProperty("Proveedor").GetProperty("CodigoSAPProveedor").GetString()=="000SAP-B/02","Nuevo presupuesto toma SAP del cliente aunque no existan datos HYT por cliente");
    var naturalId=await MasterSave("clientes",new{TipoEntidad="NATURAL",TipoIdentificacion="DNI",NumeroIdentificacion="12345678",Nombres="Ana",Apellidos="Pérez",RazonSocial="Descartar",Activo=true});
    var storeWithClient=await MasterSave("tiendas",new{ClienteId=autoClient,Nombre="Tienda de listado",Activo=true});
    var contactWithClient=await MasterSave("contactos",new{ClienteId=autoClient,Nombre="Contacto de listado",Activo=true});
    foreach(var relatedKey in new[]{"tiendas","contactos"})
    {
        var relatedList=await masters.ListAsync(GestionServicios.Application.Maintenance.MaintenanceDefinitions.Find(relatedKey)!,actor,generatedCode,1);
        Check(relatedList.Rows.Any(row=>row.ClientCode==generatedCode&&row.ClientName=="Cliente actualizado")&&relatedList.Total==1,"Listado y búsqueda por código de cliente en "+relatedKey);
    }
    var naturalContact=await MasterSave("contactos",new{ClienteId=naturalId,Nombre="Contacto natural",Activo=true});
    var naturalContacts=await masters.ListAsync(GestionServicios.Application.Maintenance.MaintenanceDefinitions.Find("contactos")!,actor,"Ana Pérez",1);
    Check(naturalContacts.Rows.Any(row=>row.Id==naturalContact&&row.ClientName=="Ana Pérez"),"Contacto muestra y busca nombres/apellidos del cliente natural");
    var naturalSnapshot=(await masters.GetAsync(GestionServicios.Application.Maintenance.MaintenanceDefinitions.Find("clientes")!,actor,naturalId))!;
    Check(naturalSnapshot.Text("RazonSocial")==""&&naturalSnapshot.Text("TipoIdentificacion")=="","Natural guarda número y nombres sin exigir tipo de documento");
    Check(await MasterSave("clientes",new{TipoEntidad="NATURAL",NumeroIdentificacion="12345678",Nombres="Ana",Apellidos="Pérez",Activo=true})>0,"SQL admite natural sin tipo de documento");
    Check(await MasterSave("clientes",new{TipoEntidad="JURIDICA",NumeroIdentificacion="20123456789",RazonSocial="Empresa",Activo=true})>0,"SQL admite jurídica sin tipo de documento");
    await ExpectDomainFailure(()=>MasterSave("clientes",new{TipoEntidad="JURIDICA",TipoIdentificacion="RUC",RazonSocial="Sin número",Activo=true}),"NumeroIdentificacion","SQL exige número de documento");
    async Task SaveLinks(object[] links)
    {
        var currentClient=(await masters.GetAsync(clientDefinition,actor,autoClient))!;
        await MasterSave("clientes",new{TipoEntidad="JURIDICA",NumeroIdentificacion="20123456789",RazonSocial="Cliente actualizado",CodigoSAP="000SAP-B/02",Activo=true,Enlaces=links},autoClient,Convert.FromBase64String(currentClient.Text("VersionFila")));
    }
    await SaveLinks([new{Id=(long?)null,Descripcion="Portal documentos",Url="https://example.com/documentos"},new{Id=(long?)null,Descripcion="Facturas",Url="https://example.com/facturas"}]);
    var linkedClient=(await masters.GetAsync(clientDefinition,actor,autoClient))!;
    var linkedRows=linkedClient.Values["Enlaces"];
    Check(linkedRows.GetArrayLength()==2&&linkedRows[0].GetProperty("Descripcion").GetString()=="Portal documentos","Cliente guarda varios enlaces ordenados con descripción");
    var firstLink=linkedRows[0].GetProperty("Id").GetInt64();
    await SaveLinks([new{Id=(long?)firstLink,Descripcion="Portal actualizado",Url="https://example.com/nuevo"}]);
    var changedLinks=(await masters.GetAsync(clientDefinition,actor,autoClient))!;
    Check(changedLinks.Values["Enlaces"].GetArrayLength()==1&&changedLinks.Values["Enlaces"][0].GetProperty("Id").GetInt64()==firstLink&&changedLinks.Audit.Any(entry=>{if(entry.Changes is null)return false;using var changes=JsonDocument.Parse(entry.Changes);var previous=changes.RootElement.GetProperty("Anterior");var current=changes.RootElement.GetProperty("Nuevo");return previous.ValueKind==JsonValueKind.Object&&previous.GetProperty("Enlaces")[0].GetProperty("Url").GetString()=="https://example.com/documentos"&&current.GetProperty("Enlaces")[0].GetProperty("Url").GetString()=="https://example.com/nuevo";}),"Editar y retirar enlaces conserva ID y auditoría anterior/nueva");
    await ExpectDomainFailure(()=>SaveLinks([new{Id=(long?)null,Descripcion="Inválido",Url="javascript:alert(1)"}]),"HTTP","SQL rechaza esquema no permitido sin perder enlaces anteriores");
    var otherClient=(await masters.GetAsync(clientDefinition,actor,naturalId))!;
    await ExpectDomainFailure(()=>MasterSave("clientes",new{TipoEntidad="NATURAL",Nombres="Ana",Apellidos="Pérez",NumeroIdentificacion="12345678",Activo=true,Enlaces=new[]{new{Id=firstLink,Descripcion="Ajeno",Url="https://example.com/ajeno"}}},naturalId,Convert.FromBase64String(otherClient.Text("VersionFila"))),"no pertenece","SQL rechaza un enlace de otro cliente");
    await SaveLinks([]);
    Check((await masters.GetAsync(clientDefinition,actor,autoClient))!.Values["Enlaces"].GetArrayLength()==0,"Cliente puede retirar todos los enlaces con auditoría");
    await Execute(connection,"UPDATE dbo.Cliente SET UrlCargaDocumentos=N'https://example.com/anterior' WHERE ClienteId=1; EXEC sys.sp_dropextendedproperty @name=N'HYT_EnlaceAnteriorMigrado',@level0type=N'SCHEMA',@level0name=N'dbo',@level1type=N'TABLE',@level1name=N'ClienteEnlaceDocumento';");
    async Task MigrateLinks()
    {
        var migration=(await File.ReadAllTextAsync(Path.Combine(sources,"Script/022_EnlacesDocumentosCliente.sql"))).Replace("SistemaWebHYT",testDatabase,StringComparison.Ordinal);
        foreach(var batch in Regex.Split(migration,@"^[ \t]*GO[ \t]*(?:--[^\r\n]*)?\r?$",RegexOptions.Multiline|RegexOptions.IgnoreCase))if(!string.IsNullOrWhiteSpace(batch))await Execute(connection,batch);
    }
    await MigrateLinks();await MigrateLinks();
    Check(Convert.ToInt64(await Scalar(connection,"SELECT COUNT_BIG(1) FROM dbo.ClienteEnlaceDocumento WHERE ClienteId=1 AND Url=N'https://example.com/anterior'"))==1,"Migración conserva el enlace anterior sin duplicarlo al repetir");
    await Execute(connection,"DELETE FROM dbo.ClienteEnlaceDocumento WHERE ClienteId=1;");await MigrateLinks();
    Check(Convert.ToInt64(await Scalar(connection,"SELECT COUNT_BIG(1) FROM dbo.ClienteEnlaceDocumento WHERE ClienteId=1"))==0,"Repetir migración no restaura enlaces retirados por el usuario");
    var parallel=await Task.WhenAll(Enumerable.Range(0,8).Select(index=>MasterSave("clientes",new{TipoEntidad="JURIDICA",TipoIdentificacion="RUC",NumeroIdentificacion="20123456789",RazonSocial="Alta paralela "+index,Activo=true})));
    Check(Convert.ToInt64(await Scalar(connection,"SELECT COUNT(DISTINCT Codigo) FROM dbo.Cliente WHERE ClienteId IN("+string.Join(',',parallel)+")"))==8,"Altas concurrentes reciben códigos distintos sin código manual");
    var unit=await MasterSave("unidades",new{Codigo="PKG",Descripcion="Paquete",Activo=true});
    Check((await masters.GetAsync(GestionServicios.Application.Maintenance.MaintenanceDefinitions.Find("unidades")!,actor,unit))!.Text("Codigo")=="PKG","Abreviatura de unidad es libre y su identificador interno es automático");
    await ExpectDomainFailure(()=>MasterSave("unidades",new{Codigo="PKG",Descripcion="Otra",Activo=true}),"abreviatura","Unidad duplicada devuelve validación de abreviatura");
    var specialization=await MasterSave("especialidades",new{Codigo="MANUAL",Descripcion="Especialidad nueva",Activo=true});
    var serviceType=await MasterSave("tipos-servicio",new{Codigo="MANUAL",Descripcion="Tipo nuevo",Activo=true});
    var zone=await MasterSave("zonas",new{Codigo="MANUAL",Descripcion="Zona nueva",Activo=true});
    Check((await masters.GetAsync(GestionServicios.Application.Maintenance.MaintenanceDefinitions.Find("especialidades")!,actor,specialization))!.Text("Codigo").StartsWith("ESP-")&&(await masters.GetAsync(GestionServicios.Application.Maintenance.MaintenanceDefinitions.Find("tipos-servicio")!,actor,serviceType))!.Text("Codigo").StartsWith("TS-")&&(await masters.GetAsync(GestionServicios.Application.Maintenance.MaintenanceDefinitions.Find("zonas")!,actor,zone))!.Text("Codigo").StartsWith("ZON-"),"Especialidades, tipos y zonas usan numeración automática");
    var material=await MasterSave("materiales",new{CodigoInterno="MANUAL",Posicion="NO_DEBE_GUARDAR",CodigoMaterialSAP="000SAP-A",Descripcion="Material con precios",UnidadMedidaId=unit,EspecialidadId=specialization,Activo=true,ZonaIds=new[]{zone}});
    var materialRecord=(await masters.GetAsync(GestionServicios.Application.Maintenance.MaintenanceDefinitions.Find("materiales")!,actor,material))!;
    Check(materialRecord.Text("CodigoInterno").StartsWith("MAT-")&&materialRecord.Text("CodigoMaterialSAP")=="000SAP-A"&&await Scalar(connection,$"SELECT Posicion FROM dbo.ServicioMaterial WHERE MaterialId={material}") is DBNull,"Material conserva SAP libre, genera código interno y no recibe posición del maestro");
    var tariff=await MasterSave("tarifarios",new{Codigo="MANUAL",Nombre="Tarifario del material",Activo=true});
    var tariffVersion=await MasterSave("versiones-tarifario",new{TarifarioId=tariff,CodigoMonedaPredeterminada="PEN",CodigoTipoIGVPredeterminado="SIN_IGV",ZonaIds=new[]{zone}});
    var price=await MasterSave("partidas-tarifario",new{VersionId=tariffVersion,MaterialId=material,Posicion="01-A",CodigoMaterialSAP="000SAP-A",TipoServicioId=serviceType,CodigoMoneda="PEN",CodigoTipoIGV="SIN_IGV",PrecioPEN=12.34m,PrecioUSD=(decimal?)null,Activo=true,ZonaIds=new[]{zone}});
    var prices=await masters.MaterialTariffsAsync(actor,material,1);
    Check(prices.Total==1&&prices.Rows.Single().GetProperty("Id").GetInt64()==price&&prices.Rows[0].GetProperty("PrecioPEN").GetDecimal()==12.34m&&prices.Rows[0].GetProperty("PrecioUSD").ValueKind==JsonValueKind.Null,"Precios de ficha recuperan moneda, zonas y USD vacío desde SQL");
    for(var priceIndex=0;priceIndex<22;priceIndex++)await MasterSave("partidas-tarifario",new{VersionId=tariffVersion,MaterialId=material,Posicion="P-"+priceIndex,CodigoMoneda="PEN",CodigoTipoIGV="SIN_IGV",PrecioPEN=10m,Activo=true,ZonaIds=Array.Empty<int>()});
    prices=await masters.MaterialTariffsAsync(actor,material,1);var nextPrices=await masters.MaterialTariffsAsync(actor,material,2);
    Check(prices.Rows.Count==20&&nextPrices.Rows.Count==3&&prices.Total==23&&!prices.Rows.Select(row=>row.GetProperty("Id").GetInt64()).Intersect(nextPrices.Rows.Select(row=>row.GetProperty("Id").GetInt64())).Any(),"Tarifas del material se paginan en SQL de veinte sin duplicados");
    await ExpectDomainFailure(()=>masters.MaterialTariffsAsync("sql-test-budget-only",material,1),"permiso","Precios del material exigen consulta de materiales y tarifarios");
    var versionRecord=(await masters.GetAsync(GestionServicios.Application.Maintenance.MaintenanceDefinitions.Find("versiones-tarifario")!,actor,tariffVersion))!;
    await masters.PublishVersionAsync(actor,tariffVersion,Convert.FromBase64String(versionRecord.Text("VersionFila")),CancellationToken.None);
    var priceRecord=(await masters.GetAsync(GestionServicios.Application.Maintenance.MaintenanceDefinitions.Find("partidas-tarifario")!,actor,price))!;
    await ExpectDomainFailure(()=>MasterSave("partidas-tarifario",new{VersionId=tariffVersion,MaterialId=material,CodigoMoneda="PEN",CodigoTipoIGV="SIN_IGV",PrecioPEN=99m,Activo=true,ZonaIds=Array.Empty<int>()},price,Convert.FromBase64String(priceRecord.Text("VersionFila"))),"borrador","Precio publicado no se modifica desde ninguna ruta");
    var publishedVersion=(await masters.GetAsync(GestionServicios.Application.Maintenance.MaintenanceDefinitions.Find("versiones-tarifario")!,actor,tariffVersion))!;
    var copiedVersion=await masters.CloneVersionAsync(actor,tariffVersion,Convert.FromBase64String(publishedVersion.Text("VersionFila")),CancellationToken.None);
    Check(copiedVersion!=tariffVersion&&(await masters.MaterialTariffsAsync(actor,material,1)).Rows.Any(row=>row.GetProperty("VersionId").GetInt64()==copiedVersion&&row.GetProperty("Estado").GetString()=="BORRADOR"),"Copiar versión deja precios borrador disponibles sin alterar el histórico");
    var ubigeoCatalogs=await masters.CatalogsAsync(actor);
    Check(ubigeoCatalogs["Departments"].Count==25&&ubigeoCatalogs["Provinces"].Count==196&&ubigeoCatalogs["Districts"].Count==1889,"Ubigeos copiados íntegros de FralseCont en HYT");
    await Execute(connection,"UPDATE dbo.Cliente SET DireccionFiscal=N'Av. Cliente 123',CodigoUbigeo='150101' WHERE ClienteId=1; UPDATE dbo.Tienda SET Direccion=N'Av. Tienda 456',CodigoUbigeo='150132' WHERE TiendaId=1; UPDATE dbo.ConfiguracionHYT SET ConsideracionesPresupuesto=N'Condiciones de prueba configuradas' WHERE Id=1;");
    var addressDefaults=await sapBudgets.DefaultsAsync(actor,new DateOnly(2026,11,5),1,1,CancellationToken.None);
    Check(addressDefaults.GetProperty("Cliente").GetProperty("DireccionFiscal").GetString()!.Contains("Av. Cliente 123 / LIMA / LIMA / LIMA")&&addressDefaults.GetProperty("Tienda").GetProperty("Direccion").GetString()!.Contains("SAN JUAN DE LURIGANCHO"),"Presupuesto recibe direcciones completas del cliente y tienda con ubigeo");
    Check(addressDefaults.GetProperty("Empresa").GetProperty("ConsideracionesPresupuesto").GetString()=="Condiciones de prueba configuradas","Datos iniciales incluyen las consideraciones de configuración HYT");
    var addressList=await masters.ListAsync(clientDefinition,actor,"CLI1",1,cancellationToken:CancellationToken.None);
    Check(addressList.Rows.Any(row=>row.Extra!.Contains("Av. Cliente 123 / LIMA / LIMA / LIMA")),"Listado de clientes muestra dirección con departamento provincia distrito");
    var addressedBudget=await Save(connection,Payload([]));
    Check(Convert.ToString(await Scalar(connection,$"SELECT ClienteDireccion+N' | '+TiendaDireccion FROM dbo.PresupuestoRevision WHERE RevisionId={addressedBudget.RevisionId}"))!.Contains("SAN JUAN DE LURIGANCHO"),"Presupuesto conserva dirección y ubigeo en su instantánea histórica");
    await ExpectDomainFailure(()=>MasterSave("tiendas",new{ClienteId=1,Nombre="Ubigeo inválido",CodigoUbigeo="999999",Activo=true}),"ubigeo","SQL rechaza distrito inexistente con mensaje de validación");
    var notesCompany=await companyRepository.GetAsync(actor,1,CancellationToken.None);
    var longNotes=new string('x',5000)+"\nCondiciones de presupuesto";
    await companyRepository.SaveAsync(actor,JsonSerializer.Serialize(new{Ruc=notesCompany!.Text("Ruc"),RazonSocial=notesCompany.Text("RazonSocial"),ConsideracionesPresupuesto=longNotes}),Convert.FromBase64String(notesCompany.Text("VersionFila")),CancellationToken.None);
    Check((await companyRepository.GetAsync(actor,1,CancellationToken.None))!.Text("ConsideracionesPresupuesto")==longNotes,"Configuración guarda consideraciones de más de 4000 caracteres sin truncarlas");
    var notesDefaults=await sapBudgets.DefaultsAsync(actor,new DateOnly(2026,11,5),1,1,CancellationToken.None);
    Check(notesDefaults.GetProperty("Empresa").GetProperty("ConsideracionesPresupuesto").GetString()==longNotes,"Nuevo presupuesto recibe las consideraciones completas de configuración");
    var beforeCompanyUbigeo=await companyRepository.GetAsync(actor,1,CancellationToken.None);
    await companyRepository.SaveAsync(actor,JsonSerializer.Serialize(new{Ruc=beforeCompanyUbigeo!.Text("Ruc"),RazonSocial=beforeCompanyUbigeo.Text("RazonSocial"),Direccion="Dirección HYT",CodigoUbigeo="150101",ConsideracionesPresupuesto=longNotes}),Convert.FromBase64String(beforeCompanyUbigeo.Text("VersionFila")),CancellationToken.None);
    var storedCompanyUbigeo=await companyRepository.GetAsync(actor,1,CancellationToken.None);
    Check(storedCompanyUbigeo!.Text("CodigoUbigeo")=="150101"&&storedCompanyUbigeo.Audit.Any(item=>item.Changes?.Contains("150101")==true),"Configuración HYT conserva y audita el ubigeo de su dirección");
    await ExpectDomainFailure(()=>companyRepository.SaveAsync(actor,JsonSerializer.Serialize(new{Ruc=storedCompanyUbigeo.Text("Ruc"),RazonSocial=storedCompanyUbigeo.Text("RazonSocial"),CodigoUbigeo="999999"}),Convert.FromBase64String(storedCompanyUbigeo.Text("VersionFila")),CancellationToken.None),"ubigeo","Configuración HYT rechaza un distrito inexistente");
    var companyMigration=(await File.ReadAllTextAsync(Path.Combine(sources,"Script/028_UbigeoConfiguracionHYT.sql"))).Replace("SistemaWebHYT",testDatabase,StringComparison.Ordinal);
    foreach(var batch in Regex.Split(companyMigration,@"^[ \t]*GO[ \t]*(?:--[^\r\n]*)?\r?$",RegexOptions.Multiline|RegexOptions.IgnoreCase))if(!string.IsNullOrWhiteSpace(batch))await Execute(connection,batch);
    Check((await companyRepository.GetAsync(actor,1,CancellationToken.None))!.Text("CodigoUbigeo")=="150101","Migración de ubigeo HYT repetible conserva la selección");
    var directPrice=await masters.SaveMaterialTariffAsync(actor,material,null,null,"CON_IGV",123.45m,null,true);
    var directRecord=(await masters.GetAsync(GestionServicios.Application.Maintenance.MaintenanceDefinitions.Find("partidas-tarifario")!,actor,directPrice))!;
    Check(directRecord.Text("EsTarifarioMaterial").Equals("true",StringComparison.OrdinalIgnoreCase)&&directRecord.Text("VersionEstado")=="VIGENTE"&&directRecord.Text("PrecioUSD")=="","Tarifario directo queda disponible sin publicación manual y permite USD vacío");
    var filteredChoice=await sapBudgets.TariffsAsync(actor,1,new DateOnly(2026,11,5),"PEN",null,1,CancellationToken.None,material,directPrice);
    Check(filteredChoice.Rows.Count==1&&filteredChoice.Rows[0].TariffLineId==directPrice,"Ayuda de tarifa filtra el material y registro elegido en SQL");
    var genericChoice=await sapBudgets.TariffsAsync(actor,0,new DateOnly(2026,11,5),"PEN",null,1,CancellationToken.None,material,directPrice);
    Check(genericChoice.Rows.Count==1,"Tarifa genérica del material disponible antes de seleccionar cliente");
    var privateChoice=await sapBudgets.TariffsAsync(actor,0,new DateOnly(2026,11,5),"PEN",null,1,CancellationToken.None,1,2);
    Check(privateChoice.Total==0,"Consulta sin cliente no ofrece tarifas privadas de otro cliente");
    var wrongChoice=await sapBudgets.TariffsAsync(actor,1,new DateOnly(2026,11,5),"PEN",null,1,CancellationToken.None,material+999,directPrice);
    Check(wrongChoice.Total==0,"Ayuda no devuelve tarifas de otro material");
    object NetDetail(long selectedMaterial,string tax="SIN_IGV")=>new{ClienteId=1,TiendaId=1,TipoAtencion="INVITACION",Descripcion="Detalle neto editable",DetalleNeto=true,Revision=new{FechaEmision="2026-11-05",CodigoMoneda="PEN",CodigoTipoIGV=tax},Partidas=new[]{new{MaterialId=selectedMaterial,PartidaTarifarioId=directPrice,OrigenPrecio="MANUAL",Descripcion="Material elegido",Unidad="UND",Cantidad=2.50m,PrecioManual=80m,CodigoMonedaOrigen="PEN",CodigoTipoIGVOrigen="SIN_IGV",MotivoPrecioManual="Precio editado en detalle",NumeroNecesidad="NEC-A1",TipoServicio=new string('N',100)}}};
    var netBudget=await Save(connection,NetDetail(material));
    Check(Convert.ToDecimal(await Scalar(connection,$"SELECT BaseImponible FROM dbo.PresupuestoPartida WHERE RevisionId={netBudget.RevisionId}"))==200m&&Convert.ToDecimal(await Scalar(connection,$"SELECT Total FROM dbo.PresupuestoRevision WHERE RevisionId={netBudget.RevisionId}"))==236m,"Detalle guarda precio neto editado, subtotal cantidad por precio e IGV por separado");
    Check(Convert.ToString(await Scalar(connection,$"SELECT TipoServicio FROM dbo.PresupuestoPartida WHERE RevisionId={netBudget.RevisionId}"))==new string('N',100),"Detalle conserva el nombre completo del servicio sin truncar a sesenta caracteres");
    await ExpectFailure(()=>Save(connection,NetDetail(material,"CON_IGV")),"sin IGV","SQL exige tratamiento neto en el nuevo detalle");
    await ExpectFailure(()=>Save(connection,NetDetail(1)),"material","SQL rechaza una tarifa perteneciente a otro material");
    async Task<bool> DirectSelectable()
    {
        for(var choicePage=1;;choicePage++)
        {
            var choices=await sapBudgets.TariffsAsync(actor,1,new DateOnly(2026,11,5),"PEN",null,choicePage,CancellationToken.None);
            if(choices.Rows.Any(row=>row.TariffLineId==directPrice&&row.Price==123.45m))return true;
            if((long)choicePage*20>=choices.Total)return false;
        }
    }
    Check(await DirectSelectable(),"Presupuesto encuentra el precio agregado desde el material");
    var allPrices=await masters.MaterialTariffsAsync(actor,0,1,search:"%no-existe-hyt-9823%");
    Check(allPrices.Total==0&&allPrices.Rows.Count==0,"Tarifarios unificado respeta filtro sin coincidencias");
    var unified=await masters.MaterialTariffsAsync(actor,0,1);
    Check(unified.Total>0&&unified.Rows.Any(row=>row.GetProperty("Id").GetInt64()==directPrice)&&unified.Rows.All(row=>row.GetProperty("MaterialId").GetInt64()>0),"Tarifarios unificado incluye precio creado desde material y su identidad");
    var usdOnly=await masters.SaveMaterialTariffAsync(actor,material,null,null,"SIN_IGV",null,9.50m,true);
    var usdRecord=await masters.GetAsync(GestionServicios.Application.Maintenance.MaintenanceDefinitions.Find("partidas-tarifario")!,actor,usdOnly);
    Check(usdRecord!.Text("PrecioPEN")==""&&usdRecord.Text("PrecioUSD")!="","Tarifario permite solo dólares sin inventar precio en soles");
    await ExpectDomainFailure(()=>masters.SaveMaterialTariffAsync(actor,material,null,null,"SIN_IGV",null,null,true),"precio","Tarifario rechaza ambos importes vacíos");
    var directOriginal=Convert.FromBase64String(directRecord.Text("VersionFila"));
    await masters.SaveMaterialTariffAsync(actor,material,directPrice,directOriginal,"SIN_IGV",75.50m,20m,false);
    var directChanged=(await masters.GetAsync(GestionServicios.Application.Maintenance.MaintenanceDefinitions.Find("partidas-tarifario")!,actor,directPrice))!;
    Check(directChanged.Text("FechaActualizacion")!=""&&directChanged.Text("UsuarioActualizacion")!=""&&directChanged.Audit.Count==2&&directChanged.Text("Activo").Equals("false",StringComparison.OrdinalIgnoreCase),"Tarifario conserva registro propio y audita actualización e inactivación");
    Check(!await DirectSelectable(),"Presupuesto excluye el tarifario inactivo");
    await ExpectDomainFailure(()=>masters.SaveMaterialTariffAsync(actor,material,directPrice,directOriginal,"SIN_IGV",10m,null,true),"cambió","Tarifario directo protege cambios concurrentes");
    await ExpectDomainFailure(()=>masters.SaveMaterialTariffAsync(actor,material+999,directPrice,Convert.FromBase64String(directChanged.Text("VersionFila")),"SIN_IGV",10m,null,true),"material","Tarifario directo valida el material propietario");
    await ExpectDomainFailure(()=>masters.SaveMaterialTariffAsync(actor,material,null,null,"SIN_IGV",1.234m,null,true),"decimales","SQL mantiene precisión de dos decimales en tarifario simple");
    await ExpectDomainFailure(()=>masters.SaveMaterialTariffAsync(actor,material,price,Convert.FromBase64String(priceRecord.Text("VersionFila")),"SIN_IGV",10m,null,true),"no permite","Tarifario simple no modifica publicaciones convencionales protegidas");
    var numberCheck=Convert.ToString(await Scalar(connection,"SELECT COUNT(1) FROM dbo.Presupuesto AS p WHERE LEFT(p.NumeroPresupuesto,6)<>CONVERT(CHAR(6),CONVERT(DATETIME2,p.FechaRegistro AT TIME ZONE 'UTC' AT TIME ZONE 'SA Pacific Standard Time'),112)"));
    Check(numberCheck=="0","Número visible usa año/mes de la fecha de registro de Lima");
    var parallelNumbers=await Task.WhenAll(Enumerable.Range(0,6).Select(async _=>{await using var concurrent=new SqlConnection(localConnection+"Initial Catalog="+testDatabase+";");await concurrent.OpenAsync();return (await Save(concurrent,Payload([]))).Number;}));
    Check(parallelNumbers.Distinct().Count()==6,"Altas concurrentes reciben correlativos mensuales distintos");
    var priorNumbers=Convert.ToString(await Scalar(connection,"SELECT p.PresupuestoId,p.NumeroPresupuesto FROM dbo.Presupuesto AS p ORDER BY p.PresupuestoId FOR JSON PATH"));
    var migrationText=(await File.ReadAllTextAsync(Path.Combine(sources,"Script/026_NumeroPresupuestoMensual.sql"))).Replace("SistemaWebHYT",testDatabase,StringComparison.Ordinal);
    foreach(var batch in Regex.Split(migrationText,@"^[ \t]*GO[ \t]*(?:--[^\r\n]*)?\r?$",RegexOptions.Multiline|RegexOptions.IgnoreCase))if(!string.IsNullOrWhiteSpace(batch))await Execute(connection,batch);
    Check(priorNumbers==Convert.ToString(await Scalar(connection,"SELECT p.PresupuestoId,p.NumeroPresupuesto FROM dbo.Presupuesto AS p ORDER BY p.PresupuestoId FOR JSON PATH")),"Reaplicar migración conserva números asignados e identificadores");

    await Execute(connection,"CREATE USER hyt_test_executor WITHOUT LOGIN; ALTER ROLE hyt_app_executor ADD MEMBER hyt_test_executor;");
    await Execute(connection,$"EXECUTE AS USER=N'hyt_test_executor'; EXEC dbo.spu_MantZona_Guardar @UsuarioId=N'{actor}',@CorrelacionId='66666666-6666-6666-6666-666666666666',@Datos=N'{{\"Descripcion\":\"Prueba permisos secuencia\",\"Activo\":true}}'; REVERT;");
    Check(true,"Rol SQL limitado ejecuta numeración automática por cadena de propiedad");
    var observed=await Save(connection,Payload([Manual()]));
    observed=await State(connection,observed,"ENVIADO");
    await ExpectFailure(()=>State(connection,observed,"VENCIDO"),"transición","SQL no permite asignar Vencido");
    foreach(var targetState in new[]{"OBSERVADO","RECHAZADO","CANCELADO"})await ExpectFailure(()=>State(connection,observed,targetState,null),"motivo","SQL exige motivo para "+targetState);
    observed=await State(connection,observed,"OBSERVADO");
    var observedStale=observed;
    var sameObserved=await Save(connection,Payload([Manual(175)],description:"Corrección directa"),observed);
    Check(sameObserved.RevisionId==observed.RevisionId&&sameObserved.RevisionNumber==observed.RevisionNumber,"Observado se edita sin generar nueva revisión");
    Check(Convert.ToString(await Scalar(connection,$"SELECT Estado FROM dbo.PresupuestoRevision WHERE RevisionId={observed.RevisionId}"))=="OBSERVADO"&&Convert.ToDecimal(await Scalar(connection,$"SELECT Total FROM dbo.PresupuestoRevision WHERE RevisionId={observed.RevisionId}"))==413m,"Edición conserva Observado y recalcula importes");
    Check(Convert.ToInt64(await Scalar(connection,$"SELECT COUNT_BIG(1) FROM dbo.EventoAuditoria WHERE Entidad=N'PresupuestoRevision' AND RegistroId=N'{observed.RevisionId}' AND EstadoNuevo=N'OBSERVADO'"))>=2,"Auditoría conserva estado real de observación y edición");
    await ExpectFailure(()=>Save(connection,Payload([Manual()]),observedStale),"cambió","Edición de observado respeta concurrencia");
    sameObserved=await State(connection,sameObserved,"ENVIADO");
    sameObserved=await State(connection,sameObserved,"APROBADO");
    Check(Convert.ToString(await Scalar(connection,$"SELECT Estado FROM dbo.PresupuestoRevision WHERE RevisionId={observed.RevisionId}"))=="APROBADO","Observado corregido puede reenviarse y aceptarse");
    var firmaCompany=(await companyRepository.GetAsync(actor,1,CancellationToken.None))!;
    var firmaVersion=Convert.FromBase64String(firmaCompany.Text("VersionFila"));
    var firmaPayload=JsonSerializer.Serialize(new{NombreOriginal="firma.png",ClaveAlmacenamiento="ConfigHYT/Firma/"+new string('a',32)+".png",TipoContenido="image/png",TamanoBytes=100,SHA256=new string('A',64)});
    await ExpectDomainFailure(()=>companyRepository.SaveSignatureAsync("usuario-inexistente",firmaPayload,firmaVersion,CancellationToken.None),"permiso","Firma SQL exige permiso de edición de configuración");
    await companyRepository.SaveSignatureAsync(actor,firmaPayload,firmaVersion,CancellationToken.None);
    var firmaStored=(await companyRepository.GetAsync(actor,1,CancellationToken.None))!;
    Check(firmaStored.Text("FirmaPresupuestoJson")==firmaPayload&&firmaStored.Audit.Any(row=>row.Action=="ACTUALIZAR_FIRMA"),"Firma conserva metadatos R2 y auditoría en SQL");
    await ExpectDomainFailure(()=>companyRepository.SaveSignatureAsync(actor,firmaPayload,firmaVersion,CancellationToken.None),"configuración","Firma protege la versión original de configuración");
    await ExpectDomainFailure(()=>companyRepository.SaveSignatureAsync(actor,firmaPayload.Replace("image/png","image/svg+xml"),Convert.FromBase64String(firmaStored.Text("VersionFila")),CancellationToken.None),"válida","Firma SQL rechaza formato diferente a PNG/JPG");
    await companyRepository.SaveAsync(actor,JsonSerializer.Serialize(new{Ruc=firmaStored.Text("Ruc"),RazonSocial=firmaStored.Text("RazonSocial")}),Convert.FromBase64String(firmaStored.Text("VersionFila")),CancellationToken.None);
    Check((await companyRepository.GetAsync(actor,1,CancellationToken.None))!.Text("FirmaPresupuestoJson")==firmaPayload,"Guardar datos corporativos no borra firma configurada");
    var subtitleCompany=(await companyRepository.GetAsync(actor,1,CancellationToken.None))!;
    Check(subtitleCompany.Text("SubtituloCotizacion").Contains("hytserviciosg@gmail.com")&&subtitleCompany.Text("SubtituloCotizacion").Contains("rene_torres35@hotmail.com"),"Subtítulo inicial conserva texto completo y ambos correos indicados");
    await companyRepository.SaveAsync(actor,JsonSerializer.Serialize(new{Ruc=subtitleCompany.Text("Ruc"),RazonSocial=subtitleCompany.Text("RazonSocial"),SubtituloCotizacion="Subtítulo modificado de prueba"}),Convert.FromBase64String(subtitleCompany.Text("VersionFila")),CancellationToken.None);
    var subtitleChanged=(await companyRepository.GetAsync(actor,1,CancellationToken.None))!;
    Check(subtitleChanged.Text("SubtituloCotizacion")=="Subtítulo modificado de prueba"&&subtitleChanged.Audit.Any(row=>row.Changes?.Contains("SubtituloCotizacion")==true),"Subtítulo configurable se guarda y audita sin alterar firma");
    await ExpectDomainFailure(()=>companyRepository.SaveAsync(actor,JsonSerializer.Serialize(new{Ruc=subtitleChanged.Text("Ruc"),RazonSocial=subtitleChanged.Text("RazonSocial"),SubtituloCotizacion=new string('X',2001)}),Convert.FromBase64String(subtitleChanged.Text("VersionFila")),CancellationToken.None),"longitud","Subtítulo valida límite de 2000 caracteres en SQL");
    var autoCodes=await Save(connection,new{ClienteId=1,TiendaId=1,TipoAtencion="INVITACION",Descripcion="Códigos desde maestros",Revision=new{FechaEmision="2026-11-05",CodigoMoneda="PEN",CodigoTipoIGV="SIN_IGV"},Partidas=new[]{new{MaterialId=1,OrigenPrecio="MANUAL",CodigoInterno="FALSO",Posicion="FALSA",CentroSAPFICO="FALSO",Descripcion="Servicio",Unidad="UND",Cantidad=1,PrecioManual=10,CodigoMonedaOrigen="PEN",CodigoTipoIGVOrigen="SIN_IGV",MotivoPrecioManual="Prueba"}}});
    Check(Convert.ToString(await Scalar(connection,$"SELECT p.Posicion FROM dbo.PresupuestoPartida AS p WHERE p.RevisionId={autoCodes.RevisionId}"))==Convert.ToString(await Scalar(connection,"SELECT m.CodigoInterno FROM dbo.ServicioMaterial AS m WHERE m.MaterialId=1")),"Posición se toma del código interno de material e ignora dato manipulado");
    Check(Convert.ToString(await Scalar(connection,$"SELECT p.CentroSAPFICO FROM dbo.PresupuestoPartida AS p WHERE p.RevisionId={autoCodes.RevisionId}"))==Convert.ToString(await Scalar(connection,"SELECT t.CentroSAPFICO FROM dbo.Tienda AS t WHERE t.TiendaId=1")),"Centro SAP FICO se obtiene de la tienda e ignora dato manipulado");
    await Execute(connection,$"UPDATE dbo.PresupuestoPartida SET CentroSAPFICO=NULL WHERE RevisionId={autoCodes.RevisionId};");
    var ficoFallback=(await sapBudgets.GetAsync(actor,autoCodes.BudgetId,autoCodes.RevisionId,1,CancellationToken.None))!;
    var masterFico=Convert.ToString(await Scalar(connection,"SELECT t.CentroSAPFICO FROM dbo.Tienda AS t WHERE t.TiendaId=1"));
    Check(ficoFallback.Data.GetProperty("Partidas")[0].GetProperty("CentroSAPFICO").GetString()==masterFico,"Consulta completa FICO vacío de presupuesto anterior desde su tienda");
    Check(await Scalar(connection,$"SELECT p.CentroSAPFICO FROM dbo.PresupuestoPartida AS p WHERE p.RevisionId={autoCodes.RevisionId}") is DBNull,"Respaldo FICO de consulta no modifica partidas históricas");
    await Execute(connection,$"UPDATE dbo.PresupuestoPartida SET CentroSAPFICO=N'FICO-HISTORICO' WHERE RevisionId={autoCodes.RevisionId};");
    var ficoHistorical=(await sapBudgets.GetAsync(actor,autoCodes.BudgetId,autoCodes.RevisionId,1,CancellationToken.None))!;
    Check(ficoHistorical.Data.GetProperty("Partidas")[0].GetProperty("CentroSAPFICO").GetString()=="FICO-HISTORICO","Consulta conserva FICO histórico cuando ya tiene valor");
    var contactStore=await MasterSave("tiendas",new{ClienteId=1,Nombre="Tienda para contacto presupuesto",Activo=true});
    var contactBudget=await Save(connection,Payload([Manual()],store:contactStore));
    var noStoreContacts=(await sapBudgets.GetAsync(actor,contactBudget.BudgetId,contactBudget.RevisionId,1,CancellationToken.None))!;
    Check(noStoreContacts.Text("ContactosTienda")=="","Sin contactos de tienda no toma un responsable interno ni inventa nombre");
    await MasterSave("contactos",new{ClienteId=1,TiendaId=contactStore,Nombre="Encargado Tienda",Activo=true});
    await MasterSave("contactos",new{ClienteId=1,TiendaId=contactStore,Nombre="Contacto inactivo",Activo=false});
    await MasterSave("contactos",new{ClienteId=1,Nombre="Contacto general del cliente",Activo=true});
    await MasterSave("contactos",new{ClienteId=1,TiendaId=1,Nombre="Contacto de otra tienda",Activo=true});
    var oneStoreContact=(await sapBudgets.GetAsync(actor,contactBudget.BudgetId,contactBudget.RevisionId,1,CancellationToken.None))!;
    Check(oneStoreContact.Text("ContactosTienda")=="Encargado Tienda","Presupuesto recupera solo contacto activo de su tienda y cliente");
    await MasterSave("contactos",new{ClienteId=1,TiendaId=contactStore,Nombre="Coordinador Tienda",Activo=true});
    var manyStoreContacts=(await sapBudgets.GetAsync(actor,contactBudget.BudgetId,contactBudget.RevisionId,1,CancellationToken.None))!;
    Check(manyStoreContacts.Text("ContactosTienda")=="Encargado Tienda / Coordinador Tienda","Varios contactos activos de tienda se presentan ordenados sin elegir uno arbitrariamente");
    Check(Convert.ToString(await Scalar(connection,"SELECT dbo.fn_ArchivoSegmento(N'OC 001/A')"))=="OC_001_A","SQL genera el mismo segmento seguro para carpeta por número");
    var folderBudget=await Save(connection,Payload([Manual()],attention:"INVITACION"));
    var folderOperation=Guid.NewGuid();
    var folderOcData=new {RevisionId=folderBudget.RevisionId,NumeroOCCliente="OC 001/A",FechaEmisionOC="2026-11-06",FechaRecepcionOC="2026-11-06",Adjunto=new{NombreOriginal="oc.pdf",ClaveAlmacenamiento="OCCliente/OC_001_A/"+folderOperation.ToString("N")+".pdf",TipoContenido="application/pdf",TamanoBytes=15,SHA256=new string('A',64)}};
    var folderOcId=await SaveDocument(connection,"OC",folderBudget.BudgetId,folderBudget.HeaderVersion,folderOcData,folderOperation);
    Check(folderOcId>0,"OC SQL acepta adjunto en carpeta derivada de su número externo");
    var folderOcVersion=(byte[])(await Scalar(connection,$"SELECT VersionFila FROM dbo.OrdenCompraCliente WHERE OrdenCompraId={folderOcId}"))!;
    var folderConfirmationOperation=Guid.NewGuid();
    var folderConfirmation=new{FechaConfirmacion="2026-11-06",NumeroConfirmacion="CONF/001",Adjunto=new{NombreOriginal="confirmacion.pdf",ClaveAlmacenamiento="ConfirmacionCliente/CONF_001/"+folderConfirmationOperation.ToString("N")+".pdf",TipoContenido="application/pdf",TamanoBytes=15,SHA256=new string('A',64)}};
    var folderConfirmationId=await SaveDocument(connection,"CONFIRMACION",folderOcId,folderOcVersion,folderConfirmation,folderConfirmationOperation);
    Check(folderConfirmationId>0,"Confirmación SQL acepta carpeta por su número y mantiene vínculo a OC");
    var ocWithGuide=await documents.ListAsync(ocKind,actor,DateTimeOffset.UtcNow.AddDays(-1),DateTimeOffset.UtcNow.AddDays(1),"OC-000A/01",1,null,null,CancellationToken.None);
    Check(ocWithGuide.Rows.Count==1&&ocWithGuide.Rows[0].GetProperty("GuiaId").GetInt64()==inviteGuide.Id&&ocWithGuide.Rows[0].GetProperty("NumeroGuia").GetString()=="INV-0008C"&&ocWithGuide.Rows[0].GetProperty("AdjuntosGuia").GetInt64()==0,"Listado SQL OC incluye guía del presupuesto y cuenta de adjuntos vigentes");
    var kpiPrefix="KPI-"+Guid.NewGuid().ToString("N");
    var kpiBudgets=new List<BudgetVersion>();
    for(var index=0;index<25;index++)
    {
        var sample=await Save(connection,Payload([Manual()],description:kpiPrefix,attention:index%2==0?"EMERGENCIA":"INVITACION"));
        if(index<5)sample=await State(connection,sample,"ENVIADO");
        kpiBudgets.Add(sample);
    }
    var kpiStart=DateTimeOffset.UtcNow.AddDays(-1);var kpiEnd=DateTimeOffset.UtcNow.AddDays(1);
    var kpiFirst=await repository.ListAsync(actor,kpiStart,kpiEnd,kpiPrefix,1,null,CancellationToken.None);
    var kpiSecond=await repository.ListAsync(actor,kpiStart,kpiEnd,kpiPrefix,2,null,CancellationToken.None);
    Check(kpiFirst.Total==25&&kpiFirst.Rows.Count==20&&kpiSecond.Rows.Count==5&&kpiFirst.StateCounts!.Values.Sum()==25&&kpiSecond.StateCounts!["BORRADOR"]==20,"KPI SQL presupuesto cuenta todos los registros y conserva paginación de 20");
    var kpiFiltered=await repository.ListAsync(actor,kpiStart,kpiEnd,kpiPrefix,1,"BORRADOR",CancellationToken.None,type:"EMERGENCIA",pendingGuide:true,pendingOrder:true);
    Check(kpiFiltered.Total==10&&kpiFiltered.StateCounts!.Values.Sum()==13&&kpiFiltered.Rows.All(row=>row.Type=="EMERGENCIA"&&!row.GuideId.HasValue&&!row.PurchaseOrderId.HasValue),"SQL combina tipo y pendientes; KPI conserva distribución previa al selector Estado");
    for(var index=0;index<21;index++)await SaveDocument(connection,"OC",kpiBudgets[index].BudgetId,kpiBudgets[index].HeaderVersion,new{RevisionId=kpiBudgets[index].RevisionId,NumeroOCCliente=kpiPrefix+"-"+index,FechaEmisionOC="2026-11-06",FechaRecepcionOC="2026-11-06"},Guid.NewGuid());
    var kpiOc=await documents.ListAsync(ocKind,actor,kpiStart,kpiEnd,kpiPrefix,1,null,null,CancellationToken.None,pendingConfirmation:true,pendingInvoice:true);
    var kpiOcSecond=await documents.ListAsync(ocKind,actor,kpiStart,kpiEnd,kpiPrefix,2,null,null,CancellationToken.None,pendingConfirmation:true,pendingInvoice:true);
    Check(kpiOc.Total==21&&kpiOc.Rows.Count==20&&kpiOcSecond.Rows.Count==1&&kpiOc.StateCounts!.Values.Sum()==21,"KPI OC y filtros pendientes cuentan 21 documentos con paginación SQL");
    var kpiWithoutOc=await repository.ListAsync(actor,kpiStart,kpiEnd,kpiPrefix,1,null,CancellationToken.None,pendingOrder:true);
    Check(kpiWithoutOc.Total==4&&kpiWithoutOc.StateCounts!.Values.Sum()==4,"Pendiente OC excluye presupuestos que ya tienen orden registrada");
    var rollbackBudget=await Save(connection,Payload([Manual()]));
    var rollbackOc=await SaveDocument(connection,"OC",rollbackBudget.BudgetId,rollbackBudget.HeaderVersion,new{RevisionId=rollbackBudget.RevisionId,NumeroOCCliente="CONF-ROLLBACK",FechaEmisionOC="2026-11-06",FechaRecepcionOC="2026-11-06"},Guid.NewGuid());
    await ExpectFailure(()=>DocumentState("oc-cliente",rollbackOc,rollbackBudget.BudgetId,"TRABAJO_COMPLETADO"),"confirmación del cliente","OC sin confirmación ni ejecución informa el requisito correcto");
    await Execute(connection,"CREATE OR ALTER TRIGGER dbo.trg_TestConfirmationAuditFailure ON dbo.EventoAuditoria AFTER INSERT AS BEGIN IF EXISTS(SELECT 1 FROM inserted AS evento WHERE evento.Accion=N'CAMBIAR_ESTADO_AUTOMATICO' AND evento.Entidad=N'OrdenCompraCliente') THROW 51999,N'Fallo auditoría automática OC',1; END;");
    try
    {
        var rollbackVersion=(byte[])(await Scalar(connection,$"SELECT VersionFila FROM dbo.OrdenCompraCliente WHERE OrdenCompraId={rollbackOc}"))!;
        await ExpectFailure(()=>SaveDocument(connection,"CONFIRMACION",rollbackOc,rollbackVersion,new{FechaConfirmacion="2026-11-06",NumeroConfirmacion="CONF-ROLLBACK"},Guid.NewGuid()),"Fallo auditoría","Error de auditoría revierte confirmación y cambio automático de OC");
        Check(Convert.ToString(await Scalar(connection,$"SELECT Estado FROM dbo.OrdenCompraCliente WHERE OrdenCompraId={rollbackOc}"))=="REGISTRADA"&&Convert.ToInt64(await Scalar(connection,$"SELECT COUNT_BIG(1) FROM dbo.ConfirmacionFacturacion WHERE OrdenCompraId={rollbackOc}"))==0,"Fallo transaccional no deja confirmación parcial ni OC completada");
    }
    finally{await Execute(connection,"DROP TRIGGER dbo.trg_TestConfirmationAuditFailure;");}
    var invoiceWithoutExecutionBudget=await Save(connection,Payload([Manual()]));
    var invoiceWithoutExecutionOc=await SaveDocument(connection,"OC",invoiceWithoutExecutionBudget.BudgetId,invoiceWithoutExecutionBudget.HeaderVersion,new{RevisionId=invoiceWithoutExecutionBudget.RevisionId,NumeroOCCliente="OC-SIN-EJECUCION",FechaEmisionOC="2026-11-06",FechaRecepcionOC="2026-11-06"},Guid.NewGuid());
    var invoiceWithoutExecutionVersion=(byte[])(await Scalar(connection,$"SELECT VersionFila FROM dbo.OrdenCompraCliente WHERE OrdenCompraId={invoiceWithoutExecutionOc}"))!;
    await SaveDocument(connection,"CONFIRMACION",invoiceWithoutExecutionOc,invoiceWithoutExecutionVersion,new{FechaConfirmacion="2026-11-06",NumeroConfirmacion="CONF-SIN-EJECUCION"},Guid.NewGuid());
    invoiceWithoutExecutionVersion=(byte[])(await Scalar(connection,$"SELECT VersionFila FROM dbo.OrdenCompraCliente WHERE OrdenCompraId={invoiceWithoutExecutionOc}"))!;
    await ExpectFailure(()=>MarkInvoice(invoiceWithoutExecutionOc,invoiceWithoutExecutionVersion,Guid.NewGuid()),"guía registrada","OC completada y confirmada no factura sin guía");
    var invoiceGuideParent=(byte[])(await Scalar(connection,$"SELECT VersionFila FROM dbo.Presupuesto WHERE PresupuestoId={invoiceWithoutExecutionBudget.BudgetId}"))!;
    await SaveDocument(connection,"GUIA",invoiceWithoutExecutionBudget.BudgetId,invoiceGuideParent,new{FechaActa="2026-11-06",NumeroActa="GUIA-SIN-EJECUCION"},Guid.NewGuid());
    var noExecutionMark=await MarkInvoice(invoiceWithoutExecutionOc,invoiceWithoutExecutionVersion,Guid.NewGuid());
    Check(noExecutionMark>0&&Convert.ToInt64(await Scalar(connection,$"SELECT COUNT_BIG(1) FROM dbo.Ejecucion WHERE PresupuestoId={invoiceWithoutExecutionBudget.BudgetId}"))==0,"OC Trabajo completado con guía y confirmación permite Facturado sin ejecución separada");
    var integralReport=await reportingRepository.ReportAsync(actor,"PRESUPUESTO",kpiStart,kpiEnd,1,CancellationToken.None,filters:new(Search:invoiceWithoutExecutionBudget.Number,State:"APROBADO"));
    var integralRow=integralReport.Rows.Single();
    Check(integralRow.GetProperty("NumeroCliente").GetString()=="OC-SIN-EJECUCION"&&integralRow.GetProperty("Guia").GetString()=="GUIA-SIN-EJECUCION"&&integralRow.GetProperty("Confirmacion").GetString()=="CONF-SIN-EJECUCION"&&integralRow.GetProperty("NumeroComprobante").GetString()=="000012A"&&integralRow.GetProperty("TipoComprobante").GetString()=="01","Reporte de presupuesto integra OC, guía, confirmación y comprobante vigente con sus números");
    Check(new[]{"PresupuestoRegistro","OCRegistro","GuiaRegistro","ConfirmacionRegistro","FacturadoRegistro"}.All(key=>integralRow.GetProperty(key).ValueKind!=JsonValueKind.Null),"Reporte conserva fechas de registro de cada relación documental");
    var integralAll=await reportingRepository.ReportAsync(actor,"PRESUPUESTO",kpiStart,kpiEnd,1,CancellationToken.None,filters:new(Export:true));
    Check(integralAll.Rows.Count==integralAll.Total&&integralAll.Rows.Count>20&&integralAll.Amounts.Sum(row=>row.Count)==integralAll.Total,"Exportación SQL devuelve todos los resultados y totales completos por moneda");
    var reportPending=await reportingRepository.ReportAsync(actor,"PRESUPUESTO",kpiStart,kpiEnd,1,CancellationToken.None,filters:new(Type:"EMERGENCIA",PendingGuide:true,PendingOrder:true,Export:true));
    Check(reportPending.Rows.Count>0&&reportPending.Rows.All(row=>row.GetProperty("TipoAtencion").GetString()=="EMERGENCIA"&&row.GetProperty("Guia").ValueKind==JsonValueKind.Null&&row.GetProperty("NumeroCliente").ValueKind==JsonValueKind.Null),"Reporte aplica tipo y pendientes simultáneamente antes de totales y exportación");
    foreach(var stateCode in new[]{"REGISTRADA","TRABAJO_COMPLETADO","CERRADA_OPERATIVAMENTE","CANCELADA"}){
        var stateReport=await reportingRepository.ReportAsync(actor,"OC_CLIENTE",kpiStart,kpiEnd,1,CancellationToken.None,filters:new(State:stateCode,Export:true));
        Check(stateReport.Rows.All(row=>row.GetProperty("Estado").GetString()==stateCode),"Reporte OC filtra estado vigente: "+stateCode);
    }
    var integralDash=await reportingRepository.DashboardAsync(actor,kpiStart,kpiEnd,CancellationToken.None);
    var pendingOcReport=await reportingRepository.ReportAsync(actor,"PRESUPUESTO",kpiStart,kpiEnd,1,CancellationToken.None,filters:new(PendingOrder:true));
    var pendingGuideReport=await reportingRepository.ReportAsync(actor,"PRESUPUESTO",kpiStart,kpiEnd,1,CancellationToken.None,filters:new(PendingGuide:true));
    var pendingInvoiceReport=await reportingRepository.ReportAsync(actor,"OC_CLIENTE",kpiStart,kpiEnd,1,CancellationToken.None,filters:new(PendingInvoice:true));
    Check(integralDash.GetProperty("PresupuestosSinOC").GetInt64()==pendingOcReport.Total&&integralDash.GetProperty("PresupuestosSinGuia").GetInt64()==pendingGuideReport.Total&&integralDash.GetProperty("PendientesFacturar").GetInt64()==pendingInvoiceReport.Total,"Pendientes del dashboard concilian con los nuevos filtros de reportes");
    var restrictedIntegral=await reportingRepository.ReportAsync("sql-test-budget-only","PRESUPUESTO",kpiStart,kpiEnd,1,CancellationToken.None,filters:new(Search:invoiceWithoutExecutionBudget.Number));
    Check(restrictedIntegral.Rows.Single().GetProperty("Guia").ValueKind==JsonValueKind.Null&&restrictedIntegral.Rows.Single().GetProperty("NumeroComprobante").ValueKind==JsonValueKind.Null,"Reporte presupuesto oculta referencias de módulos sin permiso");
    await Execute(connection, "UPDATE dbo.AspNetUsers SET PasswordHash=N'hash-fixture-admin',SecurityStamp=N'stamp-fixture-admin',ConcurrencyStamp=N'version-fixture-admin' WHERE Id=N'sql-test-editor';");
    var securityConnections = new GestionServicios.Infrastructure.SqlConnectionFactory(connection.ConnectionString);
    var security = new SqlSecurityAdministrationRepository(securityConnections);
    var access = new SqlAccessRepository(securityConnections);
    const string role = "qa-seguridad-consulta";
    const string user = "qa-seguridad-usuario";
    const string tempHash = "hash-temporal-solo-prueba";
    const string ownHash = "hash-propio-solo-prueba";
    await security.SaveRoleAsync(actor, new(role, null, "QA Consulta parcial",
        ["CLIENTES.CONSULTAR", "CLIENTES.CREAR"]), default);
    var userCommand = new SecurityUserCommand(user, null, "Ana", "Pruebas", "qa-seguridad@example.test", role, true, tempHash);
    await security.SaveUserAsync(actor, userCommand, default);
    var createdUser = await access.FindByNameAsync("QA-SEGURIDAD@EXAMPLE.TEST");
    Check(createdUser is { MustChangePassword: true } && createdUser.PasswordHash == tempHash &&
        createdUser.Permissions.Contains("CLIENTES.CREAR") && !createdUser.Permissions.Contains("CLIENTES.MODIFICAR"),
        "Alta con contraseña temporal y permisos separados por acción");
    var listedUsers = await security.UsersAsync(actor, "Ana Pruebas", true, "TEMPORAL", 1, default);
    Check(listedUsers.Total == 1 && listedUsers.Rows.Single().GetProperty("Email").GetString() == userCommand.Email,
        "Listado de usuarios busca el nombre completo y filtra contraseña pendiente");
    await ExpectDomainFailure(() => security.SaveUserAsync(actor, userCommand with { Id = "qa-seguridad-duplicado", Email = "QA-SEGURIDAD@EXAMPLE.TEST" }, default),
        "correo", "Correo duplicado se rechaza sin distinguir mayúsculas");
    await ExpectDomainFailure(() => security.RolesAsync(user, null, 1, default),
        "permiso", "Usuario temporal no consulta perfiles por SQL");
    await ExpectDomainFailure(() => security.ChangePasswordAsync(user, "sello-anterior", tempHash, ownHash, true, default),
        "Inicia sesión nuevamente", "Cambio obligatorio exige el sello vigente");
    await security.ChangePasswordAsync(user, createdUser!.SecurityStamp, tempHash, ownHash, true, default);
    var updatedUser = (await access.FindByIdAsync(user))!;
    Check(!updatedUser.MustChangePassword && updatedUser.PasswordHash == ownHash &&
        updatedUser.SecurityStamp != createdUser.SecurityStamp, "Cambio obligatorio persiste el hash y revoca la sesión temporal");
    await ExpectFailure(() => Execute(connection, "EXEC dbo.spu_MantCliente_Guardar @UsuarioId=N'qa-seguridad-usuario',@CorrelacionId=NULL,@Datos=N'{}',@Id=1;"),
        "permiso", "Permiso para crear no autoriza editar el cliente por SQL");
    await ExpectDomainFailure(() => security.AuditAsync(user, DateTimeOffset.UtcNow.AddDays(-1), DateTimeOffset.UtcNow.AddDays(1),
        null, null, null, null, 1, default), "permiso", "Auditoría restringida por SQL");
    var partialRole = (await security.RoleAsync(actor, role, default))!.Value;
    await security.SaveRoleAsync(actor, new(role, partialRole.GetProperty("ConcurrencyStamp").GetString(),
        "QA Consulta parcial", ["CLIENTES.CONSULTAR"]), default);
    var reduced = (await access.FindByIdAsync(user))!;
    Check(reduced.SecurityStamp == updatedUser.SecurityStamp && !reduced.Permissions.Contains("CLIENTES.CREAR"),
        "Retirar una acción se refleja en una sesión vigente");
    var details = (await security.UserAsync(actor, user, default))!.Value;
    await security.UserActionAsync(actor, user, details.GetProperty("ConcurrencyStamp").GetString()!,
        "RESTABLECER", tempHash, "Restablecimiento de prueba", default);
    var reset = (await access.FindByIdAsync(user))!;
    Check(reset.MustChangePassword && reset.SecurityStamp != updatedUser.SecurityStamp,
        "Restablecimiento exige cambiar la contraseña y revoca sesiones");
    await ExpectDomainFailure(() => security.UserActionAsync(actor, actor, "version-fixture-admin",
        "DESACTIVAR", null, "Prueba de protección", default), "administrador", "No se puede desactivar al último administrador");
    var adminRole = (await security.RoleAsync(actor, "HYT-ADMIN", default))!.Value;
    await ExpectDomainFailure(() => security.SaveRoleAsync(actor, new("HYT-ADMIN", adminRole.GetProperty("ConcurrencyStamp").GetString(),
        "Administrador", ["CLIENTES.CONSULTAR"]), default), "administrador", "No se pueden retirar los permisos al último administrador");
    var audit = await security.AuditAsync(actor, DateTimeOffset.UtcNow.AddDays(-1), DateTimeOffset.UtcNow.AddDays(1),
        userCommand.Email, null, "AspNetUsers", "USUARIO_RESTABLECER", 1, default);
    var auditDetail = (await security.AuditEventAsync(actor, audit.Data.Rows.Single().GetProperty("EventoId").GetInt64(), default))!.Value;
    Check(audit.Data.Total == 1 && audit.Options.Any(row => row.GetProperty("Accion").GetString() == "USUARIO_RESTABLECER") &&
        auditDetail.GetProperty("Motivo").GetString() == "Restablecimiento de prueba" &&
        !auditDetail.GetRawText().Contains(tempHash) && !auditDetail.GetRawText().Contains(ownHash),
        "Auditoría filtra por usuario y acción, conserva motivo y excluye hashes");
    var budgetAudit = await security.AuditAsync(actor, DateTimeOffset.UtcNow.AddDays(-1), DateTimeOffset.UtcNow.AddDays(1),
        null, null, "Presupuesto", "GENERAR_IMPRESION", 1, default);
    Check(budgetAudit.Data.Rows.Count > 0 &&
        budgetAudit.Data.Rows.All(row => row.GetProperty("Entidad").GetString() == "PresupuestoRevision") &&
        budgetAudit.Options.Count(row => row.GetProperty("Entidad").GetString() == "Presupuesto" &&
            row.GetProperty("Accion").GetString() == "GENERAR_IMPRESION") == 1 &&
        budgetAudit.Options.All(row => row.GetProperty("Entidad").GetString() != "PresupuestoRevision"),
        "Auditoría agrupa Presupuestos y permite filtrar eventos de revisión");
    const string anotherAdmin = "qa-seguridad-admin";
    await security.SaveUserAsync(actor, userCommand with { Id = anotherAdmin, Email = "qa-admin@example.test", RoleId = "HYT-ADMIN" }, default);
    var another = (await access.FindByIdAsync(anotherAdmin))!;
    await security.ChangePasswordAsync(anotherAdmin, another.SecurityStamp, tempHash, ownHash, true, default);
    var anotherVersion = (await security.UserAsync(actor, anotherAdmin, default))!.Value.GetProperty("ConcurrencyStamp").GetString()!;
    async Task<bool> TryDeactivate(string id, string version)
    {
        try { await security.UserActionAsync(id, id, version, "DESACTIVAR", null, "Prueba concurrente", default); return true; }
        catch (GestionServicios.Application.Maintenance.MaintenanceException error) when (error.Message.Contains("administrador", StringComparison.Ordinal))
        { return false; }
    }
    var outcomes = await Task.WhenAll(TryDeactivate(actor, "version-fixture-admin"), TryDeactivate(anotherAdmin, anotherVersion));
    Check(outcomes.Count(success => success) == 1, "Dos desactivaciones concurrentes conservan un administrador activo");
    Console.WriteLine($"SQL real: {checks} comprobaciones aprobadas en base aislada.");
}
finally
{
    if (!Regex.IsMatch(testDatabase,@"^SistemaWebHYT_Test_[a-f0-9]{32}$")) throw new InvalidOperationException("Destino temporal inesperado.");
    await using var cleanup = new SqlConnection(localConnection + "Initial Catalog=master;"); await cleanup.OpenAsync();
    await Execute(cleanup,$"IF DB_ID(N'{testDatabase}') IS NOT NULL BEGIN ALTER DATABASE [{testDatabase}] SET SINGLE_USER WITH ROLLBACK IMMEDIATE; DROP DATABASE [{testDatabase}]; END;");
    Console.WriteLine("Base temporal eliminada; SistemaWebHYT de trabajo no recibió datos de prueba.");
}

sealed record BudgetVersion(long BudgetId,long RevisionId,string Number,int RevisionNumber,byte[] HeaderVersion,byte[] RevisionVersion);
