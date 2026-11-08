using System.Text.Json;
using GestionServicios.Application.Security;

sealed class TestSecurityAdministration(TestAccessRepository access) : ISecurityAdministrationRepository
{
    public Task ChangePasswordAsync(string actor, string expectedStamp, string expectedHash, string newHash, bool mandatory, CancellationToken ct)
    {
        var account = access.Account;
        if (account.Id != actor || account.SecurityStamp != expectedStamp || account.PasswordHash != expectedHash ||
            account.MustChangePassword != mandatory)
            throw new InvalidOperationException("La prueba recibió una versión de sesión incorrecta.");
        access.Account = account with { PasswordHash = newHash, SecurityStamp = Guid.NewGuid().ToString("N"), MustChangePassword = false };
        return Task.CompletedTask;
    }

    public Task<SecurityPage> UsersAsync(string actor, string? search, bool? active, string? situation, int page, CancellationToken ct) =>
        Task.FromResult(new SecurityPage([UserFixture], 1, page));

    public Task<JsonElement?> UserAsync(string actor, string id, CancellationToken ct) => Task.FromResult<JsonElement?>(UserFixture);

    public Task<(IReadOnlyList<JsonElement> Roles, IReadOnlyList<JsonElement> Permissions)> CatalogsAsync(string actor, CancellationToken ct) =>
        Task.FromResult<(IReadOnlyList<JsonElement>, IReadOnlyList<JsonElement>)>(([RoleFixture], [
            JsonSerializer.SerializeToElement(new { Codigo = "CLIENTES.CONSULTAR" }),
            JsonSerializer.SerializeToElement(new { Codigo = "CLIENTES.CREAR" }),
            JsonSerializer.SerializeToElement(new { Codigo = "CLIENTES.MODIFICAR" })
        ]));

    public Task<SecurityPage> RolesAsync(string actor, string? search, int page, CancellationToken ct) =>
        Task.FromResult(new SecurityPage([RoleFixture], 1, page));

    public Task<JsonElement?> RoleAsync(string actor, string id, CancellationToken ct) => Task.FromResult<JsonElement?>(RoleFixture);

    public (string? Module, string? Action) LastAuditFilter { get; private set; }

    public Task<SecurityAuditPage> AuditAsync(string actor, DateTimeOffset from, DateTimeOffset until, string? search,
        string? user, string? entity, string? action, int page, CancellationToken ct)
    {
        LastAuditFilter = (entity, action);
        return Task.FromResult(new SecurityAuditPage(new SecurityPage([AuditFixture], 1, page),
            [AuditFixture, JsonSerializer.SerializeToElement(new { Entidad = "Presupuesto", Accion = "GENERAR_IMPRESION" })]));
    }

    public Task<JsonElement?> AuditEventAsync(string actor, long id, CancellationToken ct) => Task.FromResult<JsonElement?>(AuditFixture);

    public Task SaveUserAsync(string actor, SecurityUserCommand input, CancellationToken ct) => throw new NotSupportedException();
    public Task UserActionAsync(string actor, string id, string version, string action, string? hash, string reason, CancellationToken ct) => throw new NotSupportedException();
    public Task SaveRoleAsync(string actor, SecurityRoleCommand input, CancellationToken ct) => throw new NotSupportedException();

    private static readonly JsonElement UserFixture = JsonSerializer.SerializeToElement(new
    {
        Id = "test-account", Nombres = "Operador", Apellidos = "HYT", Email = "prueba@example.test",
        Activo = true, Situacion = "DISPONIBLE", Perfil = "Consulta", PerfilId = "consulta", ConcurrencyStamp = "version"
    });

    private static readonly JsonElement RoleFixture = JsonSerializer.SerializeToElement(new
    {
        Id = "consulta", Name = "Consulta", Usuarios = 1, ConcurrencyStamp = "version",
        Permisos = new[] { new { Codigo = "CLIENTES.CONSULTAR" } }
    });

    private static readonly JsonElement AuditFixture = JsonSerializer.SerializeToElement(new
    {
        EventoId = 1, Usuario = "Operador HYT", Entidad = "AspNetUsers", Accion = "ACTUALIZAR_USUARIO",
        Referencia = "prueba@example.test", FechaRegistro = new DateTime(2026, 11, 7, 18, 30, 0, DateTimeKind.Utc),
        Cambios = new { Anterior = new { Nombres = "Ana", PasswordHash = "secreto-fixture" },
            Nuevo = new { Nombres = "María", PasswordHash = "otro-secreto-fixture" } }
    });
}
