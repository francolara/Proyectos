using System.Security.Claims;
using SistemaAdministrativoWeb.Infrastructure.Empresas;
using SistemaAdministrativoWeb.Infrastructure.Suscripciones;

namespace SistemaAdministrativoWeb.Infrastructure.Security;

public sealed class ModulePermissionService(
    ICurrentCompanyAccessor currentCompanyAccessor,
    ICuentaAdministradoraRepository cuentaAdministradoraRepository) : IModulePermissionService
{
    private static readonly HashSet<string> AccountModules = new(StringComparer.OrdinalIgnoreCase)
    {
        "DASHBOARD",
        "EMPRESAS",
        "USUARIOS",
        "CONFIGURACION",
        "MISUSCRIPCION",
        "AYUDA"
    };

    private static readonly HashSet<string> CompanyModules = new(StringComparer.OrdinalIgnoreCase)
    {
        "PLANCUENTA",
        "CENTROCOSTO",
        "CUENTACORRIENTE",
        "PERSONAS",
        "TIPOCAMBIO",
        "ORIGENES",
        "CUENTASDESTINO",
        "CONFIGCONTABLE",
        "ASIENTOS",
        "COMPRAS",
        "VENTAS",
        "CAJABANCOS",
        "TRANSFERENCIAS",
        "APLICACIONES",
        "PROCESOS",
        "REPORTES",
        "LIBROELECTRONICO"
    };

    private bool initialized;
    private bool isSuperAdmin;
    private bool hasAccess;
    private bool hasCompanyContext;
    private string? aspNetUserId;
    private int? idCuentaAdministradora;
    private ContextoLoginUsuarioDto? loginContext;
    private string? accountRole;
    private Dictionary<string, UsuarioCuentaPermisoDto> accountPermissions = new(StringComparer.OrdinalIgnoreCase);
    private Dictionary<string, UsuarioCuentaPermisoDto> companyPermissions = new(StringComparer.OrdinalIgnoreCase);

    public async Task<bool> CanAccessModuleAsync(
        ClaimsPrincipal principal,
        string moduleCode,
        CancellationToken cancellationToken = default)
    {
        return await CanAccessModuleAsync(principal, moduleCode, ModulePermissionOperation.View, cancellationToken);
    }

    public async Task<bool> CanAccessModuleAsync(
        ClaimsPrincipal principal,
        string moduleCode,
        ModulePermissionOperation operation,
        CancellationToken cancellationToken = default)
    {
        var result = await EvaluateAccessAsync(principal, moduleCode, operation, cancellationToken);
        return result.IsAllowed;
    }

    public async Task<ModuleAccessResult> EvaluateAccessAsync(
        ClaimsPrincipal principal,
        string moduleCode,
        ModulePermissionOperation operation,
        CancellationToken cancellationToken = default)
    {
        await EnsureInitializedAsync(principal, cancellationToken);

        if (isSuperAdmin)
        {
            return new ModuleAccessResult
            {
                IsAllowed = true,
                Scope = ResolveScope(moduleCode),
                HasCompanyContext = true
            };
        }

        if (!hasAccess || string.IsNullOrWhiteSpace(moduleCode))
        {
            return new ModuleAccessResult
            {
                IsAllowed = false,
                Scope = ResolveScope(moduleCode),
                HasCompanyContext = hasCompanyContext,
                Message = "El usuario no tiene acceso configurado a la cuenta administradora."
            };
        }

        if (accountPermissions.TryGetValue(moduleCode, out var accountPermission))
        {
            var explicitAccountAccess = EvaluateExplicitPermission(accountPermission, operation);
            if (explicitAccountAccess.HasValue)
            {
                return new ModuleAccessResult
                {
                    IsAllowed = explicitAccountAccess.Value,
                    Scope = ModuleScope.Account,
                    HasCompanyContext = true,
                    Message = "No tiene permisos sobre esta opcion general."
                };
            }

            return new ModuleAccessResult
            {
                IsAllowed = ResolveAccountFallbackAccess(moduleCode, operation),
                Scope = ModuleScope.Account,
                HasCompanyContext = true,
                Message = "No tiene permisos sobre esta opcion general."
            };
        }

        var scope = ResolveScope(moduleCode);
        if (scope == ModuleScope.Account)
        {
            return new ModuleAccessResult
            {
                IsAllowed = ResolveAccountFallbackAccess(moduleCode, operation),
                Scope = ModuleScope.Account,
                HasCompanyContext = true,
                Message = "No tiene permisos sobre esta opcion general."
            };
        }

        if (scope == ModuleScope.Company && !hasCompanyContext)
        {
            return new ModuleAccessResult
            {
                IsAllowed = false,
                Scope = ModuleScope.Company,
                HasCompanyContext = false,
                Message = "Debe seleccionar una empresa activa para ingresar a este modulo."
            };
        }

        if (companyPermissions.TryGetValue(moduleCode, out var companyPermission))
        {
            return new ModuleAccessResult
            {
                IsAllowed = EvaluatePermission(companyPermission, operation),
                Scope = ModuleScope.Company,
                HasCompanyContext = hasCompanyContext,
                Message = "No tiene permisos operativos sobre este modulo en la empresa activa."
            };
        }

        return new ModuleAccessResult
        {
            IsAllowed = false,
            Scope = scope,
            HasCompanyContext = hasCompanyContext,
            Message = "No se encontro una configuracion de permisos para este modulo."
        };
    }

    private async Task EnsureInitializedAsync(ClaimsPrincipal principal, CancellationToken cancellationToken)
    {
        if (initialized)
        {
            return;
        }

        initialized = true;

        if (principal.Identity?.IsAuthenticated != true)
        {
            return;
        }

        isSuperAdmin = principal.IsInRole("SuperAdmin");
        if (isSuperAdmin)
        {
            hasAccess = true;
            hasCompanyContext = true;
            return;
        }

        aspNetUserId = principal.FindFirstValue(ClaimTypes.NameIdentifier);
        if (string.IsNullOrWhiteSpace(aspNetUserId))
        {
            return;
        }

        var loginContext = await cuentaAdministradoraRepository.ObtenerContextoLoginUsuarioAsync(aspNetUserId, cancellationToken);
        if (loginContext is null || !loginContext.TieneAcceso || !loginContext.IdCuentaAdministradora.HasValue)
        {
            return;
        }

        hasAccess = true;
        this.loginContext = loginContext;
        idCuentaAdministradora = loginContext.IdCuentaAdministradora.Value;

        var usuariosCuenta = await cuentaAdministradoraRepository.ListarUsuariosCuentaAdministradoraAsync(idCuentaAdministradora.Value, cancellationToken);
        var usuarioCuenta = usuariosCuenta.FirstOrDefault(x => string.Equals(x.AspNetUserId, aspNetUserId, StringComparison.Ordinal));
        if (usuarioCuenta is null)
        {
            hasAccess = false;
            return;
        }

        accountRole = usuarioCuenta.RolCuenta;
        var permisosCuenta = await cuentaAdministradoraRepository.ListarPermisosUsuarioCuentaAsync(usuarioCuenta.IdUsuarioCuentaAdministradora, cancellationToken);
        accountPermissions = permisosCuenta.ToDictionary(x => x.CodigoModulo, StringComparer.OrdinalIgnoreCase);

        if (!currentCompanyAccessor.TieneEmpresaActiva || !currentCompanyAccessor.EmpresaId.HasValue)
        {
            return;
        }

        var empresasUsuario = await cuentaAdministradoraRepository.ListarEmpresasUsuarioCuentaAdministradoraAsync(
            idCuentaAdministradora.Value,
            aspNetUserId,
            cancellationToken);

        var empresaActiva = empresasUsuario.FirstOrDefault(x => x.IdEmpresa == currentCompanyAccessor.EmpresaId.Value && x.IdUsuarioEmpresa.HasValue);
        if (empresaActiva?.IdUsuarioEmpresa is not int idUsuarioEmpresa)
        {
            return;
        }

        hasCompanyContext = true;
        var permisosEmpresa = await cuentaAdministradoraRepository.ListarPermisosUsuarioEmpresaAsync(idUsuarioEmpresa, cancellationToken);
        companyPermissions = permisosEmpresa.ToDictionary(x => x.CodigoModulo, StringComparer.OrdinalIgnoreCase);
    }

    private static bool EvaluatePermission(UsuarioCuentaPermisoDto permission, ModulePermissionOperation operation)
    {
        return operation switch
        {
            ModulePermissionOperation.View => permission.PuedeVerEfectivo,
            ModulePermissionOperation.Create => permission.PuedeCrearEfectivo,
            ModulePermissionOperation.Edit => permission.PuedeEditarEfectivo,
            ModulePermissionOperation.Delete => permission.PuedeEliminarEfectivo,
            ModulePermissionOperation.CreateOrEdit => permission.PuedeCrearEfectivo || permission.PuedeEditarEfectivo,
            _ => permission.PuedeVerEfectivo
        };
    }

    private static bool? EvaluateExplicitPermission(UsuarioCuentaPermisoDto permission, ModulePermissionOperation operation)
    {
        return operation switch
        {
            ModulePermissionOperation.View => ResolveExplicitFlag(permission.PuedeVerOverride, permission.PuedeVerRol),
            ModulePermissionOperation.Create => ResolveExplicitFlag(permission.PuedeCrearOverride, permission.PuedeCrearRol),
            ModulePermissionOperation.Edit => ResolveExplicitFlag(permission.PuedeEditarOverride, permission.PuedeEditarRol),
            ModulePermissionOperation.Delete => ResolveExplicitFlag(permission.PuedeEliminarOverride, permission.PuedeEliminarRol),
            ModulePermissionOperation.CreateOrEdit => ResolveCreateOrEditExplicit(permission),
            _ => ResolveExplicitFlag(permission.PuedeVerOverride, permission.PuedeVerRol)
        };
    }

    private static bool? ResolveExplicitFlag(bool? overrideValue, bool? roleValue)
    {
        if (overrideValue.HasValue)
        {
            return overrideValue.Value;
        }

        if (roleValue.HasValue)
        {
            return roleValue.Value;
        }

        return null;
    }

    private static bool? ResolveCreateOrEditExplicit(UsuarioCuentaPermisoDto permission)
    {
        var create = ResolveExplicitFlag(permission.PuedeCrearOverride, permission.PuedeCrearRol);
        var edit = ResolveExplicitFlag(permission.PuedeEditarOverride, permission.PuedeEditarRol);

        if (!create.HasValue && !edit.HasValue)
        {
            return null;
        }

        return (create ?? false) || (edit ?? false);
    }

    private bool ResolveAccountFallbackAccess(string moduleCode, ModulePermissionOperation operation)
    {
        if (!hasAccess)
        {
            return false;
        }

        if (!IsGeneralFallbackOperationAllowed(moduleCode, operation))
        {
            return false;
        }

        if (string.Equals(moduleCode, "AYUDA", StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        if (string.Equals(moduleCode, "MISUSCRIPCION", StringComparison.OrdinalIgnoreCase))
        {
            return IsAdministrativeRole(accountRole);
        }

        if (string.Equals(moduleCode, "CONFIGURACION", StringComparison.OrdinalIgnoreCase))
        {
            return IsAdministrativeRole(accountRole) || loginContext?.SoloModulosCuenta == true;
        }

        if (string.Equals(moduleCode, "USUARIOS", StringComparison.OrdinalIgnoreCase))
        {
            return IsAdministrativeRole(accountRole);
        }

        if (string.Equals(moduleCode, "EMPRESAS", StringComparison.OrdinalIgnoreCase))
        {
            return loginContext?.SoloModulosCuenta == false
                && (loginContext.CantidadEmpresasAsignadas > 0 || IsAdministrativeRole(accountRole));
        }

        return false;
    }

    private static bool IsGeneralFallbackOperationAllowed(string moduleCode, ModulePermissionOperation operation)
    {
        if (operation == ModulePermissionOperation.View)
        {
            return true;
        }

        if (operation == ModulePermissionOperation.CreateOrEdit)
        {
            return string.Equals(moduleCode, "CONFIGURACION", StringComparison.OrdinalIgnoreCase);
        }

        return false;
    }

    private static bool IsAdministrativeRole(string? roleCode)
    {
        return string.Equals(roleCode, "ADMINISTRADORCUENTA", StringComparison.OrdinalIgnoreCase)
            || string.Equals(roleCode, "SUPERVISOR", StringComparison.OrdinalIgnoreCase);
    }

    private static ModuleScope ResolveScope(string? moduleCode)
    {
        if (string.IsNullOrWhiteSpace(moduleCode))
        {
            return ModuleScope.Unknown;
        }

        if (AccountModules.Contains(moduleCode))
        {
            return ModuleScope.Account;
        }

        if (CompanyModules.Contains(moduleCode))
        {
            return ModuleScope.Company;
        }

        return ModuleScope.Unknown;
    }
}
