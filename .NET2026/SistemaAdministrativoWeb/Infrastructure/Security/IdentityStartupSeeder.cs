using Microsoft.AspNetCore.Identity;
using Microsoft.Extensions.Options;
using SistemaAdministrativoWeb.Configuration;

namespace SistemaAdministrativoWeb.Infrastructure.Security;

public sealed class IdentityStartupSeeder(
    RoleManager<IdentityRole> roleManager,
    UserManager<IdentityUser> userManager,
    IOptions<IdentitySeedOptions> options)
{
    private static readonly string[] RolesBase =
    [
        "SuperAdmin",
        "AdministradorEmpresa"
    ];

    public async Task SeedAsync(CancellationToken cancellationToken = default)
    {
        foreach (var roleName in RolesBase)
        {
            cancellationToken.ThrowIfCancellationRequested();

            if (await roleManager.RoleExistsAsync(roleName))
            {
                continue;
            }

            await roleManager.CreateAsync(new IdentityRole(roleName));
        }

        foreach (var email in options.Value.SuperAdminEmails
                     .Where(x => !string.IsNullOrWhiteSpace(x))
                     .Select(x => x.Trim())
                     .Distinct(StringComparer.OrdinalIgnoreCase))
        {
            cancellationToken.ThrowIfCancellationRequested();

            var user = await userManager.FindByEmailAsync(email);
            if (user is null)
            {
                continue;
            }

            if (!await userManager.IsInRoleAsync(user, "SuperAdmin"))
            {
                await userManager.AddToRoleAsync(user, "SuperAdmin");
            }
        }
    }
}
