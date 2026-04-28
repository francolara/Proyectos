namespace SistemaControlEspaciosDeportivosWeb.Services;

public class JobsSettings
{
    public string Token { get; set; } = string.Empty;
    public bool AutoCancelEnabled { get; set; }
    public string UsuarioSistema { get; set; } = "job_scheduler";
    public string[] AllowedEnvironments { get; set; } = ["Production", "Staging"];

    public bool IsEnvironmentAllowed(string? environmentName)
    {
        if (string.IsNullOrWhiteSpace(environmentName) || AllowedEnvironments is null || AllowedEnvironments.Length == 0)
            return false;

        return AllowedEnvironments.Any(x =>
            !string.IsNullOrWhiteSpace(x) &&
            string.Equals(x.Trim(), environmentName.Trim(), StringComparison.OrdinalIgnoreCase));
    }
}

