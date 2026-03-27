namespace SistemaControlEspaciosDeportivosWeb.Services;

public class AutomationSettings
{
    public bool Enabled { get; set; }
    public int IntervalSeconds { get; set; } = 300;
    public string UsuarioSistema { get; set; } = "worker_auto";
}
