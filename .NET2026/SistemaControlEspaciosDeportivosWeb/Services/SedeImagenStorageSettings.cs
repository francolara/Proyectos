namespace SistemaControlEspaciosDeportivosWeb.Services;

public class SedeImagenStorageSettings
{
    public bool Enabled { get; set; }
    public string Provider { get; set; } = "R2";
    public string Endpoint { get; set; } = string.Empty;
    public string BucketName { get; set; } = string.Empty;
    public string AccessKey { get; set; } = string.Empty;
    public string SecretKey { get; set; } = string.Empty;
    public string Region { get; set; } = "us-east-1";
    public string PublicBaseUrl { get; set; } = string.Empty;
    public int MaxImageBytes { get; set; } = 12 * 1024 * 1024;
    public int MaxOutputBytes { get; set; } = 400 * 1024;
    public int TargetWidth { get; set; } = 1200;
    public int TargetHeight { get; set; } = 900;
}
