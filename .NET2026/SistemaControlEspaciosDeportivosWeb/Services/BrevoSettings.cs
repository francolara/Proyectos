namespace SistemaControlEspaciosDeportivosWeb.Services;

public class BrevoSettings
{
    public string ApiKey { get; set; } = string.Empty;
    public string SenderEmail { get; set; } = string.Empty;
    public string SenderName { get; set; } = string.Empty;
    public List<string> AllowedSenderEmails { get; set; } = new();
    public int AttachmentMaxBytes { get; set; } = 5242880;
    public int AttachmentDownloadTimeoutSeconds { get; set; } = 20;
    public List<string> AllowedAttachmentContentTypes { get; set; } = new()
    {
        "application/pdf",
        "image/jpeg",
        "image/png",
        "text/plain"
    };
}
