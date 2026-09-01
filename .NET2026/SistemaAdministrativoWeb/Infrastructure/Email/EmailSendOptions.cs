namespace SistemaAdministrativoWeb.Infrastructure.Email;

public sealed class EmailSendOptions
{
    public string? SenderEmail { get; set; }
    public string? SenderName { get; set; }
    public List<EmailAttachmentUrlOption> AttachmentUrls { get; set; } = new();
}

public sealed class EmailAttachmentUrlOption
{
    public string Url { get; set; } = string.Empty;
    public string? FileName { get; set; }
}
