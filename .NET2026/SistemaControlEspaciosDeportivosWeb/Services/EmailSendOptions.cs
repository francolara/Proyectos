namespace SistemaControlEspaciosDeportivosWeb.Services;

public class EmailSendOptions
{
    public string? SenderEmail { get; set; }
    public string? SenderName { get; set; }
    public List<EmailAttachmentUrlOption> AttachmentUrls { get; set; } = new();
}

public class EmailAttachmentUrlOption
{
    public string Url { get; set; } = string.Empty;
    public string? FileName { get; set; }
}
