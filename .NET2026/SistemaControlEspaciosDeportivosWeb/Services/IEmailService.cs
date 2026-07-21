namespace SistemaControlEspaciosDeportivosWeb.Services;

public interface IEmailService
{
    bool IsEnabled { get; }

    Task SendEmailAsync(
        string toEmail,
        string toName,
        string subject,
        string htmlContent,
        EmailSendOptions? options = null);
}
