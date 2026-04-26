namespace SistemaControlEspaciosDeportivosWeb.Services;

public interface IEmailService
{
    Task SendEmailAsync(
        string toEmail,
        string toName,
        string subject,
        string htmlContent,
        EmailSendOptions? options = null);
}
