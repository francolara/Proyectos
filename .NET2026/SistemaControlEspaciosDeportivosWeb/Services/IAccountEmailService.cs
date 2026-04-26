namespace SistemaControlEspaciosDeportivosWeb.Services;

public interface IAccountEmailService
{
    Task SendConfirmationEmailAsync(string toEmail, string? toName, string confirmationUrl);
    Task SendResetPasswordEmailAsync(string toEmail, string? toName, string resetUrl);
}
