namespace SistemaAdministrativoWeb.Infrastructure.Email;

public interface IAccountEmailService
{
    bool IsEnabled { get; }

    Task SendConfirmationEmailAsync(
        string toEmail,
        string? toName,
        string confirmationUrl,
        CancellationToken cancellationToken = default);

    Task SendResetPasswordEmailAsync(
        string toEmail,
        string? toName,
        string resetUrl,
        CancellationToken cancellationToken = default);

    Task SendWelcomeEmailAsync(
        string toEmail,
        string? toName,
        string loginUrl,
        CancellationToken cancellationToken = default);
}
