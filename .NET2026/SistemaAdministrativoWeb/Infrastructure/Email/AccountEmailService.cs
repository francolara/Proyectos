using SistemaAdministrativoWeb.Infrastructure.Assets;

namespace SistemaAdministrativoWeb.Infrastructure.Email;

public sealed class AccountEmailService(
    IEmailService emailService,
    IConfiguration configuration) : IAccountEmailService
{
    public bool IsEnabled => emailService.IsEnabled;

    public Task SendConfirmationEmailAsync(
        string toEmail,
        string? toName,
        string confirmationUrl,
        CancellationToken cancellationToken = default)
    {
        var recipientName = ResolveRecipientName(toEmail, toName);
        var html = AccountEmailTemplateBuilder.BuildConfirmationTemplate(
            recipientName,
            confirmationUrl,
            ResolveLogoUrl());

        return emailService.SendEmailAsync(
            toEmail,
            recipientName,
            "Confirma tu cuenta de FralseCont",
            html,
            cancellationToken: cancellationToken);
    }

    public Task SendResetPasswordEmailAsync(
        string toEmail,
        string? toName,
        string resetUrl,
        CancellationToken cancellationToken = default)
    {
        var recipientName = ResolveRecipientName(toEmail, toName);
        var html = AccountEmailTemplateBuilder.BuildResetPasswordTemplate(
            recipientName,
            resetUrl,
            ResolveLogoUrl());

        return emailService.SendEmailAsync(
            toEmail,
            recipientName,
            "Recupera tu contrasena de FralseCont",
            html,
            cancellationToken: cancellationToken);
    }

    public Task SendWelcomeEmailAsync(
        string toEmail,
        string? toName,
        string loginUrl,
        CancellationToken cancellationToken = default)
    {
        var recipientName = ResolveRecipientName(toEmail, toName);
        var html = AccountEmailTemplateBuilder.BuildWelcomeTemplate(
            recipientName,
            loginUrl,
            ResolveLogoUrl());

        return emailService.SendEmailAsync(
            toEmail,
            recipientName,
            "Bienvenido a FralseCont",
            html,
            cancellationToken: cancellationToken);
    }

    private string ResolveLogoUrl()
        => PublicAssetUrlBuilder.Build(configuration, "Logo/LogoFralseTech.webp");

    private static string ResolveRecipientName(string toEmail, string? toName)
        => string.IsNullOrWhiteSpace(toName) ? toEmail.Trim() : toName.Trim();
}
