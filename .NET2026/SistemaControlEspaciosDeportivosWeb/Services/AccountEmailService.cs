namespace SistemaControlEspaciosDeportivosWeb.Services;

public class AccountEmailService(IEmailService emailService) : IAccountEmailService
{
    private const string SenderEmail = "noreply@lazonadeportiva.com";
    private const string SenderName = "La Zona Deportiva";

    public Task SendConfirmationEmailAsync(string toEmail, string? toName, string confirmationUrl)
    {
        var html = AccountEmailTemplateBuilder.BuildConfirmEmailTemplate(
            string.IsNullOrWhiteSpace(toName) ? toEmail : toName!,
            confirmationUrl);

        return emailService.SendEmailAsync(
            toEmail,
            string.IsNullOrWhiteSpace(toName) ? toEmail : toName!,
            "Confirma tu cuenta - La Zona Deportiva",
            html,
            new EmailSendOptions
            {
                SenderEmail = SenderEmail,
                SenderName = SenderName
            });
    }

    public Task SendResetPasswordEmailAsync(string toEmail, string? toName, string resetUrl)
    {
        var html = AccountEmailTemplateBuilder.BuildResetPasswordTemplate(
            string.IsNullOrWhiteSpace(toName) ? toEmail : toName!,
            resetUrl);

        return emailService.SendEmailAsync(
            toEmail,
            string.IsNullOrWhiteSpace(toName) ? toEmail : toName!,
            "Recupera tu contrasena - La Zona Deportiva",
            html,
            new EmailSendOptions
            {
                SenderEmail = SenderEmail,
                SenderName = SenderName
            });
    }
}
