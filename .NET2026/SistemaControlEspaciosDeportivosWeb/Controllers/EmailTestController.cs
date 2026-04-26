using Microsoft.AspNetCore.Mvc;
using SistemaControlEspaciosDeportivosWeb.Services;

namespace SistemaControlEspaciosDeportivosWeb.Controllers;

[ApiController]
[Route("dev/email-test")]
public class EmailTestController(
    IEmailService emailService,
    IWebHostEnvironment environment,
    ILogger<EmailTestController> logger) : ControllerBase
{
    [HttpPost("send")]
    public async Task<IActionResult> Send(
        [FromQuery] string toEmail,
        [FromQuery] string? toName,
        [FromQuery] string? senderEmail,
        [FromQuery] string? senderName,
        [FromQuery] string[]? attachmentUrl)
    {
        if (!environment.IsDevelopment())
        {
            return NotFound();
        }

        if (string.IsNullOrWhiteSpace(toEmail))
        {
            return BadRequest(new { ok = false, mensaje = "Debes indicar el parametro toEmail." });
        }

        const string subject = "Prueba de correo - La Zona Deportiva";
        const string htmlContent =
"""
<h2>La Zona Deportiva</h2>
<p>Este es un correo de prueba enviado desde Brevo.</p>
<p><small>Este correo fue enviado automaticamente.</small></p>
""";

        var options = new EmailSendOptions
        {
            SenderEmail = senderEmail,
            SenderName = senderName
        };
        if (attachmentUrl is not null && attachmentUrl.Length > 0)
        {
            options.AttachmentUrls = attachmentUrl
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Select(x => new EmailAttachmentUrlOption { Url = x })
                .ToList();
        }

        try
        {
            await emailService.SendEmailAsync(
                toEmail,
                string.IsNullOrWhiteSpace(toName) ? toEmail : toName,
                subject,
                htmlContent,
                options);

            return Ok(new { ok = true, mensaje = $"Correo de prueba enviado a {toEmail}." });
        }
        catch (EmailDeliveryException ex)
        {
            logger.LogWarning(ex, "No se pudo enviar correo de prueba a {ToEmail}.", toEmail);
            return StatusCode(StatusCodes.Status502BadGateway, new { ok = false, mensaje = "No se pudo enviar el correo de prueba." });
        }
    }
}
