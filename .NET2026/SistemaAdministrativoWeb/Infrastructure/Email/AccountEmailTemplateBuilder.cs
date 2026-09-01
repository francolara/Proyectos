using System.Net;

namespace SistemaAdministrativoWeb.Infrastructure.Email;

public static class AccountEmailTemplateBuilder
{
    public static string BuildConfirmationTemplate(string recipientName, string confirmationUrl, string logoUrl)
        => BuildActionTemplate(
            title: "Confirma tu correo",
            eyebrow: "Tu cuenta esta casi lista",
            greeting: $"Hola {recipientName}",
            description: "Confirma tu correo electronico para activar tu cuenta y comenzar a organizar la gestion contable de tu empresa con FralseCont.",
            buttonText: "Confirmar mi cuenta",
            buttonUrl: confirmationUrl,
            securityNote: "Por tu seguridad, este enlace es personal y solo puede utilizarse una vez.",
            logoUrl: logoUrl);

    public static string BuildResetPasswordTemplate(string recipientName, string resetUrl, string logoUrl)
        => BuildActionTemplate(
            title: "Recupera tu contrasena",
            eyebrow: "Solicitud de seguridad",
            greeting: $"Hola {recipientName}",
            description: "Recibimos una solicitud para crear una nueva contrasena de acceso a FralseCont. Usa el boton para continuar de forma segura.",
            buttonText: "Restablecer contrasena",
            buttonUrl: resetUrl,
            securityNote: "Si no solicitaste este cambio, ignora el mensaje. Tu contrasena actual permanecera sin modificaciones.",
            logoUrl: logoUrl);

    public static string BuildWelcomeTemplate(string recipientName, string loginUrl, string logoUrl)
        => BuildActionTemplate(
            title: "Bienvenido a FralseCont",
            eyebrow: "Cuenta confirmada",
            greeting: $"Hola {recipientName}",
            description: "Tu correo fue confirmado correctamente. Ya puedes ingresar a FralseCont y continuar con la configuracion de tus empresas, usuarios y procesos contables.",
            buttonText: "Ingresar a FralseCont",
            buttonUrl: loginUrl,
            securityNote: "FralseCont es una solucion de FRALSE TECH S.A.C. para empresas y contadores.",
            logoUrl: logoUrl);

    private static string BuildActionTemplate(
        string title,
        string eyebrow,
        string greeting,
        string description,
        string buttonText,
        string buttonUrl,
        string securityNote,
        string logoUrl)
    {
        var safeButtonUrl = Escape(buttonUrl);
        var logoMarkup = string.IsNullOrWhiteSpace(logoUrl)
            ? "<div style=\"font-size:22px;font-weight:900;letter-spacing:-.02em;color:#0b2040;\">Fralse<span style=\"color:#0d6efd;\">Cont</span></div>"
            : $"<img src=\"{Escape(logoUrl)}\" width=\"178\" alt=\"FRALSE TECH S.A.C.\" style=\"display:block;width:178px;max-width:100%;height:auto;border:0;\">";

        return
$"""
<!doctype html>
<html lang="es">
  <head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width,initial-scale=1">
    <title>{Escape(title)}</title>
  </head>
  <body style="margin:0;padding:0;background:#eef4fb;font-family:'Segoe UI',Arial,sans-serif;color:#122033;">
    <div style="display:none;max-height:0;overflow:hidden;opacity:0;">{Escape(description)}</div>
    <table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="width:100%;background:#eef4fb;padding:28px 12px;">
      <tr>
        <td align="center">
          <table role="presentation" width="620" cellspacing="0" cellpadding="0" style="width:100%;max-width:620px;background:#ffffff;border:1px solid #d8e4f2;border-radius:18px;overflow:hidden;box-shadow:0 16px 42px rgba(15,35,66,.10);">
            <tr>
              <td style="height:8px;background:linear-gradient(90deg,#12376f 0%,#0d6efd 58%,#25b7e8 100%);font-size:0;line-height:0;">&nbsp;</td>
            </tr>
            <tr>
              <td style="padding:24px 30px 18px;border-bottom:1px solid #e4ebf4;">
                <table role="presentation" width="100%" cellspacing="0" cellpadding="0">
                  <tr>
                    <td>{logoMarkup}</td>
                    <td align="right" style="font-size:12px;font-weight:800;letter-spacing:.08em;text-transform:uppercase;color:#0d6efd;">FralseCont</td>
                  </tr>
                </table>
              </td>
            </tr>
            <tr>
              <td style="padding:34px 30px 30px;">
                <div style="margin:0 0 10px;font-size:12px;font-weight:800;letter-spacing:.10em;text-transform:uppercase;color:#1f75d6;">{Escape(eyebrow)}</div>
                <h1 style="margin:0 0 18px;font-size:30px;line-height:1.16;color:#0b2040;letter-spacing:-.02em;">{Escape(title)}</h1>
                <p style="margin:0 0 12px;font-size:18px;line-height:1.5;font-weight:700;color:#173962;">{Escape(greeting)},</p>
                <p style="margin:0 0 26px;font-size:15px;line-height:1.7;color:#536983;">{Escape(description)}</p>
                <table role="presentation" cellspacing="0" cellpadding="0" style="margin:0 0 24px;">
                  <tr>
                    <td style="border-radius:12px;background:#0d6efd;">
                      <a href="{safeButtonUrl}" style="display:inline-block;padding:14px 24px;border-radius:12px;color:#ffffff;font-size:15px;font-weight:800;text-decoration:none;">{Escape(buttonText)}</a>
                    </td>
                  </tr>
                </table>
                <div style="padding:15px 16px;border-radius:12px;background:#f3f7fc;border:1px solid #dce7f4;">
                  <p style="margin:0 0 7px;font-size:12px;line-height:1.55;color:#607089;">Si el boton no funciona, copia y pega este enlace en tu navegador:</p>
                  <p style="margin:0;word-break:break-all;font-size:12px;line-height:1.55;color:#0d6efd;">{Escape(buttonUrl)}</p>
                </div>
                <p style="margin:20px 0 0;font-size:12px;line-height:1.6;color:#7a8ba3;">{Escape(securityNote)}</p>
              </td>
            </tr>
            <tr>
              <td style="padding:20px 30px;background:#0f2342;color:#c8deff;">
                <p style="margin:0 0 5px;font-size:13px;font-weight:800;color:#ffffff;">FralseCont</p>
                <p style="margin:0;font-size:11px;line-height:1.55;">Tu contabilidad en la nube, simple y segura.<br>Correo generado automaticamente; no respondas a este mensaje.</p>
              </td>
            </tr>
          </table>
        </td>
      </tr>
    </table>
  </body>
</html>
""";
    }

    private static string Escape(string? value)
        => WebUtility.HtmlEncode(value ?? string.Empty).Replace("\"", "&quot;", StringComparison.Ordinal);
}
