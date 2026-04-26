namespace SistemaControlEspaciosDeportivosWeb.Services;

public static class AccountEmailTemplateBuilder
{
    public static string BuildConfirmEmailTemplate(string nombreDestino, string confirmationUrl)
    {
        return BuildBaseTemplate(
            "Confirma tu cuenta",
            $"Hola {Escape(nombreDestino)}, ya casi terminas tu registro en La Zona Deportiva.",
            "Confirma tu correo para activar tu cuenta y empezar a gestionar reservas.",
            "Confirmar cuenta",
            confirmationUrl);
    }

    public static string BuildResetPasswordTemplate(string nombreDestino, string resetUrl)
    {
        return BuildBaseTemplate(
            "Recupera tu contrasena",
            $"Hola {Escape(nombreDestino)}, recibimos una solicitud para restablecer tu contrasena.",
            "Haz clic en el siguiente boton para crear una nueva contrasena de forma segura.",
            "Restablecer contrasena",
            resetUrl);
    }

    private static string BuildBaseTemplate(
        string title,
        string intro,
        string body,
        string buttonText,
        string buttonUrl)
    {
        var safeUrl = EscapeAttribute(buttonUrl);
        return
$"""
<!doctype html>
<html lang="es">
  <body style="margin:0;padding:0;background-color:#f4f7fb;font-family:Manrope,'Segoe UI',Arial,sans-serif;color:#1f2937;">
    <table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="background-color:#f4f7fb;padding:24px 12px;">
      <tr>
        <td align="center">
          <table role="presentation" width="600" cellspacing="0" cellpadding="0" style="max-width:600px;background:#ffffff;border:1px solid #dbe6f4;border-radius:14px;overflow:hidden;">
            <tr>
              <td style="background:linear-gradient(135deg,#0d3b66 0%,#164f86 60%,#17a2b8 100%);padding:22px 24px;">
                <div style="font-size:12px;letter-spacing:.08em;text-transform:uppercase;color:#dbeafe;font-weight:800;">La Zona Deportiva</div>
                <h1 style="margin:8px 0 0;font-size:28px;line-height:1.1;color:#ffffff;">{Escape(title)}</h1>
              </td>
            </tr>
            <tr>
              <td style="padding:26px 24px 30px;">
                <p style="margin:0 0 10px;font-size:18px;line-height:1.45;color:#0f172a;font-weight:700;">{intro}</p>
                <p style="margin:0 0 22px;font-size:15px;line-height:1.65;color:#475569;">{body}</p>
                <table role="presentation" cellspacing="0" cellpadding="0" style="margin:0 0 18px;">
                  <tr>
                    <td style="border-radius:999px;background:linear-gradient(135deg,#1f66dc 0%,#1d4ed8 100%);">
                      <a href="{safeUrl}" style="display:inline-block;padding:13px 24px;color:#ffffff;font-weight:800;font-size:15px;text-decoration:none;border-radius:999px;">{Escape(buttonText)}</a>
                    </td>
                  </tr>
                </table>
                <p style="margin:0;font-size:13px;line-height:1.55;color:#64748b;">Si el boton no funciona, copia y pega este enlace en tu navegador:</p>
                <p style="margin:6px 0 0;word-break:break-all;font-size:13px;line-height:1.55;color:#1d4ed8;">{Escape(buttonUrl)}</p>
                <hr style="margin:22px 0 14px;border:0;border-top:1px solid #e2e8f0;">
                <p style="margin:0;font-size:12px;line-height:1.5;color:#94a3b8;">Este correo fue enviado automaticamente por La Zona Deportiva.</p>
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

    private static string Escape(string value)
    {
        return System.Net.WebUtility.HtmlEncode(value ?? string.Empty);
    }

    private static string EscapeAttribute(string value)
    {
        return Escape(value).Replace("\"", "&quot;");
    }
}
