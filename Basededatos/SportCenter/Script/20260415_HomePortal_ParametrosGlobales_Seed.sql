USE [DbSportCenter]
GO

-- Firma: Codex - 15/04/2026 | Seed inicial de parametros globales HOME_PORTAL_* para beneficios, CTA y barra final del Home.

SET NOCOUNT ON;

DECLARE @Usuario NVARCHAR(120) = N'seed-codex';

DECLARE @Params TABLE
(
    NombreParametro NVARCHAR(100) NOT NULL,
    Descripcion NVARCHAR(500) NOT NULL,
    ValorParametro NVARCHAR(100) NOT NULL
);

INSERT INTO @Params (NombreParametro, Descripcion, ValorParametro)
VALUES
(N'HOME_PORTAL_BENEF_TITULO', N'HOME_PORTAL_BENEF_TITULO', N'Todo lo que necesitas para gestionar tus canchas deportivas'),
(N'HOME_PORTAL_BENEF_SUBTITULO', N'HOME_PORTAL_BENEF_SUBTITULO', N'SportCenter integra reservas, sedes, pagos y reportes para crecer tu operacion.'),
(N'HOME_PORTAL_BENEF_1_TITULO', N'HOME_PORTAL_BENEF_1_TITULO', N'Sistema de reservas'),
(N'HOME_PORTAL_BENEF_1_DETALLE', N'HOME_PORTAL_BENEF_1_DETALLE', N'Controla la disponibilidad por horario con agenda visual y registro de clientes.'),
(N'HOME_PORTAL_BENEF_2_TITULO', N'HOME_PORTAL_BENEF_2_TITULO', N'Multiples sedes'),
(N'HOME_PORTAL_BENEF_2_DETALLE', N'HOME_PORTAL_BENEF_2_DETALLE', N'Administra distintos complejos deportivos desde un solo panel operativo.'),
(N'HOME_PORTAL_BENEF_3_TITULO', N'HOME_PORTAL_BENEF_3_TITULO', N'Pagos seguros'),
(N'HOME_PORTAL_BENEF_3_DETALLE', N'HOME_PORTAL_BENEF_3_DETALLE', N'Gestiona adelantos, saldos y comprobantes con trazabilidad por reserva.'),
(N'HOME_PORTAL_BENEF_4_TITULO', N'HOME_PORTAL_BENEF_4_TITULO', N'Promociones especiales'),
(N'HOME_PORTAL_BENEF_4_DETALLE', N'HOME_PORTAL_BENEF_4_DETALLE', N'Crea descuentos por sede, dia y horario para mejorar ocupacion.'),
(N'HOME_PORTAL_BENEF_5_TITULO', N'HOME_PORTAL_BENEF_5_TITULO', N'Estadisticas detalladas'),
(N'HOME_PORTAL_BENEF_5_DETALLE', N'HOME_PORTAL_BENEF_5_DETALLE', N'Analiza ingresos, ocupacion y rendimiento para decisiones con datos.'),
(N'HOME_PORTAL_BENEF_6_TITULO', N'HOME_PORTAL_BENEF_6_TITULO', N'Mayor visibilidad'),
(N'HOME_PORTAL_BENEF_6_DETALLE', N'HOME_PORTAL_BENEF_6_DETALLE', N'Publica tu negocio en el portal y recibe nuevas solicitudes online.'),
(N'HOME_PORTAL_CTA_TITULO', N'HOME_PORTAL_CTA_TITULO', N'Unete a la comunidad de SportCenter'),
(N'HOME_PORTAL_CTA_SUBTITULO', N'HOME_PORTAL_CTA_SUBTITULO', N'Registra tu club deportivo y comienza a gestionar tus canchas de manera eficiente.'),
(N'HOME_PORTAL_CTA_BTN_CLUB_TEXTO', N'HOME_PORTAL_CTA_BTN_CLUB_TEXTO', N'Registrar mi club'),
(N'HOME_PORTAL_CTA_BTN_CLUB_URL', N'HOME_PORTAL_CTA_BTN_CLUB_URL', N'/Home/SoftwareClubes'),
(N'HOME_PORTAL_CTA_BTN_USUARIO_TEXTO', N'HOME_PORTAL_CTA_BTN_USUARIO_TEXTO', N'Crear cuenta personal'),
(N'HOME_PORTAL_CTA_BTN_USUARIO_URL', N'HOME_PORTAL_CTA_BTN_USUARIO_URL', N'/Identity/Account/Register'),
(N'HOME_PORTAL_MARCA_TITULO', N'HOME_PORTAL_MARCA_TITULO', N'SportCenter'),
(N'HOME_PORTAL_MARCA_DESC', N'HOME_PORTAL_MARCA_DESC', N'La plataforma lider para la reserva y gestion de canchas deportivas.'),
(N'HOME_PORTAL_CONTACTO_EMAIL', N'HOME_PORTAL_CONTACTO_EMAIL', N'contacto@sportcenter.com'),
(N'HOME_PORTAL_CONTACTO_TELEFONO', N'HOME_PORTAL_CONTACTO_TELEFONO', N'+51 900 000 000'),
(N'HOME_PORTAL_FACEBOOK_URL', N'HOME_PORTAL_FACEBOOK_URL', N''),
(N'HOME_PORTAL_INSTAGRAM_URL', N'HOME_PORTAL_INSTAGRAM_URL', N''),
(N'HOME_PORTAL_WHATSAPP_URL', N'HOME_PORTAL_WHATSAPP_URL', N'');

MERGE dbo.ParametrosGlobales AS T
USING @Params AS S
ON T.NombreParametro = S.NombreParametro
WHEN MATCHED THEN
    UPDATE SET
        T.Descripcion = S.Descripcion,
        T.ValorParametro = S.ValorParametro
WHEN NOT MATCHED BY TARGET THEN
    INSERT (NombreParametro, Descripcion, ValorParametro)
    VALUES (S.NombreParametro, S.Descripcion, S.ValorParametro);

SELECT NombreParametro, Descripcion, ValorParametro
FROM dbo.ParametrosGlobales
WHERE NombreParametro LIKE N'HOME_PORTAL_%'
ORDER BY NombreParametro;
