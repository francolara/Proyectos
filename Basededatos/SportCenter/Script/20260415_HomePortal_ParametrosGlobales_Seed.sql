USE [DbSportCenter]
GO

-- Firma: Codex - 15/04/2026 | Seed inicial de parametros globales HOME_PORTAL_* para beneficios, CTA y barra final del Home.
-- Firma: Codex - 22/09/2026 | Actualiza los textos iniciales de beneficios para comunicar la operacion integral de complejos deportivos.

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
(N'HOME_PORTAL_BENEF_TITULO', N'HOME_PORTAL_BENEF_TITULO', N'Gestiona todo tu complejo deportivo desde un solo lugar'),
(N'HOME_PORTAL_BENEF_SUBTITULO', N'HOME_PORTAL_BENEF_SUBTITULO', N'Reservas, tarifas, clientes, cobros y reportes para operar con control total.'),
(N'HOME_PORTAL_BENEF_1_TITULO', N'HOME_PORTAL_BENEF_1_TITULO', N'Reservas sin cruces'),
(N'HOME_PORTAL_BENEF_1_DETALLE', N'HOME_PORTAL_BENEF_1_DETALLE', N'Evita horarios duplicados y bloquea automaticamente los espacios compartidos.'),
(N'HOME_PORTAL_BENEF_2_TITULO', N'HOME_PORTAL_BENEF_2_TITULO', N'Tarifas inteligentes'),
(N'HOME_PORTAL_BENEF_2_DETALLE', N'HOME_PORTAL_BENEF_2_DETALLE', N'Define precios por espacio y turno, con tarifas especiales para domingos y feriados.'),
(N'HOME_PORTAL_BENEF_3_TITULO', N'HOME_PORTAL_BENEF_3_TITULO', N'Cobros y comprobantes'),
(N'HOME_PORTAL_BENEF_3_DETALLE', N'HOME_PORTAL_BENEF_3_DETALLE', N'Registra adelantos y saldos, y emite comprobantes electronicos por cada reserva.'),
(N'HOME_PORTAL_BENEF_4_TITULO', N'HOME_PORTAL_BENEF_4_TITULO', N'Promociones y cupones'),
(N'HOME_PORTAL_BENEF_4_DETALLE', N'HOME_PORTAL_BENEF_4_DETALLE', N'Crea descuentos y cupones para impulsar tus horarios de menor demanda.'),
(N'HOME_PORTAL_BENEF_5_TITULO', N'HOME_PORTAL_BENEF_5_TITULO', N'Clientes y reportes'),
(N'HOME_PORTAL_BENEF_5_DETALLE', N'HOME_PORTAL_BENEF_5_DETALLE', N'Revisa ingresos diarios, cancelaciones y el consumo de cada cliente.'),
(N'HOME_PORTAL_BENEF_6_TITULO', N'HOME_PORTAL_BENEF_6_TITULO', N'Multiples deportes y espacios'),
(N'HOME_PORTAL_BENEF_6_DETALLE', N'HOME_PORTAL_BENEF_6_DETALLE', N'Administra sedes, canchas y disciplinas deportivas desde un solo panel.'),
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
