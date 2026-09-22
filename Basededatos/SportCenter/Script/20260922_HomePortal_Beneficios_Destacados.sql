USE [DbSportCenter]
GO

-- Firma: Codex - 22/09/2026 | Actualiza los beneficios destacados del portal para comunicar reservas sin cruces, tarifas, comprobantes, promociones, reportes y gestion multideportiva.
SET NOCOUNT ON;
SET XACT_ABORT ON;

BEGIN TRANSACTION;

DECLARE @Parametros TABLE
(
    NombreParametro NVARCHAR(100) NOT NULL,
    ValorParametro NVARCHAR(500) NOT NULL
);

INSERT INTO @Parametros (NombreParametro, ValorParametro)
VALUES
    (N'HOME_PORTAL_BENEF_TITULO', N'Gestiona todo tu complejo deportivo desde un solo lugar'),
    (N'HOME_PORTAL_BENEF_SUBTITULO', N'Reservas, tarifas, clientes, cobros y reportes para operar con control total.'),
    (N'HOME_PORTAL_BENEF_1_TITULO', N'Reservas sin cruces'),
    (N'HOME_PORTAL_BENEF_1_DETALLE', N'Evita horarios duplicados y bloquea automaticamente los espacios compartidos.'),
    (N'HOME_PORTAL_BENEF_2_TITULO', N'Tarifas inteligentes'),
    (N'HOME_PORTAL_BENEF_2_DETALLE', N'Define precios por espacio y turno, con tarifas especiales para domingos y feriados.'),
    (N'HOME_PORTAL_BENEF_3_TITULO', N'Cobros y comprobantes'),
    (N'HOME_PORTAL_BENEF_3_DETALLE', N'Registra adelantos y saldos, y emite comprobantes electronicos por cada reserva.'),
    (N'HOME_PORTAL_BENEF_4_TITULO', N'Promociones y cupones'),
    (N'HOME_PORTAL_BENEF_4_DETALLE', N'Crea descuentos y cupones para impulsar tus horarios de menor demanda.'),
    (N'HOME_PORTAL_BENEF_5_TITULO', N'Clientes y reportes'),
    (N'HOME_PORTAL_BENEF_5_DETALLE', N'Revisa ingresos diarios, cancelaciones y el consumo de cada cliente.'),
    (N'HOME_PORTAL_BENEF_6_TITULO', N'Multiples deportes y espacios'),
    (N'HOME_PORTAL_BENEF_6_DETALLE', N'Administra sedes, canchas y disciplinas deportivas desde un solo panel.');

IF EXISTS
(
    SELECT 1
    FROM @Parametros p
    LEFT JOIN dbo.ParametrosGlobales pg
        ON pg.NombreParametro = p.NombreParametro
    WHERE pg.ParametroId IS NULL
)
BEGIN
    THROW 50001, 'Faltan parametros HOME_PORTAL_BENEF_*. Ejecute primero el seed inicial del portal.', 1;
END

UPDATE pg
SET pg.ValorParametro = p.ValorParametro
FROM dbo.ParametrosGlobales pg
INNER JOIN @Parametros p
    ON p.NombreParametro = pg.NombreParametro;

COMMIT TRANSACTION;

SELECT pg.NombreParametro, pg.ValorParametro
FROM dbo.ParametrosGlobales pg
INNER JOIN @Parametros p
    ON p.NombreParametro = pg.NombreParametro
ORDER BY pg.NombreParametro;
