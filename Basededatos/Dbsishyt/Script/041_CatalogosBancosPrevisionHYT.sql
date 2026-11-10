-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/11/2026
-- Description:   Carga entidades bancarias y previsionales para Planillas.
-- =============================================
USE [SistemaWebHYT];
GO
-- Fuentes: directorio de captación de depósitos y listado de AFP de la SBS.
INSERT INTO dbo.BancoPeruHYT (Nombre)
SELECT nuevo.Nombre FROM (VALUES
    (N'Alfin Banco'),
    (N'BanBif'),
    (N'Banco Bci'),
    (N'Banco de Crédito del Perú (BCP)'),
    (N'Banco Efectiva'),
    (N'Banco Falabella'),
    (N'Banco GNB'),
    (N'Banco Pichincha'),
    (N'Banco Ripley'),
    (N'Bancom'),
    (N'Bank of China (Perú)'),
    (N'BBVA Perú'),
    (N'Santander Consumer'),
    (N'Citibank del Perú'),
    (N'Compartamos Banco'),
    (N'ICBC Peru Bank'),
    (N'Interbank'),
    (N'Mibanco'),
    (N'Banco de la Nación'),
    (N'Santander Perú'),
    (N'Scotiabank Perú')
) AS nuevo (Nombre)
WHERE NOT EXISTS (SELECT 1 FROM dbo.BancoPeruHYT AS banco WHERE banco.Nombre = nuevo.Nombre);

INSERT INTO dbo.EntidadPrevisionalHYT (Nombre, Regimen)
SELECT nuevo.Nombre, nuevo.Regimen FROM (VALUES
    (N'AFP Hábitat', 'AFP'), (N'AFP Integra', 'AFP'), (N'Prima AFP', 'AFP'),
    (N'Profuturo AFP', 'AFP'), (N'ONP', 'ONP')
) AS nuevo (Nombre, Regimen)
WHERE NOT EXISTS (SELECT 1 FROM dbo.EntidadPrevisionalHYT AS entidad WHERE entidad.Nombre = nuevo.Nombre);
GO
