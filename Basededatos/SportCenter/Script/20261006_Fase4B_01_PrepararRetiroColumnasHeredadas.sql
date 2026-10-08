-- =============================================
-- Author:        FRANCO LARA
-- Create date:   06/10/2026
-- Description:   Prepara el retiro fisico de IDs locales sin interrumpir el despliegue coordinado.
-- =============================================
-- Orden obligatorio:
-- 1. Confirmar backup verificado y cambiar @RespaldoVerificado a 1.
-- 2. Ejecutar este script.
-- 3. Publicar los SP de Fase 4B y la aplicacion compilada con contratos canonicos.
-- 4. Ejecutar 20261006_Fase4B_02_RetirarColumnasHeredadas.sql.

SET NOCOUNT ON;
SET XACT_ABORT ON;

DECLARE @RespaldoVerificado BIT = 1;

IF @RespaldoVerificado <> 1
BEGIN
    RAISERROR('Ejecucion bloqueada: confirme un backup verificado y establezca @RespaldoVerificado = 1.', 16, 1);
    RETURN;
END;

BEGIN TRY
    IF COL_LENGTH(N'dbo.Negocios', N'CodigoMoneda') IS NULL
        RAISERROR('Fase 4B bloqueada: falta Negocios.CodigoMoneda.', 16, 1);

    IF COL_LENGTH(N'dbo.ComprobantesElectronicos', N'CodigoMoneda') IS NULL
        RAISERROR('Fase 4B bloqueada: falta ComprobantesElectronicos.CodigoMoneda.', 16, 1);

    IF COL_LENGTH(N'dbo.ComprobantesElectronicos', N'CodigoTipoComprobante') IS NULL
        RAISERROR('Fase 4B bloqueada: falta ComprobantesElectronicos.CodigoTipoComprobante.', 16, 1);

    IF EXISTS
    (
        SELECT 1
        FROM dbo.Negocios n
        WHERE n.CodigoMoneda IS NULL
           OR LTRIM(RTRIM(n.CodigoMoneda)) = N''
    )
        RAISERROR('Fase 4B bloqueada: existen negocios sin CodigoMoneda.', 16, 1);

    IF EXISTS
    (
        SELECT 1
        FROM dbo.ComprobantesElectronicos ce
        WHERE ce.CodigoMoneda IS NULL
           OR LTRIM(RTRIM(ce.CodigoMoneda)) = N''
           OR ce.CodigoTipoComprobante IS NULL
           OR LTRIM(RTRIM(ce.CodigoTipoComprobante)) = N''
    )
        RAISERROR('Fase 4B bloqueada: existen comprobantes sin codigos canonicos.', 16, 1);

    IF EXISTS
    (
        SELECT
            ce.NegocioId,
            ce.CodigoTipoComprobante,
            ce.Serie,
            ce.Numero
        FROM dbo.ComprobantesElectronicos ce
        GROUP BY
            ce.NegocioId,
            ce.CodigoTipoComprobante,
            ce.Serie,
            ce.Numero
        HAVING COUNT_BIG(*) > 1
    )
        RAISERROR('Fase 4B bloqueada: existen correlativos duplicados para el codigo SUNAT canonico.', 16, 1);

    BEGIN TRANSACTION;

    ALTER TABLE dbo.Negocios ALTER COLUMN CodigoMoneda NVARCHAR(10) NOT NULL;

    IF NOT EXISTS
    (
        SELECT 1
        FROM sys.default_constraints dc
        INNER JOIN sys.columns c
            ON c.object_id = dc.parent_object_id
           AND c.column_id = dc.parent_column_id
        WHERE dc.parent_object_id = OBJECT_ID(N'dbo.Negocios')
          AND c.name = N'CodigoMoneda'
    )
    BEGIN
        ALTER TABLE dbo.Negocios
            ADD CONSTRAINT DF_Negocios_CodigoMoneda DEFAULT (N'PEN') FOR CodigoMoneda;
    END;

    IF COL_LENGTH(N'dbo.ComprobantesElectronicos', N'TipoComprobante') IS NOT NULL
        ALTER TABLE dbo.ComprobantesElectronicos ALTER COLUMN TipoComprobante INT NULL;

    IF COL_LENGTH(N'dbo.ComprobantesElectronicos', N'TipoMoneda') IS NOT NULL
        ALTER TABLE dbo.ComprobantesElectronicos ALTER COLUMN TipoMoneda INT NULL;

    IF NOT EXISTS
    (
        SELECT 1
        FROM sys.indexes i
        WHERE i.object_id = OBJECT_ID(N'dbo.ComprobantesElectronicos')
          AND i.name = N'UX_ComprobantesElectronicos_Negocio_CodigoTipo_Serie_Numero'
    )
    BEGIN
        CREATE UNIQUE NONCLUSTERED INDEX UX_ComprobantesElectronicos_Negocio_CodigoTipo_Serie_Numero
            ON dbo.ComprobantesElectronicos (NegocioId, CodigoTipoComprobante, Serie, Numero);
    END;

    COMMIT TRANSACTION;

    SELECT
        N'FASE_4B_PREPARADA' AS Resultado,
        SYSUTCDATETIME() AS FechaUtc;
END TRY
BEGIN CATCH
    IF XACT_STATE() <> 0
        ROLLBACK TRANSACTION;

    DECLARE @ErrorMessage NVARCHAR(4000);
    DECLARE @ErrorSeverity INT;
    DECLARE @ErrorState INT;

    SELECT
        @ErrorMessage = ERROR_MESSAGE(),
        @ErrorSeverity = ERROR_SEVERITY(),
        @ErrorState = ERROR_STATE();

    RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
END CATCH;
