-- =============================================
-- Author:        FRANCO LARA
-- Create date:   06/10/2026
-- Description:   Valida el retiro de dependencias runtime de IDs locales antes de la Fase 4B.
-- Firma:         FRANCO LARA - 08/10/2026 | Permite negocios pendientes de configuracion aunque conserven un CodigoMoneda aparente sin asociacion activa.
-- =============================================

SET NOCOUNT ON;
SET XACT_ABORT ON;

BEGIN TRY
    IF EXISTS (SELECT 1 FROM dbo.Reservas WHERE CodigoMoneda IS NULL OR LTRIM(RTRIM(CodigoMoneda)) = N'')
        RAISERROR('Fase 4A bloqueada: existen reservas sin CodigoMoneda.', 16, 1);

    IF EXISTS (SELECT 1 FROM dbo.Pagos WHERE CodigoMoneda IS NULL OR LTRIM(RTRIM(CodigoMoneda)) = N'')
        RAISERROR('Fase 4A bloqueada: existen pagos sin CodigoMoneda.', 16, 1);

    IF EXISTS
    (
        SELECT 1
        FROM dbo.ComprobantesElectronicos ce
        WHERE ce.CodigoMoneda IS NULL
           OR LTRIM(RTRIM(ce.CodigoMoneda)) = N''
           OR ce.CodigoTipoComprobante IS NULL
           OR LTRIM(RTRIM(ce.CodigoTipoComprobante)) = N''
    )
        RAISERROR('Fase 4A bloqueada: existen comprobantes sin codigos canonicos.', 16, 1);

    IF COALESCE(OBJECT_DEFINITION(OBJECT_ID(N'dbo.Sp_ConfiguracionClub_Actualizar')), N'') NOT LIKE N'%@CodigoMoneda%'
        RAISERROR('Fase 4A bloqueada: Sp_ConfiguracionClub_Actualizar no usa @CodigoMoneda.', 16, 1);

    IF COALESCE(OBJECT_DEFINITION(OBJECT_ID(N'dbo.Sp_Comprobantes_Crear')), N'') LIKE N'%@TipoComprobante INT%'
        RAISERROR('Fase 4A bloqueada: Sp_Comprobantes_Crear aun expone @TipoComprobante.', 16, 1);

    IF COALESCE(OBJECT_DEFINITION(OBJECT_ID(N'dbo.Sp_Comprobantes_Crear')), N'') LIKE N'%@TipoMoneda INT%'
        RAISERROR('Fase 4A bloqueada: Sp_Comprobantes_Crear aun expone @TipoMoneda.', 16, 1);

    SELECT
        N'FASE_4A_VALIDADA' AS Resultado,
        SYSUTCDATETIME() AS FechaValidacionUtc;
END TRY
BEGIN CATCH
    DECLARE @ErrorMessage NVARCHAR(4000);
    DECLARE @ErrorSeverity INT;
    DECLARE @ErrorState INT;

    SELECT
        @ErrorMessage = ERROR_MESSAGE(),
        @ErrorSeverity = ERROR_SEVERITY(),
        @ErrorState = ERROR_STATE();

    RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
END CATCH;
