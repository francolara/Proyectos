-- =============================================
-- Author:        FRANCO LARA
-- Create date:   18/06/2026
-- Description:   Copia origenes maestros internos hacia una empresa cuando aun no tiene origenes.
-- =============================================
-- =============================================
-- Author:        FRANCO LARA / Codex
-- Create date:   25/08/2026
-- Description:   Carga origenes y configuracion contable maestra en una sola transaccion, resolviendo CodigoOrigen dentro de la empresa.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_CON_CargarOrigenesDefaultEmpresa
    @IdEmpresa INT,
    @UsuarioRegistro NVARCHAR(450) = NULL
AS
BEGIN

    SET NOCOUNT ON;
    SET XACT_ABORT ON;

    BEGIN TRY

        DECLARE @CodigoOrigenFaltante VARCHAR(10)

        IF NOT EXISTS
        (
            SELECT 1
            FROM dbo.SEG_Empresa AS empresa
            WHERE empresa.IdEmpresa = @IdEmpresa
        )
        BEGIN
            RAISERROR(N'La empresa indicada no existe.', 16, 1);
        END;

        IF EXISTS
        (
            SELECT 1
            FROM dbo.CON_Origen AS o
            WHERE o.IdEmpresa = @IdEmpresa
        )
        BEGIN
            RAISERROR(N'La empresa ya tiene origenes registrados.', 16, 1);
        END;

        BEGIN TRANSACTION;

        INSERT INTO dbo.CON_Origen
        (
            IdEmpresa,
            CodigoOrigen,
            NombreOrigen,
            ModuloOrigen,
            PermiteRegistroManual,
            Estado,
            UsuarioRegistro
        )
        SELECT
            @IdEmpresa,
            om.CodigoOrigen,
            om.NombreOrigen,
            om.ModuloOrigen,
            om.PermiteRegistroManual,
            om.Estado,
            @UsuarioRegistro
        FROM dbo.CON_OrigenMaestro AS om
        WHERE om.Estado = 1
        ORDER BY om.Orden, om.CodigoOrigen;

        SELECT TOP (1)
            @CodigoOrigenFaltante = maestro.CodigoOrigen
        FROM dbo.CON_ConfiguracionContabilizacionMaestro AS maestro
        LEFT JOIN dbo.CON_Origen AS origen
            ON origen.IdEmpresa = @IdEmpresa
           AND origen.CodigoOrigen = maestro.CodigoOrigen
           AND origen.Estado = 1
        WHERE maestro.Activo = 1
          AND origen.IdOrigen IS NULL
        ORDER BY maestro.Orden, maestro.ModuloOperacion;

        IF @CodigoOrigenFaltante IS NOT NULL
        BEGIN
            RAISERROR(N'El origen maestro %s requerido por la configuracion contable no existe o esta inactivo.', 16, 1, @CodigoOrigenFaltante);
        END;

        INSERT INTO dbo.CON_ConfiguracionContabilizacion
        (
            IdEmpresa,
            ModuloOperacion,
            EscenarioOperacion,
            IdOrigen,
            Descripcion,
            GeneraAsientoAutomatico,
            UsaTipoCambio,
            Activo,
            UsuarioRegistro
        )
        SELECT
            @IdEmpresa,
            maestro.ModuloOperacion,
            maestro.EscenarioOperacion,
            origen.IdOrigen,
            maestro.Descripcion,
            maestro.GeneraAsientoAutomatico,
            maestro.UsaTipoCambio,
            maestro.Activo,
            @UsuarioRegistro
        FROM dbo.CON_ConfiguracionContabilizacionMaestro AS maestro
        INNER JOIN dbo.CON_Origen AS origen
            ON origen.IdEmpresa = @IdEmpresa
           AND origen.CodigoOrigen = maestro.CodigoOrigen
           AND origen.Estado = 1
        WHERE maestro.Activo = 1
          AND NOT EXISTS
          (
              SELECT 1
              FROM dbo.CON_ConfiguracionContabilizacion AS configuracion
              WHERE configuracion.IdEmpresa = @IdEmpresa
                AND configuracion.ModuloOperacion = maestro.ModuloOperacion
                AND configuracion.EscenarioOperacion = maestro.EscenarioOperacion
          );

        COMMIT TRANSACTION;

    END TRY

    BEGIN CATCH

        IF XACT_STATE() <> 0
        BEGIN
            ROLLBACK TRANSACTION;
        END;

        DECLARE @ErrorMessage NVARCHAR(4000)
        DECLARE @ErrorSeverity INT
        DECLARE @ErrorState INT

        SELECT
            @ErrorMessage = ERROR_MESSAGE(),
            @ErrorSeverity = ERROR_SEVERITY(),
            @ErrorState = ERROR_STATE()

        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState)

    END CATCH

END
