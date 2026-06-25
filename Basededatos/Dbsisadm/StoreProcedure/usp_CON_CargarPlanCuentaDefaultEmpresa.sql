-- =============================================
-- Author:        FRANCO LARA
-- Create date:   18/06/2026
-- Description:   Copia plan de cuentas maestro interno hacia una empresa con ColBalance, moneda y tipo de cambio.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_CON_CargarPlanCuentaDefaultEmpresa
    @IdEmpresa INT,
    @UsuarioRegistro NVARCHAR(450) = NULL
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        IF EXISTS
        (
            SELECT 1
            FROM dbo.CON_PlanCuenta AS pc
            WHERE pc.IdEmpresa = @IdEmpresa
        )
        BEGIN
            RAISERROR(N'La empresa ya tiene plan de cuentas registrado.', 16, 1);
        END;

        INSERT INTO dbo.CON_PlanCuenta
        (
            IdEmpresa,
            IdPlanCuentaPadre,
            CodigoCuenta,
            NombreCuenta,
            NivelCuenta,
            ColBalance,
            IdMoneda,
            TipoCambio,
            AceptaMovimiento,
            RequiereCentroCosto,
            Estado,
            UsuarioRegistro
        )
        SELECT
            @IdEmpresa,
            NULL,
            pcm.CodigoCuenta,
            pcm.NombreCuenta,
            pcm.NivelCuenta,
            pcm.ColBalance,
            pcm.IdMoneda,
            pcm.TipoCambio,
            pcm.AceptaMovimiento,
            pcm.RequiereCentroCosto,
            pcm.Estado,
            @UsuarioRegistro
        FROM dbo.CON_PlanCuentaMaestro AS pcm
        WHERE pcm.Estado = 1
        ORDER BY pcm.NivelCuenta, pcm.Orden, pcm.CodigoCuenta;

        UPDATE hijo
        SET IdPlanCuentaPadre = padre.IdPlanCuenta
        FROM dbo.CON_PlanCuenta AS hijo
        INNER JOIN dbo.CON_PlanCuentaMaestro AS maestroHijo
            ON maestroHijo.CodigoCuenta = hijo.CodigoCuenta
        INNER JOIN dbo.CON_PlanCuenta AS padre
            ON padre.IdEmpresa = hijo.IdEmpresa
           AND padre.CodigoCuenta = maestroHijo.CodigoCuentaPadre
        WHERE hijo.IdEmpresa = @IdEmpresa;

    END TRY

    BEGIN CATCH

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
