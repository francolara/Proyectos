-- =============================================
-- Author:        FRANCO LARA
-- Create date:   30/06/2026
-- Description:   Obtiene el estado abierto o cerrado de un periodo contable por empresa.
-- =============================================
-- Firma: FRANCO LARA - 30/06/2026 | Consulta el estado del periodo contable para habilitar o bloquear la operativa de registros.

CREATE OR ALTER PROCEDURE dbo.usp_CON_ObtenerPeriodoContableEstado
    @IdEmpresa INT,
    @Periodo CHAR(6)
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        SET @Periodo = LTRIM(RTRIM(@Periodo));

        IF @Periodo NOT LIKE '[1-2][0-9][0-9][0-9][0-1][0-9]'
        BEGIN
            RAISERROR(N'El periodo debe tener formato YYYYMM.', 16, 1);
        END;

        SELECT
            pce.IdPeriodoContableEstado,
            pce.IdEmpresa,
            pce.Periodo,
            pce.Cerrado,
            pce.FechaRegistro,
            pce.UsuarioRegistro,
            pce.FechaCierre,
            pce.UsuarioCierre,
            pce.FechaApertura,
            pce.UsuarioApertura
        FROM dbo.CON_PeriodoContableEstado AS pce
        WHERE pce.IdEmpresa = @IdEmpresa
          AND pce.Periodo = @Periodo;

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
