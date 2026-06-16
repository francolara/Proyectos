-- =============================================
-- Author:        FRANCO LARA
-- Create date:   15/06/2026
-- Description:   Lista los origenes contables activos de una empresa.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_CON_ListarOrigenesActivos
    @IdEmpresa INT
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        SELECT
            o.IdOrigen,
            o.CodigoOrigen,
            o.NombreOrigen,
            o.ModuloOrigen,
            o.PermiteRegistroManual
        FROM dbo.CON_Origen AS o
        WHERE o.IdEmpresa = @IdEmpresa
          AND o.Estado = 1
        ORDER BY o.CodigoOrigen ASC;

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
