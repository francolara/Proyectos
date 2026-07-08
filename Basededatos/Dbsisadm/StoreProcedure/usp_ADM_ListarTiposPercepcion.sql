-- =============================================
-- Author:        FRANCO LARA
-- Create date:   29/06/2026
-- Description:   Lista los tipos de percepcion activos para compras.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_ADM_ListarTiposPercepcion
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        SELECT
            tp.IdTipoPercepcion,
            tp.Codigo,
            tp.Descripcion,
            tp.Porcentaje,
            tp.Estado
        FROM dbo.ADM_TipoPercepcion AS tp
        WHERE tp.Estado = 1
        ORDER BY
            tp.Codigo ASC;

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
