-- =============================================
-- Author:      FRANCO LARA / Codex
-- Create date: 15/10/2026
-- Description: Lista tipos de cambio para su mantenimiento.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.spu_MantTiposCambio_Listar
    @Busqueda VARCHAR(255) = ''
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        SET @Busqueda = LTRIM(RTRIM(ISNULL(@Busqueda, '')));

        SELECT
            CAST(TC.Fecha AS DATE) AS Fecha,
            TC.TCCompra,
            TC.TCVenta,
            TC.TCFacturacion
        FROM dbo.TiposDeCambio AS TC
        WHERE @Busqueda = ''
           OR CONVERT(CHAR(10), TC.Fecha, 103) LIKE '%' + @Busqueda + '%'
        ORDER BY TC.Fecha DESC;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;

        SELECT
            @ErrorMessage = ERROR_MESSAGE(),
            @ErrorSeverity = ERROR_SEVERITY(),
            @ErrorState = ERROR_STATE();

        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
