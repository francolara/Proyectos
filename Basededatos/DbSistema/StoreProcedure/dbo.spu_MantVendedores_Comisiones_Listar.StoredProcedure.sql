-- =============================================
-- Author:      FRANCO LARA / Codex
-- Create date: 16/10/2026
-- Description: Lista la configuración de comisiones de un vendedor.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.spu_MantVendedores_Comisiones_Listar
    @IdEmpresa CHAR(2),
    @IdVendedor CHAR(8)
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        SELECT
            C.Item,
            C.PeriodoIni,
            C.Monto1,
            C.Monto2,
            C.Porcentaje1,
            C.Porcentaje2,
            C.Porcentaje3,
            C.IndJefe
        FROM dbo.ComisionxVendedor AS C
        WHERE C.IdEmpresa = @IdEmpresa
          AND C.IdVendedor = @IdVendedor
        ORDER BY C.Item;
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
    END CATCH
END
