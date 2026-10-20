-- =============================================
-- Author:      FRANCO LARA / Codex
-- Create date: 15/10/2026
-- Description: Obtiene una forma de pago para su mantenimiento.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.spu_MantFormasPago_Obtener
    @IdEmpresa CHAR(2),
    @IdFormaPago CHAR(8)
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        SELECT
            F.IdFormaPago,
            F.GlsFormaPago,
            F.IdTipoFormaPago,
            F.DiasVcto,
            F.Visible_Ventas,
            F.Cheque
        FROM dbo.FormasPagos AS F
        WHERE F.IdEmpresa = @IdEmpresa
          AND F.IdFormaPago = @IdFormaPago;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000);
        DECLARE @ErrorSeverity INT;
        DECLARE @ErrorState INT;

        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
