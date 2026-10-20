-- =============================================
-- Author:      FRANCO LARA / Codex
-- Create date: 15/10/2026
-- Description: Lista formas de pago para su mantenimiento.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.spu_MantFormasPago_Listar
    @IdEmpresa CHAR(2),
    @Busqueda VARCHAR(255) = ''
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        SET @Busqueda = LTRIM(RTRIM(ISNULL(@Busqueda, '')));

        SELECT
            F.IdFormaPago,
            F.GlsFormaPago,
            T.GlsTipoFormaPago,
            F.DiasVcto
        FROM dbo.FormasPagos AS F
        INNER JOIN dbo.TipoFormasPago AS T
            ON F.IdTipoFormaPago = T.IdTipoFormaPago
        WHERE F.IdEmpresa = @IdEmpresa
          AND (@Busqueda = '' OR F.GlsFormaPago LIKE '%' + @Busqueda + '%')
        ORDER BY F.IdFormaPago;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000);
        DECLARE @ErrorSeverity INT;
        DECLARE @ErrorState INT;

        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
