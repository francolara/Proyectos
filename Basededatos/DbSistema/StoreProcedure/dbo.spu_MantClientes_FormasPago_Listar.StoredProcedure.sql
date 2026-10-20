-- =============================================
-- Author:      FRANCO LARA / Codex
-- Create date: 15/10/2026
-- Description: Lista formas de pago del cliente sin consultas por fila.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.spu_MantClientes_FormasPago_Listar
    @IdEmpresa CHAR(2),
    @IdCliente CHAR(8)
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        SELECT
            P.GlsPersona,
            C.IdFormaPago,
            F.GlsFormaPago,
            C.IndEstado,
            C.FecRegistro,
            C.IdUsuario,
            CAST(0 AS BIGINT) AS IndPrincipal
        FROM dbo.ClientesFormaPagos AS C
        INNER JOIN dbo.Personas AS P
            ON C.IdCliente = P.IdPersona
        LEFT JOIN dbo.FormasPagos AS F
            ON C.IdEmpresa = F.IdEmpresa
           AND C.IdFormaPago = F.IdFormaPago
        WHERE C.IdEmpresa = @IdEmpresa
          AND C.IdCliente = @IdCliente
        ORDER BY C.IdFormaPago;
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
