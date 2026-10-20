-- =============================================
-- Author:      FRANCO LARA / Codex
-- Create date: 15/10/2026
-- Description: Obtiene la linea de credito y moneda efectiva del cliente.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.spu_MantClientes_LineaCredito_Obtener
    @IdEmpresa CHAR(2),
    @IdCliente CHAR(8)
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        SELECT
            ISNULL(L.IdMoneda, ISNULL(PA.ValParametro, '')) AS IdMoneda,
            ISNULL(L.Linea_Actual, 0) AS Linea_Actual,
            ISNULL(L.Linea_Usada, 0) AS Linea_Usada,
            ISNULL(L.Saldo, 0) AS Saldo,
            ISNULL(L.IndSuspension, '0') AS IndSuspension,
            ISNULL(D.GlsDato, '') AS GlsDato,
            CASE WHEN L.IdCliente IS NULL THEN 0 ELSE 1 END AS ExisteControl
        FROM (SELECT @IdEmpresa AS IdEmpresa, @IdCliente AS IdCliente) AS I
        LEFT JOIN dbo.ControlLineaCredito AS L
            ON I.IdEmpresa = L.IdEmpresa
           AND I.IdCliente = L.IdCliente
        LEFT JOIN dbo.Datos AS D
            ON L.Estado = D.IdDato
        LEFT JOIN dbo.Parametros AS PA
            ON PA.IdEmpresa = I.IdEmpresa
           AND PA.GlsParametro = 'MONEDA_LINEA_CREDITO';
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
