-- =============================================
-- Author:      FRANCO LARA / Codex
-- Create date: 15/10/2026
-- Description: Lista las tiendas asociadas a un cliente.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.spu_MantClientes_Tiendas_Listar
    @IdEmpresa CHAR(2),
    @IdCliente CHAR(8)
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        SELECT
            T.IdPais,
            T.IdDistrito,
            U.IdDpto,
            U.IdProv,
            T.IdTdaCli,
            T.Item,
            T.GlsNombre,
            T.GlsDireccion,
            T.GlsTelefonos,
            T.GlsContacto,
            T.IdVendedor,
            PV.GlsPersona AS GlsVendedor
        FROM dbo.TiendasCliente AS T
        LEFT JOIN dbo.Ubigeo AS U
            ON T.IdDistrito = U.IdDistrito
           AND T.IdPais = U.IdPais
        LEFT JOIN dbo.Personas AS PV
            ON T.IdVendedor = PV.IdPersona
        WHERE T.IdEmpresa = @IdEmpresa
          AND T.IdPersona = @IdCliente
        ORDER BY T.Item, T.IdTdaCli;
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
