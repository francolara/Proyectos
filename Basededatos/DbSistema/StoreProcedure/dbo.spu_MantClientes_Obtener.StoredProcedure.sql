-- =============================================
-- Author:      FRANCO LARA / Codex
-- Create date: 15/10/2026
-- Description: Obtiene la cabecera de un cliente para su mantenimiento.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.spu_MantClientes_Obtener
    @IdEmpresa CHAR(2),
    @IdCliente CHAR(8)
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        SELECT
            C.IdProvCliente,
            C.IdCliente,
            C.IdVendedorCampo,
            C.IdEmpTrans,
            C.IdLista,
            C.IdFormaPago,
            C.Val_LineaCredito,
            C.IndAgenteRetencion,
            C.IdMonedaLineaCredito,
            C.IdClienteInterno,
            C.Val_Dscto,
            C.IdCanal,
            C.IndVentasTerceros,
            P.Telefonos,
            P.Mail,
            P.GlsContacto,
            C.GlsObservacion,
            C.IndComision,
            C.IndAnula,
            C.IndEspecial
        FROM dbo.Clientes AS C
        INNER JOIN dbo.Personas AS P
            ON C.IdCliente = P.IdPersona
        WHERE C.IdEmpresa = @IdEmpresa
          AND C.IdCliente = @IdCliente;
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
