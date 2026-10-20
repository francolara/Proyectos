
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   02/10/2026
-- Description:   Recalcula las cantidades y estados de importacion de un pedido.
-- =============================================
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   19/10/2026
-- Description:   Adapta el procedimiento anterior al recalculo universal por item origen.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.spu_RecalculaImportacionPedido
    @IdEmpresa VARCHAR(10),
    @IdSucursal VARCHAR(10),
    @IdSeriePedido VARCHAR(10),
    @IdDocVentasPedido VARCHAR(20)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        EXECUTE dbo.spu_RecalculaImportacionDocumento
            @IdEmpresa,
            @IdSucursal,
            '40',
            @IdSeriePedido,
            @IdDocVentasPedido;
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
GO
