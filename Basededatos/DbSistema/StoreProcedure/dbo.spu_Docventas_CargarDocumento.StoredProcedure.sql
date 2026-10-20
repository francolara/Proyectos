
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- =============================================
-- Author:      FRANCO LARA / Codex
-- Create date: 13/10/2026
-- Description: Devuelve cabecera, detalle y referencias de un documento de ventas en una sola llamada.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.spu_Docventas_CargarDocumento
    @IdEmpresa CHAR(2),
    @IdSucursal VARCHAR(8),
    @IdDocumento CHAR(2),
    @IdSerie CHAR(4),
    @IdDocVentas CHAR(8)
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        EXEC dbo.spu_Docventas_Mostrar_Docventas_cab
            @idEmpresa = @IdEmpresa,
            @idSucursal = @IdSucursal,
            @idDocumento = @IdDocumento,
            @idSerie = @IdSerie,
            @iddocventas = @IdDocVentas;

        EXEC dbo.spu_Docventas_Mostrar_Docventas_Det
            @idEmpresa = @IdEmpresa,
            @idSucursal = @IdSucursal,
            @idDocumento = @IdDocumento,
            @idSerie = @IdSerie,
            @iddocventas = @IdDocVentas;

        EXEC dbo.spu_Docventas_Mostrar_Docventas_Ref
            @idEmpresa = @IdEmpresa,
            @idSucursal = @IdSucursal,
            @idDocumento = @IdDocumento,
            @idSerie = @IdSerie,
            @iddocventas = @IdDocVentas;
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
