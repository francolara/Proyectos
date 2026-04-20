USE [DbSportCenter]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- Firma: Codex - 20/04/2026 | Activa o desactiva anuncios popup sin perder configuracion ni imagen asociada.
CREATE OR ALTER PROCEDURE [dbo].[Sp_PopupPromociones_CambiarEstado]
    @IdPopupPromocion INT,
    @Activo BIT,
    @Usuario NVARCHAR(120) = NULL
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        UPDATE dbo.PopupPromocion
        SET
            Activo = @Activo,
            FechaModificacion = SYSDATETIME()
        WHERE IdPopupPromocion = @IdPopupPromocion;

        IF @@ROWCOUNT = 0
            RAISERROR(N'No se encontro el anuncio a actualizar.', 16, 1);
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
