USE [DbSportCenter]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- Firma: Codex - 20/04/2026 | Elimina anuncios popup del home publico por Id.
CREATE OR ALTER PROCEDURE [dbo].[Sp_PopupPromociones_Eliminar]
    @IdPopupPromocion INT,
    @Usuario NVARCHAR(120) = NULL
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        DELETE FROM dbo.PopupPromocion
        WHERE IdPopupPromocion = @IdPopupPromocion;

        IF @@ROWCOUNT = 0
            RAISERROR(N'No se encontro el anuncio a eliminar.', 16, 1);
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
