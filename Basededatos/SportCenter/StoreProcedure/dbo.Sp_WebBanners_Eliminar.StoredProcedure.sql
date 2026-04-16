USE [DbSportCenter]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- Firma: Codex - 14/04/2026 | Elimina banner web por Id para mantenimiento del carrusel.
CREATE OR ALTER PROCEDURE [dbo].[Sp_WebBanners_Eliminar]
    @Id INT,
    @Usuario NVARCHAR(120) = NULL
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        DELETE FROM dbo.WebBanners
        WHERE Id = @Id;

        IF @@ROWCOUNT = 0
            RAISERROR(N'No se encontro el banner a eliminar.', 16, 1);
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT
            @ErrorMessage = ERROR_MESSAGE(),
            @ErrorSeverity = ERROR_SEVERITY(),
            @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO
