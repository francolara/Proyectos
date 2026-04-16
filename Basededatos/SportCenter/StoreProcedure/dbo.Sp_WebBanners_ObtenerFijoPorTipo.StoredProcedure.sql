USE [DbSportCenter]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- Firma: Codex - 14/04/2026 | Obtiene banner fijo por tipo (Login/Registro), priorizando activo y menor orden.
CREATE OR ALTER PROCEDURE [dbo].[Sp_WebBanners_ObtenerFijoPorTipo]
    @TipoBanner TINYINT
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        SET @TipoBanner = CASE WHEN @TipoBanner IN (2, 3) THEN @TipoBanner ELSE 2 END;

        SELECT TOP (1)
            b.Id,
            b.Titulo,
            b.Subtitulo,
            b.Descripcion,
            b.BotonTexto,
            b.BotonUrl,
            b.ImagenUrl,
            b.ImagenUrlMobile,
            b.Orden
        FROM dbo.WebBanners b
        WHERE b.TipoBanner = @TipoBanner
        ORDER BY
            CASE WHEN b.Activo = 1 THEN 0 ELSE 1 END,
            b.Orden,
            b.Id;
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