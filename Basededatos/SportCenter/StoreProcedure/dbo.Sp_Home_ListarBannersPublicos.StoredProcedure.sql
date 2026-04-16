USE [DbSportCenter]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- Firma: Codex - 14/04/2026 | Lista banners configurados para Home publico (TipoBanner=1), incluye imagen mobile.
CREATE OR ALTER PROCEDURE [dbo].[Sp_Home_ListarBannersPublicos]
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        SELECT
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
        WHERE b.TipoBanner = 1
        ORDER BY b.Orden, b.Id;
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
