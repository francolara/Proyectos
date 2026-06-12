
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- Firma: Codex - 14/04/2026 | Lista banners configurados para Home publico (TipoBanner=1), incluye imagen mobile.
-- Firma: Codex - 11/06/2026 | Filtra solo banners Home publicos activos y vigentes segun FechaInicio/FechaFin.
CREATE OR ALTER PROCEDURE [dbo].[Sp_Home_ListarBannersPublicos]
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        DECLARE @FechaActual DATE = CAST(GETDATE() AS DATE);

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
          AND b.Activo = 1
          AND (b.FechaInicio IS NULL OR b.FechaInicio <= @FechaActual)
          AND (b.FechaFin IS NULL OR b.FechaFin >= @FechaActual)
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
