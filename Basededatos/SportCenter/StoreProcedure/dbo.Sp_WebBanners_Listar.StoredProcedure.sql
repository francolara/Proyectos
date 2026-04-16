USE [DbSportCenter]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- Firma: Codex - 14/04/2026 | Lista mantenimiento de banners web con tipo (Home/Login/Registro) y filtro opcional por estado.
CREATE OR ALTER PROCEDURE [dbo].[Sp_WebBanners_Listar]
    @SoloActivos BIT = NULL
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
            b.TipoBanner,
            CASE b.TipoBanner
                WHEN 2 THEN N'Login fijo'
                WHEN 3 THEN N'Registro fijo'
                ELSE N'Home publico'
            END AS TipoBannerNombre,
            b.Orden,
            b.Activo,
            b.FechaInicio,
            b.FechaFin
        FROM dbo.WebBanners b
        WHERE (@SoloActivos IS NULL OR b.Activo = @SoloActivos)
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
