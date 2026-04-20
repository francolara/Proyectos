USE [DbSportCenter]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- Firma: Codex - 20/04/2026 | Lista anuncios popup vigentes para home publico, con subtitulo y orientacion por pieza, ordenados por prioridad y fecha de creacion.
CREATE OR ALTER PROCEDURE [dbo].[Sp_Home_ListarPopupPromocionesActivas]
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        DECLARE @Hoy DATE = CONVERT(DATE, SYSDATETIME());

        SELECT
            p.IdPopupPromocion,
            p.Titulo,
            p.Subtitulo,
            p.Descripcion,
            p.ImagenUrl,
            p.TextoBoton,
            p.UrlBoton,
            p.UrlImagen,
            p.Orden,
            p.AbrirNuevaPestana,
            p.Orientacion
        FROM dbo.PopupPromocion p
        WHERE p.Activo = 1
          AND (p.FechaInicio IS NULL OR p.FechaInicio <= @Hoy)
          AND (p.FechaFin IS NULL OR p.FechaFin >= @Hoy)
        ORDER BY p.Orden ASC, p.FechaCreacion DESC;
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
