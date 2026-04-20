USE [DbSportCenter]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- Firma: Codex - 20/04/2026 | Lista anuncios popup para super admin con filtro opcional por estado activo, subtitulo y orientacion por pieza.
CREATE OR ALTER PROCEDURE [dbo].[Sp_PopupPromociones_ListarAdmin]
    @SoloActivos BIT = NULL
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
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
            p.Activo,
            p.FechaInicio,
            p.FechaFin,
            p.AbrirNuevaPestana,
            p.FechaCreacion,
            p.FechaModificacion,
            p.Orientacion
        FROM dbo.PopupPromocion p
        WHERE @SoloActivos IS NULL
           OR p.Activo = @SoloActivos
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
