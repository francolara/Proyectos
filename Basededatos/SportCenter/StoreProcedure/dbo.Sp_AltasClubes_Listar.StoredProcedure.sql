
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- SOURCE: 21_Altas_Clubes.sql (linea 90)
-- Firma: Codex - 20/04/2026 | Agrega paginacion backend y KPIs de estado por outputs para panel superadmin.
-- Firma: FRANCO LARA - 21/07/2026 | Expone el plan comercial publico de cada solicitud.
CREATE OR ALTER PROCEDURE [dbo].[Sp_AltasClubes_Listar]
    @Estado INT = NULL,
    @Pagina INT = 1,
    @TamanoPagina INT = 20,
    @TotalRegistros INT OUTPUT,
    @TotalPendientes INT OUTPUT,
    @TotalAprobados INT OUTPUT,
    @TotalRechazados INT OUTPUT
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SET @Pagina = CASE WHEN ISNULL(@Pagina, 0) < 1 THEN 1 ELSE @Pagina END;
        SET @TamanoPagina = CASE WHEN ISNULL(@TamanoPagina, 0) < 1 THEN 20 ELSE @TamanoPagina END;

        SELECT @TotalPendientes = COUNT(1) FROM dbo.SolicitudesAltaClub WHERE Estado = 1;
        SELECT @TotalAprobados = COUNT(1) FROM dbo.SolicitudesAltaClub WHERE Estado = 2;
        SELECT @TotalRechazados = COUNT(1) FROM dbo.SolicitudesAltaClub WHERE Estado = 3;

        SELECT
            @TotalRegistros = COUNT(1)
        FROM dbo.SolicitudesAltaClub ac
        WHERE (@Estado IS NULL OR ac.Estado = @Estado);

        SELECT
            ac.Id,
            ac.CodigoSolicitud,
            ac.NombreContacto,
            ac.Telefono,
            ac.Correo,
            ac.RelacionClub,
            ac.NombreClub,
            ac.Pais,
            ac.ProvinciaEstado,
            ac.Ciudad,
            ac.Direccion,
            ac.Estado,
            ac.ComentarioGestion,
            ac.NegocioId,
            ac.SedeId,
            ac.FechaRegistro,
            ac.FechaGestion,
            COALESCE(ac.PlanComercial, N'PRUEBA') AS PlanComercial
        FROM dbo.SolicitudesAltaClub ac
        WHERE (@Estado IS NULL OR ac.Estado = @Estado)
        ORDER BY ac.FechaRegistro DESC
        OFFSET ((@Pagina - 1) * @TamanoPagina) ROWS
        FETCH NEXT @TamanoPagina ROWS ONLY;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
