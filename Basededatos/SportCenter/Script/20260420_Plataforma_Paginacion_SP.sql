-- =============================================
-- Author:        FRANCO LARA
-- Create date:   20/04/2026
-- Firma:         Paginacion backend para listados de negocios y altas de clubes del panel superadmin.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.Sp_Plataforma_Negocios_Listar
    @Buscar NVARCHAR(200) = NULL,
    @EstadoContrato NVARCHAR(30) = N'todos',
    @Pagina INT = 1,
    @TamanoPagina INT = 20,
    @TotalRegistros INT OUTPUT
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        DECLARE @EstadoContratoNorm NVARCHAR(30) = LOWER(LTRIM(RTRIM(COALESCE(@EstadoContrato, N'todos'))));
        DECLARE @Hoy DATE = CAST(GETDATE() AS DATE);

        SET @Pagina = CASE WHEN ISNULL(@Pagina, 0) < 1 THEN 1 ELSE @Pagina END;
        SET @TamanoPagina = CASE WHEN ISNULL(@TamanoPagina, 0) < 1 THEN 20 ELSE @TamanoPagina END;
        SET @Buscar = NULLIF(LTRIM(RTRIM(@Buscar)), N'');

        SELECT
            @TotalRegistros = COUNT(1)
        FROM dbo.Negocios n
        LEFT JOIN dbo.NegociosSuscripcion ns ON ns.NegocioId = n.Id
        WHERE (@Buscar IS NULL OR n.NombreComercial LIKE N'%' + @Buscar + N'%')
          AND (
                @EstadoContratoNorm = N'todos'
                OR (@EstadoContratoNorm = N'con-contrato' AND COALESCE(ns.EsPrueba, 0) = 0 AND ns.TipoCobro IS NOT NULL AND ns.FechaFinPlan IS NOT NULL)
                OR (@EstadoContratoNorm = N'sin-contrato' AND NOT (COALESCE(ns.EsPrueba, 0) = 0 AND ns.TipoCobro IS NOT NULL AND ns.FechaFinPlan IS NOT NULL))
                OR (@EstadoContratoNorm = N'prueba-por-vencer' AND COALESCE(ns.EsPrueba, 0) = 1 AND ns.FechaFinPrueba IS NOT NULL AND ns.FechaFinPrueba >= @Hoy AND ns.FechaFinPrueba <= DATEADD(DAY, 7, @Hoy))
              );

        SELECT
            n.Id,
            n.NombreComercial,
            n.Activo,
            CAST(COALESCE(n.SedesPermitidas, 2) AS INT) AS SedesPermitidas,
            CAST(COALESCE(n.EspaciosPermitidos, 6) AS INT) AS EspaciosPermitidos,
            CAST(COALESCE(n.UsuariosPermitidos, 3) AS INT) AS UsuariosPermitidos,
            CAST(COALESCE(ns.EstadoSuscripcion, 0) AS INT) AS EstadoSuscripcion,
            CAST(COALESCE(ns.EsPrueba, 0) AS BIT) AS EsPrueba,
            ns.FechaInicioPrueba,
            ns.FechaFinPrueba,
            ns.TipoCobro,
            ns.FechaInicioPlan,
            ns.FechaFinPlan,
            CAST(COALESCE(ns.DiasGracia, 5) AS INT) AS DiasGracia,
            ns.FechaFinGracia
        FROM dbo.Negocios n
        LEFT JOIN dbo.NegociosSuscripcion ns ON ns.NegocioId = n.Id
        WHERE (@Buscar IS NULL OR n.NombreComercial LIKE N'%' + @Buscar + N'%')
          AND (
                @EstadoContratoNorm = N'todos'
                OR (@EstadoContratoNorm = N'con-contrato' AND COALESCE(ns.EsPrueba, 0) = 0 AND ns.TipoCobro IS NOT NULL AND ns.FechaFinPlan IS NOT NULL)
                OR (@EstadoContratoNorm = N'sin-contrato' AND NOT (COALESCE(ns.EsPrueba, 0) = 0 AND ns.TipoCobro IS NOT NULL AND ns.FechaFinPlan IS NOT NULL))
                OR (@EstadoContratoNorm = N'prueba-por-vencer' AND COALESCE(ns.EsPrueba, 0) = 1 AND ns.FechaFinPrueba IS NOT NULL AND ns.FechaFinPrueba >= @Hoy AND ns.FechaFinPrueba <= DATEADD(DAY, 7, @Hoy))
              )
        ORDER BY n.NombreComercial, n.Id
        OFFSET ((@Pagina - 1) * @TamanoPagina) ROWS
        FETCH NEXT @TamanoPagina ROWS ONLY;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END

CREATE OR ALTER PROCEDURE dbo.Sp_AltasClubes_Listar
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
            ac.FechaGestion
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
