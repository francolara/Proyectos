USE [DbSportCenter]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   27/04/2026
-- Description:   Lista referenciales externos del Home para superadmin, con filtros por ubigeo/nombre y paginacion.
-- Firma: Codex - 27/04/2026
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   29/04/2026
-- Description:   Incluye telefono, CodigoUbigeo y TipoDeporteSuperId en listado admin para gestion de activacion/inactivacion/edicion con paginacion.
-- Firma: Codex - 29/04/2026
-- =============================================
CREATE OR ALTER PROCEDURE [dbo].[Sp_Home_ReferencialesExternos_ListarAdmin]
    @CodigoDepartamento CHAR(2) = NULL,
    @CodigoProvincia CHAR(4) = NULL,
    @CodigoUbigeo CHAR(6) = NULL,
    @BuscarNombre NVARCHAR(180) = NULL,
    @Pagina INT = 1,
    @TamanoPagina INT = 20,
    @SoloActivos BIT = 1,
    @TotalRegistros INT OUTPUT
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        SET @CodigoDepartamento = NULLIF(LTRIM(RTRIM(@CodigoDepartamento)), '');
        SET @CodigoProvincia = NULLIF(LTRIM(RTRIM(@CodigoProvincia)), '');
        SET @CodigoUbigeo = NULLIF(LTRIM(RTRIM(@CodigoUbigeo)), '');
        SET @BuscarNombre = NULLIF(LTRIM(RTRIM(@BuscarNombre)), '');
        SET @Pagina = CASE WHEN @Pagina IS NULL OR @Pagina < 1 THEN 1 ELSE @Pagina END;
        SET @TamanoPagina = CASE WHEN @TamanoPagina IS NULL OR @TamanoPagina < 1 THEN 20 ELSE @TamanoPagina END;

        IF OBJECT_ID('tempdb..#BaseReferenciales') IS NOT NULL
            DROP TABLE #BaseReferenciales;

        CREATE TABLE #BaseReferenciales
        (
            Id INT NOT NULL,
            NombreComplejo NVARCHAR(180) NOT NULL,
            NombreEspacio NVARCHAR(150) NULL,
            CodigoUbigeo CHAR(6) NOT NULL,
            TipoDeporteSuperId INT NOT NULL,
            TipoDeporte NVARCHAR(120) NOT NULL,
            Departamento NVARCHAR(120) NOT NULL,
            Provincia NVARCHAR(120) NOT NULL,
            Distrito NVARCHAR(120) NOT NULL,
            Direccion NVARCHAR(250) NULL,
            TelefonoContacto NVARCHAR(40) NULL,
            GoogleMapsUrl NVARCHAR(500) NULL,
            Activo BIT NOT NULL,
            FechaActualizacion DATETIME2(7) NULL,
            UsuarioActualizacion NVARCHAR(200) NULL
        );

        INSERT INTO #BaseReferenciales
        (
            Id,
            NombreComplejo,
            NombreEspacio,
            CodigoUbigeo,
            TipoDeporteSuperId,
            TipoDeporte,
            Departamento,
            Provincia,
            Distrito,
            Direccion,
            TelefonoContacto,
            GoogleMapsUrl,
            Activo,
            FechaActualizacion,
            UsuarioActualizacion
        )
        SELECT
            he.Id,
            he.NombreComplejo,
            he.NombreEspacio,
            he.CodigoUbigeo,
            he.TipoDeporteSuperId,
            tsm.Nombre AS TipoDeporte,
            udp.Nombre AS Departamento,
            upp.Nombre AS Provincia,
            ud.Nombre AS Distrito,
            he.Direccion,
            he.TelefonoContacto,
            he.GoogleMapsUrl,
            he.Activo,
            COALESCE(he.FechaActualizacion, he.FechaCreacion) AS FechaActualizacion,
            COALESCE(he.UsuarioActualizacion, he.UsuarioCreacion) AS UsuarioActualizacion
        FROM dbo.HomeEspaciosReferencialesExternos he
        INNER JOIN dbo.UbigeoDistritos ud ON ud.CodigoUbigeo = he.CodigoUbigeo
        INNER JOIN dbo.UbigeoProvincias upp ON upp.CodigoProvincia = ud.CodigoProvincia
        INNER JOIN dbo.UbigeoDepartamentos udp ON udp.CodigoDepartamento = ud.CodigoDepartamento
        INNER JOIN dbo.TiposDeporteSuperMaestro tsm ON tsm.Id = he.TipoDeporteSuperId
        WHERE (@SoloActivos IS NULL OR he.Activo = @SoloActivos)
          AND (@CodigoDepartamento IS NULL OR ud.CodigoDepartamento = @CodigoDepartamento)
          AND (@CodigoProvincia IS NULL OR ud.CodigoProvincia = @CodigoProvincia)
          AND (@CodigoUbigeo IS NULL OR he.CodigoUbigeo = @CodigoUbigeo)
          AND (
                @BuscarNombre IS NULL
                OR he.NombreComplejo LIKE N'%' + @BuscarNombre + N'%'
                OR ISNULL(he.NombreEspacio, N'') LIKE N'%' + @BuscarNombre + N'%'
                OR ISNULL(he.Direccion, N'') LIKE N'%' + @BuscarNombre + N'%'
              );

        SELECT @TotalRegistros = COUNT(1)
        FROM #BaseReferenciales;

        SELECT
            b.Id,
            b.NombreComplejo,
            b.NombreEspacio,
            b.CodigoUbigeo,
            b.TipoDeporteSuperId,
            b.TipoDeporte,
            b.Departamento,
            b.Provincia,
            b.Distrito,
            b.Direccion,
            b.TelefonoContacto,
            b.GoogleMapsUrl,
            b.Activo,
            b.FechaActualizacion,
            b.UsuarioActualizacion
        FROM #BaseReferenciales b
        ORDER BY
            b.Activo DESC,
            b.NombreComplejo ASC,
            b.Id DESC
        OFFSET (@Pagina - 1) * @TamanoPagina ROWS
        FETCH NEXT @TamanoPagina ROWS ONLY;
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
