USE [DbSportCenter]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- Firma: Codex - 14/04/2026 | Inserta/actualiza banners web (Home/Login/Registro) y retorna Id generado/actualizado.
-- Firma: Codex - 02/05/2026 | Corrige validacion por tipo: Home exige imagen horizontal; Login/Registro exigen imagen vertical y permiten fallback de ImagenUrl desde ImagenUrlMobile.
CREATE OR ALTER PROCEDURE [dbo].[Sp_WebBanners_Guardar]
    @Id INT = NULL,
    @Titulo NVARCHAR(120),
    @Subtitulo NVARCHAR(220) = NULL,
    @Descripcion NVARCHAR(400) = NULL,
    @BotonTexto NVARCHAR(40) = NULL,
    @BotonUrl NVARCHAR(300) = NULL,
    @ImagenUrl NVARCHAR(500),
    @ImagenUrlMobile NVARCHAR(500) = NULL,
    @TipoBanner TINYINT = 1,
    @Orden INT = 1,
    @Activo BIT = 1,
    @FechaInicio DATE = NULL,
    @FechaFin DATE = NULL,
    @Usuario NVARCHAR(120) = NULL
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        SET @Titulo = LTRIM(RTRIM(COALESCE(@Titulo, N'')));
        SET @Subtitulo = NULLIF(LTRIM(RTRIM(COALESCE(@Subtitulo, N''))), N'');
        SET @Descripcion = NULLIF(LTRIM(RTRIM(COALESCE(@Descripcion, N''))), N'');
        SET @BotonTexto = NULLIF(LTRIM(RTRIM(COALESCE(@BotonTexto, N''))), N'');
        SET @BotonUrl = NULLIF(LTRIM(RTRIM(COALESCE(@BotonUrl, N''))), N'');
        SET @ImagenUrl = LTRIM(RTRIM(COALESCE(@ImagenUrl, N'')));
        SET @ImagenUrlMobile = NULLIF(LTRIM(RTRIM(COALESCE(@ImagenUrlMobile, N''))), N'');
        SET @TipoBanner = CASE WHEN @TipoBanner IN (1, 2, 3) THEN @TipoBanner ELSE 1 END;
        SET @Orden = CASE WHEN @Orden IS NULL OR @Orden < 1 THEN 1 ELSE @Orden END;

        IF @Titulo = N''
            RAISERROR(N'El titulo del banner es obligatorio.', 16, 1);

        IF @TipoBanner = 1 AND @ImagenUrl = N''
            RAISERROR(N'La imagen del banner es obligatoria para Home.', 16, 1);

        IF @TipoBanner IN (2, 3) AND ISNULL(@ImagenUrlMobile, N'') = N''
            RAISERROR(N'La imagen vertical del banner es obligatoria para Login/Registro.', 16, 1);

        IF @TipoBanner IN (2, 3) AND @ImagenUrl = N'' AND ISNULL(@ImagenUrlMobile, N'') <> N''
            SET @ImagenUrl = @ImagenUrlMobile;

        IF @FechaInicio IS NOT NULL AND @FechaFin IS NOT NULL AND @FechaInicio > @FechaFin
            RAISERROR(N'La fecha inicio no puede ser mayor a la fecha fin.', 16, 1);

        IF @Id IS NULL OR @Id <= 0
        BEGIN
            INSERT INTO dbo.WebBanners
            (
                Titulo, Subtitulo, Descripcion, BotonTexto, BotonUrl, ImagenUrl, ImagenUrlMobile, TipoBanner,
                Orden, Activo, FechaInicio, FechaFin, UsuarioRegistro
            )
            VALUES
            (
                @Titulo, @Subtitulo, @Descripcion, @BotonTexto, @BotonUrl, @ImagenUrl, @ImagenUrlMobile, @TipoBanner,
                @Orden, @Activo, @FechaInicio, @FechaFin, @Usuario
            );

            SELECT CAST(SCOPE_IDENTITY() AS INT) AS Id;
            RETURN;
        END

        UPDATE dbo.WebBanners
        SET
            Titulo = @Titulo,
            Subtitulo = @Subtitulo,
            Descripcion = @Descripcion,
            BotonTexto = @BotonTexto,
            BotonUrl = @BotonUrl,
            ImagenUrl = @ImagenUrl,
            ImagenUrlMobile = @ImagenUrlMobile,
            TipoBanner = @TipoBanner,
            Orden = @Orden,
            Activo = @Activo,
            FechaInicio = @FechaInicio,
            FechaFin = @FechaFin,
            FechaActualizacion = SYSDATETIME(),
            UsuarioActualizacion = @Usuario
        WHERE Id = @Id;

        IF @@ROWCOUNT = 0
            RAISERROR(N'No se encontro el banner a actualizar.', 16, 1);

        SELECT @Id AS Id;
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
