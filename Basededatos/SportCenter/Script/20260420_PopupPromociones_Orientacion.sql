USE [DbSportCenter]
GO

-- Firma: Codex - 20/04/2026 | Agrega orientacion por anuncio (vertical/horizontal) al modulo PopupPromocion y ajusta los SPs para soportar piezas mixtas.
IF COL_LENGTH('dbo.PopupPromocion', 'Orientacion') IS NULL
BEGIN
    ALTER TABLE dbo.PopupPromocion
    ADD Orientacion CHAR(1) NOT NULL CONSTRAINT DF_PopupPromocion_Orientacion DEFAULT ('V');
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.check_constraints WHERE name = N'CK_PopupPromocion_Orientacion' AND parent_object_id = OBJECT_ID(N'dbo.PopupPromocion'))
BEGIN
    ALTER TABLE dbo.PopupPromocion
    ADD CONSTRAINT CK_PopupPromocion_Orientacion CHECK (Orientacion IN ('V', 'H'));
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Home_ListarPopupPromocionesActivas
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        DECLARE @Hoy DATE = CONVERT(DATE, SYSDATETIME());

        SELECT
            p.IdPopupPromocion,
            p.Titulo,
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

CREATE OR ALTER PROCEDURE dbo.Sp_PopupPromociones_ListarAdmin
    @SoloActivos BIT = NULL
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        SELECT
            p.IdPopupPromocion,
            p.Titulo,
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

CREATE OR ALTER PROCEDURE dbo.Sp_PopupPromociones_Guardar
    @IdPopupPromocion INT = NULL,
    @Titulo NVARCHAR(120),
    @Descripcion NVARCHAR(260) = NULL,
    @ImagenUrl NVARCHAR(500),
    @Orientacion CHAR(1) = 'V',
    @TextoBoton NVARCHAR(40) = NULL,
    @UrlBoton NVARCHAR(300) = NULL,
    @UrlImagen NVARCHAR(300) = NULL,
    @Orden INT = 1,
    @Activo BIT = 1,
    @FechaInicio DATE = NULL,
    @FechaFin DATE = NULL,
    @AbrirNuevaPestana BIT = 1,
    @Usuario NVARCHAR(120) = NULL
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        SET @Titulo = LTRIM(RTRIM(COALESCE(@Titulo, N'')));
        SET @Descripcion = NULLIF(LTRIM(RTRIM(COALESCE(@Descripcion, N''))), N'');
        SET @ImagenUrl = LTRIM(RTRIM(COALESCE(@ImagenUrl, N'')));
        SET @Orientacion = UPPER(LTRIM(RTRIM(COALESCE(@Orientacion, 'V'))));
        SET @TextoBoton = NULLIF(LTRIM(RTRIM(COALESCE(@TextoBoton, N''))), N'');
        SET @UrlBoton = NULLIF(LTRIM(RTRIM(COALESCE(@UrlBoton, N''))), N'');
        SET @UrlImagen = NULLIF(LTRIM(RTRIM(COALESCE(@UrlImagen, N''))), N'');
        SET @Orden = CASE WHEN @Orden IS NULL OR @Orden < 1 THEN 1 ELSE @Orden END;

        IF @Titulo = N''
            RAISERROR(N'El titulo del anuncio es obligatorio.', 16, 1);

        IF @ImagenUrl = N''
            RAISERROR(N'La imagen del anuncio es obligatoria.', 16, 1);

        IF @Orientacion NOT IN ('V', 'H')
            RAISERROR(N'La orientacion del anuncio debe ser V o H.', 16, 1);

        IF @FechaInicio IS NOT NULL AND @FechaFin IS NOT NULL AND @FechaInicio > @FechaFin
            RAISERROR(N'La fecha inicio no puede ser mayor a la fecha fin.', 16, 1);

        IF @IdPopupPromocion IS NULL OR @IdPopupPromocion <= 0
        BEGIN
            INSERT INTO dbo.PopupPromocion
            (
                Titulo,
                Descripcion,
                ImagenUrl,
                Orientacion,
                TextoBoton,
                UrlBoton,
                UrlImagen,
                Orden,
                Activo,
                FechaInicio,
                FechaFin,
                AbrirNuevaPestana,
                FechaCreacion
            )
            VALUES
            (
                @Titulo,
                @Descripcion,
                @ImagenUrl,
                @Orientacion,
                @TextoBoton,
                @UrlBoton,
                @UrlImagen,
                @Orden,
                @Activo,
                @FechaInicio,
                @FechaFin,
                @AbrirNuevaPestana,
                SYSDATETIME()
            );

            SELECT CAST(SCOPE_IDENTITY() AS INT) AS IdPopupPromocion;
            RETURN;
        END

        UPDATE dbo.PopupPromocion
        SET
            Titulo = @Titulo,
            Descripcion = @Descripcion,
            ImagenUrl = @ImagenUrl,
            Orientacion = @Orientacion,
            TextoBoton = @TextoBoton,
            UrlBoton = @UrlBoton,
            UrlImagen = @UrlImagen,
            Orden = @Orden,
            Activo = @Activo,
            FechaInicio = @FechaInicio,
            FechaFin = @FechaFin,
            AbrirNuevaPestana = @AbrirNuevaPestana,
            FechaModificacion = SYSDATETIME()
        WHERE IdPopupPromocion = @IdPopupPromocion;

        IF @@ROWCOUNT = 0
            RAISERROR(N'No se encontro el anuncio a actualizar.', 16, 1);

        SELECT @IdPopupPromocion AS IdPopupPromocion;
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
