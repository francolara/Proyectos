
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- Firma: Codex - 10/06/2026 | Registra boletines deportivos y limita a una carga semanal para usuarios publicos no administrativos.
CREATE OR ALTER PROCEDURE dbo.Sp_BoletinesDeportivos_Guardar
    @IdBoletin INT = NULL,
    @UsuarioId NVARCHAR(450),
    @Titulo NVARCHAR(160) = NULL,
    @Descripcion NVARCHAR(500) = NULL,
    @ImagenUrl NVARCHAR(500),
    @FechaEvento DATE,
    @CodigoUbigeo CHAR(6),
    @TipoRegistro CHAR(1) = 'U',
    @Activo BIT = 1,
    @EsAdministradorCarga BIT = 0,
    @Usuario NVARCHAR(120)
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        DECLARE @UsuarioIdNorm NVARCHAR(450) = NULLIF(LTRIM(RTRIM(@UsuarioId)), N'');
        DECLARE @TituloNorm NVARCHAR(160) = NULLIF(LTRIM(RTRIM(@Titulo)), N'');
        DECLARE @DescripcionNorm NVARCHAR(500) = NULLIF(LTRIM(RTRIM(@Descripcion)), N'');
        DECLARE @ImagenUrlNorm NVARCHAR(500) = NULLIF(LTRIM(RTRIM(@ImagenUrl)), N'');
        DECLARE @CodigoUbigeoNorm CHAR(6) = NULLIF(LTRIM(RTRIM(@CodigoUbigeo)), '');
        DECLARE @TipoRegistroNorm CHAR(1) = UPPER(ISNULL(NULLIF(LTRIM(RTRIM(@TipoRegistro)), ''), 'U'));
        DECLARE @PerfilPublicoId INT = NULL;
        DECLARE @HoyLocal DATE = CONVERT(DATE, DATEADD(HOUR, -5, SYSUTCDATETIME()));
        DECLARE @FechaCorteSemanal DATE = DATEADD(DAY, -6, @HoyLocal);

        IF @UsuarioIdNorm IS NULL
            RAISERROR(N'El usuario es obligatorio.', 16, 1);

        IF @ImagenUrlNorm IS NULL
            RAISERROR(N'La imagen del boletin es obligatoria.', 16, 1);

        IF @FechaEvento IS NULL
            RAISERROR(N'La fecha del evento es obligatoria.', 16, 1);

        IF @CodigoUbigeoNorm IS NULL OR LEN(@CodigoUbigeoNorm) <> 6
            RAISERROR(N'El distrito del evento es obligatorio.', 16, 1);

        IF @TipoRegistroNorm NOT IN ('U', 'A')
            RAISERROR(N'El tipo de registro del boletin no es valido.', 16, 1);

        IF NOT EXISTS (SELECT 1 FROM dbo.AspNetUsers u WHERE u.Id = @UsuarioIdNorm)
            RAISERROR(N'No se encontro el usuario del boletin.', 16, 1);

        IF NOT EXISTS (SELECT 1 FROM dbo.UbigeoDistritos d WHERE d.CodigoUbigeo = @CodigoUbigeoNorm AND d.Activo = 1)
            RAISERROR(N'El ubigeo del boletin no existe o esta inactivo.', 16, 1);

        SELECT TOP (1) @PerfilPublicoId = p.Id
        FROM dbo.UsuariosPublicosPerfil p
        WHERE p.UsuarioId = @UsuarioIdNorm;

        IF ISNULL(@EsAdministradorCarga, 0) = 0
           AND @TipoRegistroNorm = 'U'
           AND @IdBoletin IS NULL
           AND EXISTS
           (
               SELECT 1
               FROM dbo.BoletinesDeportivos b
               WHERE b.UsuarioId = @UsuarioIdNorm
                 AND CONVERT(DATE, DATEADD(HOUR, -5, b.FechaCreacion)) >= @FechaCorteSemanal
           )
        BEGIN
            RAISERROR(N'Solo puedes publicar un boletin por semana.', 16, 1);
        END;

        IF @IdBoletin IS NULL
        BEGIN
            INSERT INTO dbo.BoletinesDeportivos
            (
                UsuarioId, PerfilPublicoId, Titulo, Descripcion, ImagenUrl, FechaEvento, CodigoUbigeo,
                TipoRegistro, Activo, FechaCreacion, UsuarioCreacion
            )
            VALUES
            (
                @UsuarioIdNorm, @PerfilPublicoId, @TituloNorm, @DescripcionNorm, @ImagenUrlNorm, @FechaEvento, @CodigoUbigeoNorm,
                @TipoRegistroNorm, ISNULL(@Activo, 1), SYSDATETIME(), @Usuario
            );

            SELECT CAST(SCOPE_IDENTITY() AS INT);
            RETURN;
        END;

        IF NOT EXISTS (SELECT 1 FROM dbo.BoletinesDeportivos b WHERE b.IdBoletin = @IdBoletin)
            RAISERROR(N'No se encontro el boletin a actualizar.', 16, 1);

        IF ISNULL(@EsAdministradorCarga, 0) = 0
           AND EXISTS
           (
               SELECT 1
               FROM dbo.BoletinesDeportivos b
               WHERE b.IdBoletin = @IdBoletin
                 AND b.UsuarioId <> @UsuarioIdNorm
           )
        BEGIN
            RAISERROR(N'No puedes editar un boletin que pertenece a otro usuario.', 16, 1);
        END;

        UPDATE dbo.BoletinesDeportivos
        SET Titulo = @TituloNorm,
            Descripcion = @DescripcionNorm,
            ImagenUrl = @ImagenUrlNorm,
            FechaEvento = @FechaEvento,
            CodigoUbigeo = @CodigoUbigeoNorm,
            TipoRegistro = @TipoRegistroNorm,
            Activo = ISNULL(@Activo, Activo),
            PerfilPublicoId = @PerfilPublicoId,
            FechaActualizacion = SYSDATETIME(),
            UsuarioActualizacion = @Usuario
        WHERE IdBoletin = @IdBoletin;

        SELECT @IdBoletin;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO
