USE [DbSportCenter]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   18/04/2026
-- Description:   Registra un nuevo desafio entre usuarios publicos.
-- =============================================
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   11/05/2026
-- Description:   Valida IdDeporte contra TiposDeporteSuperMaestro para mantener coherencia del catalogo publico.
-- =============================================
CREATE OR ALTER PROCEDURE [dbo].[Sp_Desafios_Crear]
    @UsuarioIdRetador NVARCHAR(450),
    @PerfilRetadoId INT,
    @IdDeporte INT,
    @IdNivel INT,
    @Distrito CHAR(6),
    @FechaTentativa DATE,
    @HoraTentativa TIME(7),
    @CanchaSugerida NVARCHAR(150) = NULL,
    @Modalidad NVARCHAR(120),
    @Mensaje NVARCHAR(500) = NULL,
    @FormaPago NVARCHAR(120),
    @Usuario NVARCHAR(120)
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        DECLARE @UsuarioIdRetado NVARCHAR(450);
        DECLARE @CanchaSugeridaNorm NVARCHAR(150) = NULLIF(LTRIM(RTRIM(@CanchaSugerida)), N'');
        DECLARE @ModalidadNorm NVARCHAR(120) = NULLIF(LTRIM(RTRIM(@Modalidad)), N'');
        DECLARE @MensajeNorm NVARCHAR(500) = NULLIF(LTRIM(RTRIM(@Mensaje)), N'');
        DECLARE @FormaPagoNorm NVARCHAR(120) = NULLIF(LTRIM(RTRIM(@FormaPago)), N'');
        DECLARE @DistritoNorm CHAR(6) = NULLIF(LTRIM(RTRIM(@Distrito)), '');
        DECLARE @NuevoId INT;

        IF @PerfilRetadoId IS NULL OR @PerfilRetadoId <= 0
            RAISERROR('Debes seleccionar un rival valido.', 16, 1);

        IF @ModalidadNorm IS NULL
            RAISERROR('La modalidad es obligatoria.', 16, 1);

        IF @FormaPagoNorm IS NULL
            RAISERROR('La forma de pago es obligatoria.', 16, 1);

        IF NOT EXISTS (
            SELECT 1
            FROM dbo.TiposDeporteSuperMaestro tsm
            WHERE tsm.Id = @IdDeporte
              AND tsm.Activo = 1
        )
            RAISERROR('El deporte seleccionado no existe en el catalogo publico.', 16, 1);

        IF @FechaTentativa < CAST(GETDATE() AS DATE)
            RAISERROR('La fecha tentativa no puede ser anterior a hoy.', 16, 1);

        SELECT @UsuarioIdRetado = p.UsuarioId
        FROM dbo.UsuariosPublicosPerfil p
        WHERE p.Id = @PerfilRetadoId;

        IF @UsuarioIdRetado IS NULL
            RAISERROR('No se encontro el perfil retado.', 16, 1);

        IF @UsuarioIdRetador = @UsuarioIdRetado
            RAISERROR('No puedes desafiarte a ti mismo.', 16, 1);

        IF NOT EXISTS (
            SELECT 1
            FROM dbo.UsuariosPublicosPerfil p
            WHERE p.Id = @PerfilRetadoId
              AND p.BuscarDesafios = 1
        )
            RAISERROR('El equipo seleccionado no esta disponible para desafios.', 16, 1);

        INSERT INTO dbo.Desafio
        (
            IdUsuarioRetador, IdUsuarioRetado, IdDeporte, IdNivel, Distrito, FechaTentativa, HoraTentativa,
            CanchaSugerida, Modalidad, Mensaje, FormaPago, Estado, FechaCreacion, Activo, UsuarioCreacion
        )
        VALUES
        (
            @UsuarioIdRetador, @UsuarioIdRetado, @IdDeporte, @IdNivel, @DistritoNorm, @FechaTentativa, @HoraTentativa,
            @CanchaSugeridaNorm, @ModalidadNorm, @MensajeNorm, @FormaPagoNorm, N'Pendiente', SYSDATETIME(), 1, @Usuario
        );

        SET @NuevoId = SCOPE_IDENTITY();
        SELECT @NuevoId;
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
