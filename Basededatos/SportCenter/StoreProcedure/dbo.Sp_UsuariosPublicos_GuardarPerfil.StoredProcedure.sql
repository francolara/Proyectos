USE [DbSportCenter]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   16/04/2026
-- Description:   Inserta o actualiza perfil del usuario publico.
-- =============================================
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   18/04/2026
-- Description:   Guarda configuracion de desafios y detalle general del equipo dentro del perfil publico.
-- =============================================
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   18/04/2026
-- Description:   Guarda la ubicacion y WhatsApp del equipo y usa ese ubigeo como referencia del modulo Desafios.
-- =============================================
CREATE OR ALTER PROCEDURE [dbo].[Sp_UsuariosPublicos_GuardarPerfil]
    @UsuarioId NVARCHAR(450),
    @TipoDocumento NVARCHAR(20) = N'0',
    @NumeroDocumento NVARCHAR(20) = NULL,
    @Nombres NVARCHAR(120),
    @Apellidos NVARCHAR(120),
    @NombreEquipo NVARCHAR(120) = NULL,
    @Telefono NVARCHAR(30) = NULL,
    @Correo NVARCHAR(200) = NULL,
    @FechaNacimiento DATE = NULL,
    @CodigoUbigeo CHAR(6) = NULL,
    @BuscarDesafios BIT = 0,
    @IdDeporteDesafio INT = NULL,
    @IdNivelDesafio INT = NULL,
    @ObservacionDesafio NVARCHAR(500) = NULL,
    @DetalleEquipo NVARCHAR(1000) = NULL,
    @CodigoUbigeoEquipo CHAR(6) = NULL,
    @WhatsappEquipo NVARCHAR(30) = NULL,
    @Usuario NVARCHAR(120)
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        DECLARE @TipoDocumentoNorm NVARCHAR(20) = COALESCE(NULLIF(LTRIM(RTRIM(@TipoDocumento)), N''), N'0');
        DECLARE @NumeroDocumentoNorm NVARCHAR(20) = NULLIF(LTRIM(RTRIM(@NumeroDocumento)), N'');
        DECLARE @NombresNorm NVARCHAR(120) = NULLIF(LTRIM(RTRIM(@Nombres)), N'');
        DECLARE @ApellidosNorm NVARCHAR(120) = NULLIF(LTRIM(RTRIM(@Apellidos)), N'');
        DECLARE @NombreEquipoNorm NVARCHAR(120) = NULLIF(LTRIM(RTRIM(@NombreEquipo)), N'');
        DECLARE @TelefonoNorm NVARCHAR(30) = NULLIF(LTRIM(RTRIM(@Telefono)), N'');
        DECLARE @CorreoNorm NVARCHAR(200) = NULLIF(LTRIM(RTRIM(@Correo)), N'');
        DECLARE @CodigoUbigeoNorm CHAR(6) = NULLIF(LTRIM(RTRIM(@CodigoUbigeo)), '');
        DECLARE @ObservacionDesafioNorm NVARCHAR(500) = NULLIF(LTRIM(RTRIM(@ObservacionDesafio)), N'');
        DECLARE @DetalleEquipoNorm NVARCHAR(1000) = NULLIF(LTRIM(RTRIM(@DetalleEquipo)), N'');
        DECLARE @CodigoUbigeoEquipoNorm CHAR(6) = NULLIF(LTRIM(RTRIM(@CodigoUbigeoEquipo)), '');
        DECLARE @WhatsappEquipoNorm NVARCHAR(30) = NULLIF(LTRIM(RTRIM(@WhatsappEquipo)), N'');
        DECLARE @Id INT;

        IF @NombresNorm IS NULL OR @ApellidosNorm IS NULL
            RAISERROR('Nombres y apellidos son obligatorios.', 16, 1);

        IF @BuscarDesafios = 1
        BEGIN
            IF @CodigoUbigeoEquipoNorm IS NULL
                RAISERROR('Debes seleccionar la ubicacion del equipo para habilitar desafios.', 16, 1);
            IF @IdDeporteDesafio IS NULL OR @IdDeporteDesafio <= 0
                RAISERROR('Debes seleccionar un deporte para desafios.', 16, 1);
            IF @IdNivelDesafio IS NULL OR @IdNivelDesafio <= 0
                RAISERROR('Debes seleccionar un nivel para desafios.', 16, 1);
        END

        IF EXISTS (SELECT 1 FROM dbo.UsuariosPublicosPerfil WHERE UsuarioId = @UsuarioId)
        BEGIN
            UPDATE p
            SET p.TipoDocumento = @TipoDocumentoNorm,
                p.NumeroDocumento = @NumeroDocumentoNorm,
                p.Nombres = @NombresNorm,
                p.Apellidos = @ApellidosNorm,
                p.NombreEquipo = @NombreEquipoNorm,
                p.Telefono = @TelefonoNorm,
                p.Correo = @CorreoNorm,
                p.FechaNacimiento = @FechaNacimiento,
                p.CodigoUbigeo = @CodigoUbigeoNorm,
                p.BuscarDesafios = @BuscarDesafios,
                p.IdDeporteDesafio = CASE WHEN @BuscarDesafios = 1 THEN @IdDeporteDesafio ELSE NULL END,
                p.IdNivelDesafio = CASE WHEN @BuscarDesafios = 1 THEN @IdNivelDesafio ELSE NULL END,
                p.ObservacionDesafio = @ObservacionDesafioNorm,
                p.DetalleEquipo = @DetalleEquipoNorm,
                p.CodigoUbigeoEquipo = @CodigoUbigeoEquipoNorm,
                p.WhatsappEquipo = @WhatsappEquipoNorm,
                p.FechaActualizacion = SYSDATETIME(),
                p.UsuarioActualizacion = @Usuario
            FROM dbo.UsuariosPublicosPerfil p
            WHERE p.UsuarioId = @UsuarioId;
        END
        ELSE
        BEGIN
            INSERT INTO dbo.UsuariosPublicosPerfil
            (
                UsuarioId, TipoDocumento, NumeroDocumento, Nombres, Apellidos, NombreEquipo,
                Telefono, Correo, FechaNacimiento, CodigoUbigeo,
                BuscarDesafios, IdDeporteDesafio, IdNivelDesafio, ObservacionDesafio, DetalleEquipo, CodigoUbigeoEquipo, WhatsappEquipo,
                FechaCreacion, UsuarioCreacion
            )
            VALUES
            (
                @UsuarioId, @TipoDocumentoNorm, @NumeroDocumentoNorm, @NombresNorm, @ApellidosNorm, @NombreEquipoNorm,
                @TelefonoNorm, @CorreoNorm, @FechaNacimiento, @CodigoUbigeoNorm,
                @BuscarDesafios, CASE WHEN @BuscarDesafios = 1 THEN @IdDeporteDesafio ELSE NULL END, CASE WHEN @BuscarDesafios = 1 THEN @IdNivelDesafio ELSE NULL END, @ObservacionDesafioNorm, @DetalleEquipoNorm, @CodigoUbigeoEquipoNorm, @WhatsappEquipoNorm,
                SYSDATETIME(), @Usuario
            );
        END

        SELECT TOP 1 @Id = Id
        FROM dbo.UsuariosPublicosPerfil
        WHERE UsuarioId = @UsuarioId;

        SELECT @Id;
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
