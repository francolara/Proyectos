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
        DECLARE @Id INT;

        IF @NombresNorm IS NULL OR @ApellidosNorm IS NULL
            RAISERROR('Nombres y apellidos son obligatorios.', 16, 1);

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
                FechaCreacion, UsuarioCreacion
            )
            VALUES
            (
                @UsuarioId, @TipoDocumentoNorm, @NumeroDocumentoNorm, @NombresNorm, @ApellidosNorm, @NombreEquipoNorm,
                @TelefonoNorm, @CorreoNorm, @FechaNacimiento, @CodigoUbigeoNorm,
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
