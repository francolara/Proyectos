
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   08/06/2026
-- Description:   Registra una sola resena por reserva publica cuando la reserva ya fue confirmada, pagada o completada.
-- =============================================
-- Firma:         FRANCO LARA - 08/06/2026 | Valida pertenencia de la reserva al usuario publico autenticado, asegura una sola reseña por reserva y guarda alias publico editable sin revelar el nombre real.
CREATE OR ALTER PROCEDURE [dbo].[Sp_UsuariosPublicos_ResenaCrear]
    @UsuarioId NVARCHAR(450),
    @ReservaId INT,
    @AliasPublico NVARCHAR(120),
    @Comentario NVARCHAR(800),
    @Usuario NVARCHAR(120)
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        DECLARE @AliasPublicoNorm NVARCHAR(120) = NULLIF(REPLACE(LTRIM(RTRIM(@AliasPublico)), N' ', N''), N'');
        DECLARE @ComentarioNorm NVARCHAR(800) = NULLIF(LTRIM(RTRIM(@Comentario)), N'');
        DECLARE @ResenaId INT;

        IF NULLIF(LTRIM(RTRIM(@UsuarioId)), N'') IS NULL
            RAISERROR('El usuario publico es obligatorio.', 16, 1);

        IF @ReservaId IS NULL OR @ReservaId <= 0
            RAISERROR('La reserva es obligatoria.', 16, 1);

        IF @ComentarioNorm IS NULL
            RAISERROR('El comentario de la resena es obligatorio.', 16, 1);

        IF NOT EXISTS (
            SELECT 1
            FROM dbo.ReservasUsuariosPublicos rup
            WHERE rup.ReservaId = @ReservaId
              AND rup.UsuarioId = @UsuarioId
        )
            RAISERROR('La reserva no pertenece al usuario autenticado.', 16, 1);

        IF NOT EXISTS (
            SELECT 1
            FROM dbo.Reservas r
            WHERE r.Id = @ReservaId
              AND r.Estado IN (2, 3, 4)
        )
            RAISERROR('Solo se puede registrar una resena para reservas confirmadas, pagadas o completadas.', 16, 1);

        IF EXISTS (
            SELECT 1
            FROM dbo.ReservasUsuariosPublicosResenas rr
            WHERE rr.ReservaId = @ReservaId
        )
            RAISERROR('La reserva ya cuenta con una resena registrada.', 16, 1);

        INSERT INTO dbo.ReservasUsuariosPublicosResenas
        (
            ReservaId,
            UsuarioId,
            AliasPublico,
            Comentario,
            FechaCreacion,
            UsuarioCreacion
        )
        VALUES
        (
            @ReservaId,
            @UsuarioId,
            COALESCE(@AliasPublicoNorm, N'@JugadorAnonimo'),
            @ComentarioNorm,
            SYSDATETIME(),
            @Usuario
        );

        SET @ResenaId = SCOPE_IDENTITY();
        SELECT @ResenaId;
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
