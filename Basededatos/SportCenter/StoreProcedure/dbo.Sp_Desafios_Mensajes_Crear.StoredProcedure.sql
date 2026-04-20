USE [DbSportCenter]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   18/04/2026
-- Description:   Registra mensajes internos de desafios permitiendo coordinacion antes y despues de la aceptacion.
-- =============================================
CREATE OR ALTER PROCEDURE [dbo].[Sp_Desafios_Mensajes_Crear]
    @IdDesafio INT,
    @UsuarioId NVARCHAR(450),
    @Mensaje NVARCHAR(500),
    @Usuario NVARCHAR(120)
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        DECLARE @MensajeNorm NVARCHAR(500) = NULLIF(LTRIM(RTRIM(@Mensaje)), N'');
        DECLARE @EstadoActual NVARCHAR(20);
        DECLARE @IdMensaje INT;

        IF @IdDesafio IS NULL OR @IdDesafio <= 0
            RAISERROR('Debes indicar un desafio valido.', 16, 1);

        IF @MensajeNorm IS NULL
            RAISERROR('Debes escribir un mensaje para el desafio.', 16, 1);

        SELECT
            @EstadoActual = d.Estado
        FROM dbo.Desafio d
        WHERE d.Id = @IdDesafio
          AND d.Activo = 1
          AND (d.IdUsuarioRetador = @UsuarioId OR d.IdUsuarioRetado = @UsuarioId);

        IF @EstadoActual IS NULL
            RAISERROR('No tienes acceso al desafio indicado.', 16, 1);

        IF @EstadoActual NOT IN (N'Pendiente', N'Aceptado')
            RAISERROR('Solo se permiten mensajes en desafios pendientes o aceptados.', 16, 1);

        INSERT INTO dbo.DesafioMensaje
        (
            IdDesafio,
            UsuarioIdEmisor,
            Mensaje,
            FechaRegistro,
            Activo,
            UsuarioCreacion
        )
        VALUES
        (
            @IdDesafio,
            @UsuarioId,
            @MensajeNorm,
            SYSDATETIME(),
            1,
            @Usuario
        );

        SET @IdMensaje = SCOPE_IDENTITY();

        SELECT @IdMensaje;
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
