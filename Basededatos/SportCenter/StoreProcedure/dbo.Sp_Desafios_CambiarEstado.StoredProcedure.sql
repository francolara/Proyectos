USE [DbSportCenter]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   18/04/2026
-- Description:   Actualiza el estado del desafio segun el rol del usuario participante.
-- =============================================
CREATE OR ALTER PROCEDURE [dbo].[Sp_Desafios_CambiarEstado]
    @Id INT,
    @UsuarioId NVARCHAR(450),
    @Estado NVARCHAR(20),
    @Usuario NVARCHAR(120)
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        DECLARE @EstadoNorm NVARCHAR(20) = UPPER(LTRIM(RTRIM(@Estado)));
        DECLARE @EstadoActual NVARCHAR(20);
        DECLARE @IdUsuarioRetador NVARCHAR(450);
        DECLARE @IdUsuarioRetado NVARCHAR(450);

        SELECT
            @EstadoActual = d.Estado,
            @IdUsuarioRetador = d.IdUsuarioRetador,
            @IdUsuarioRetado = d.IdUsuarioRetado
        FROM dbo.Desafio d
        WHERE d.Id = @Id
          AND d.Activo = 1;

        IF @EstadoActual IS NULL
            RAISERROR('No se encontro el desafio solicitado.', 16, 1);

        IF @EstadoNorm = N'CANCELADO'
        BEGIN
            IF @IdUsuarioRetador <> @UsuarioId
                RAISERROR('Solo el retador puede cancelar un desafio.', 16, 1);
            IF @EstadoActual <> N'Pendiente'
                RAISERROR('Solo se pueden cancelar desafios pendientes.', 16, 1);
        END
        ELSE IF @EstadoNorm IN (N'ACEPTADO', N'RECHAZADO')
        BEGIN
            IF @IdUsuarioRetado <> @UsuarioId
                RAISERROR('Solo el usuario retado puede responder el desafio.', 16, 1);
            IF @EstadoActual <> N'Pendiente'
                RAISERROR('Solo se pueden responder desafios pendientes.', 16, 1);
        END
        ELSE IF @EstadoNorm = N'FINALIZADO'
        BEGIN
            IF @UsuarioId NOT IN (@IdUsuarioRetador, @IdUsuarioRetado)
                RAISERROR('Solo un participante puede finalizar el desafio.', 16, 1);
            IF @EstadoActual <> N'Aceptado'
                RAISERROR('Solo se pueden finalizar desafios aceptados.', 16, 1);
        END
        ELSE
        BEGIN
            RAISERROR('Estado de desafio no permitido.', 16, 1);
        END

        UPDATE dbo.Desafio
        SET Estado = CASE @EstadoNorm
                        WHEN N'CANCELADO' THEN N'Cancelado'
                        WHEN N'ACEPTADO' THEN N'Aceptado'
                        WHEN N'RECHAZADO' THEN N'Rechazado'
                        WHEN N'FINALIZADO' THEN N'Finalizado'
                     END,
            FechaRespuesta = CASE WHEN @EstadoNorm IN (N'ACEPTADO', N'RECHAZADO', N'FINALIZADO') THEN SYSDATETIME() ELSE FechaRespuesta END,
            UsuarioActualizacion = @Usuario
        WHERE Id = @Id;
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
