-- =============================================
-- Author:        FRANCO LARA
-- Create date:   27/03/2026
-- Description:   Sprint 6 - Cambio rapido de estado de reservas (check-in/check-out/no-show).
-- Firma:         Codex - 30/03/2026 | Sp_Reservas_CambiarEstadoRapido ahora devuelve error si la reserva no existe para el negocio.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.Sp_Reservas_CambiarEstadoRapido
    @NegocioId INT,
    @Id INT,
    @NuevoEstado INT,
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        DECLARE @EstadoActual INT;

        SELECT @EstadoActual = r.Estado
        FROM dbo.Reservas r
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        WHERE r.Id = @Id
          AND s.NegocioId = @NegocioId;

        IF @EstadoActual IS NULL
            RAISERROR('No se encontro la reserva para cambio de estado.', 16, 1);

        IF @NuevoEstado NOT IN (3, 4, 6)
            RAISERROR('Estado no permitido para cambio rapido.', 16, 1);

        IF @EstadoActual IN (5, 6)
            RAISERROR('La reserva ya esta cancelada o marcada como no asistio.', 16, 1);

        IF @NuevoEstado = 3 AND @EstadoActual NOT IN (1, 2)
            RAISERROR('Check-in solo permitido para reservas pendientes o confirmadas.', 16, 1);

        IF @NuevoEstado = 4 AND @EstadoActual <> 3
            RAISERROR('Check-out solo permitido para reservas en uso.', 16, 1);

        IF @NuevoEstado = 6 AND @EstadoActual NOT IN (1, 2)
            RAISERROR('No-show solo permitido para reservas pendientes o confirmadas.', 16, 1);

        UPDATE dbo.Reservas
        SET Estado = @NuevoEstado,
            FechaActualizacion = SYSUTCDATETIME(),
            UsuarioActualizacion = @Usuario
        WHERE Id = @Id;

        IF @@ROWCOUNT > 0
        BEGIN
            DECLARE @EntidadIdAudit NVARCHAR(80);
            DECLARE @AccionAudit NVARCHAR(30);

            SET @EntidadIdAudit = CONVERT(NVARCHAR(80), @Id);
            SET @AccionAudit =
                CASE @NuevoEstado
                    WHEN 3 THEN N'CHECKIN'
                    WHEN 4 THEN N'CHECKOUT'
                    WHEN 6 THEN N'NOSHOW'
                    ELSE N'EDIT'
                END;

            EXEC dbo.Sp_Auditoria_Registrar
                @NegocioId = @NegocioId,
                @Modulo = N'RESERVAS',
                @Accion = @AccionAudit,
                @Entidad = N'Reserva',
                @EntidadId = @EntidadIdAudit,
                @Usuario = @Usuario,
                @DetalleJson = NULL;
        END;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO
