USE [DbSportCenter]
GO
/****** Object:  StoredProcedure [dbo].[Sp_Reservas_CambiarEstadoRapido]    Script Date: 3/04/2026 23:18:34 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- SOURCE: 16_Reservas_CheckIn_CheckOut.sql (linea 8)
-- Firma: Codex - 05/04/2026 | Se retira estado En uso/Check-in y se normaliza Finalizada->Pagada, No Show->No Asistio.
-- Firma: Codex - 06/04/2026 | Confirmar reserva valida politica de pago del negocio (sin pago, adelanto minimo %, o pago total 100%).
-- Firma: Codex - 10/04/2026 | Bloquea cambio a Cancelada cuando la reserva tiene pagos registrados.
CREATE OR ALTER   PROCEDURE [dbo].[Sp_Reservas_CambiarEstadoRapido]
    @NegocioId INT,
    @Id INT,
    @NuevoEstado INT,
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        DECLARE @EstadoActual INT;
        DECLARE @TotalReserva DECIMAL(10,2);
        DECLARE @PagadoActual DECIMAL(10,2);
        DECLARE @PoliticaConfirmacionPago TINYINT;
        DECLARE @PorcentajeAdelantoMinimo DECIMAL(5,2);
        DECLARE @PagoMinimoRequerido DECIMAL(10,2);

        SELECT
            @EstadoActual = r.Estado,
            @TotalReserva = COALESCE(r.Total, 0),
            @PagadoActual = COALESCE(r.Adelanto, 0),
            @PoliticaConfirmacionPago = COALESCE(n.PoliticaConfirmacionPago, 0),
            @PorcentajeAdelantoMinimo = n.PorcentajeAdelantoMinimo
        FROM dbo.Reservas r
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        INNER JOIN dbo.Negocios n ON n.Id = s.NegocioId
        WHERE r.Id = @Id
          AND s.NegocioId = @NegocioId;

        IF @EstadoActual IS NULL
            RAISERROR('No se encontro la reserva para cambio de estado.', 16, 1);

        IF @NuevoEstado NOT IN (2, 4, 5, 6)
            RAISERROR('Estado no permitido para cambio rapido.', 16, 1);

        IF @EstadoActual IN (5, 6)
            RAISERROR('La reserva ya esta cancelada o marcada como no asistio.', 16, 1);

        IF @NuevoEstado = 4 AND @EstadoActual NOT IN (1, 2, 3)
            RAISERROR('Pagada solo permitido para reservas pendientes, confirmadas o en uso historico.', 16, 1);

        IF @NuevoEstado = 6 AND @EstadoActual NOT IN (1, 2, 3)
            RAISERROR('No Asistio solo permitido para reservas pendientes, confirmadas o en uso historico.', 16, 1);

        IF @NuevoEstado = 5
        BEGIN
            IF EXISTS (SELECT 1 FROM dbo.Pagos p WHERE p.ReservaId = @Id AND COALESCE(p.Monto, 0) > 0)
                RAISERROR('No se puede cancelar la reserva porque tiene pagos registrados. Elimina los pagos para continuar.', 16, 1);
        END

        IF @NuevoEstado = 2
        BEGIN
            IF @PoliticaConfirmacionPago = 1
            BEGIN
                IF @PorcentajeAdelantoMinimo IS NULL OR @PorcentajeAdelantoMinimo <= 0 OR @PorcentajeAdelantoMinimo > 100
                    RAISERROR('La configuracion del porcentaje minimo de adelanto no es valida para confirmar.', 16, 1);

                SET @PagoMinimoRequerido = ROUND(@TotalReserva * (@PorcentajeAdelantoMinimo / 100.0), 2);
                IF @PagadoActual < @PagoMinimoRequerido
                    RAISERROR('No se puede confirmar: el pago actual no alcanza el adelanto minimo configurado.', 16, 1);
            END
            ELSE IF @PoliticaConfirmacionPago = 2
            BEGIN
                IF @PagadoActual < @TotalReserva
                    RAISERROR('No se puede confirmar: se requiere pago total (100%).', 16, 1);
            END
        END

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
                    WHEN 2 THEN N'CONFIRMAR'
                    WHEN 4 THEN N'PAGADA'
                    WHEN 5 THEN N'CANCELAR'
                    WHEN 6 THEN N'NOASISTIO'
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
