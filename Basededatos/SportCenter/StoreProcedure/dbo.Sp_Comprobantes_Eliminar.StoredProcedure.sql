USE [DbSportCenter]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- Firma: Codex - 09/04/2026 | Anula comprobante y lo deja fuera de activos para permitir nueva emision sobre la misma reserva pagada.
-- Firma: Codex - 10/04/2026 | Al anular comprobante, libera reserva asociada y la mantiene en estado Pagada cuando el pago acumulado es 100%.
CREATE OR ALTER PROCEDURE [dbo].[Sp_Comprobantes_Eliminar]
    @NegocioId INT,
    @Id INT,
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        DECLARE @ReservaId INT;

        SELECT @ReservaId = c.ReservaId
        FROM dbo.ComprobantesElectronicos c
        WHERE c.Id = @Id
          AND c.NegocioId = @NegocioId;

        UPDATE dbo.ComprobantesElectronicos
        SET Estado = 5,
            FechaActualizacion = SYSUTCDATETIME(),
            UsuarioActualizacion = @Usuario
        WHERE Id = @Id
          AND NegocioId = @NegocioId;

        IF @@ROWCOUNT = 0
            RAISERROR('No se encontro el comprobante para eliminar en el negocio.', 16, 1);

        IF @ReservaId IS NOT NULL
        BEGIN
            UPDATE r
            SET r.Estado = CASE WHEN COALESCE(r.Adelanto, 0) >= COALESCE(r.Total, 0) THEN 4 ELSE r.Estado END,
                r.Saldo = COALESCE(r.Total, 0) - COALESCE(r.Adelanto, 0),
                r.FechaActualizacion = SYSUTCDATETIME(),
                r.UsuarioActualizacion = @Usuario
            FROM dbo.Reservas r
            INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
            INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
            WHERE r.Id = @ReservaId
              AND s.NegocioId = @NegocioId;
        END

        DECLARE @EntidadIdAudit NVARCHAR(80);
        SET @EntidadIdAudit = CONVERT(NVARCHAR(80), @Id);
        EXEC dbo.Sp_Auditoria_Registrar
            @NegocioId = @NegocioId,
            @Modulo = N'COMPROBANTES',
            @Accion = N'DELETE',
            @Entidad = N'ComprobanteElectronico',
            @EntidadId = @EntidadIdAudit,
            @Usuario = @Usuario,
            @DetalleJson = NULL;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO
