/*
Firma: Codex - 09/04/2026
Descripcion: Actualiza Sp_Comprobantes_Eliminar para anular comprobante activo y liberar la reserva para nueva emision.

Firma: Codex - 10/04/2026
Descripcion: Al anular comprobante, libera reserva asociada para nueva emision y recalcula saldo manteniendo estado Pagada si ya cubre 100%.
*/
USE [DbSportCenter]
GO

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
