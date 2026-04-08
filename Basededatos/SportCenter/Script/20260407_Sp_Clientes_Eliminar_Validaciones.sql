USE [DbSportCenter]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- Firma: Codex - 06/04/2026 | Reemplaza eliminacion por inactivacion logica usando Clientes.NegocioId, sin tabla puente.
-- Firma: Codex - 07/04/2026 | Bloquea inactivacion cuando existen reservas activas futuras del cliente.
CREATE OR ALTER PROCEDURE dbo.Sp_Clientes_Eliminar
    @NegocioId INT,
    @Id INT,
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        DECLARE @Hoy DATE = CAST(GETDATE() AS DATE);
        DECLARE @HoraActual TIME = CAST(GETDATE() AS TIME);
        DECLARE @ReservasActivas NVARCHAR(MAX);

        SELECT @ReservasActivas =
            STRING_AGG(
                CONCAT(
                    N'#', CONVERT(NVARCHAR(20), r.Id),
                    N' ', CONVERT(NVARCHAR(10), r.Fecha, 103),
                    N' ', LEFT(CONVERT(NVARCHAR(8), r.HoraInicio, 108), 5),
                    N'-', LEFT(CONVERT(NVARCHAR(8), r.HoraFin, 108), 5),
                    N' ', e.Nombre
                ),
                N'; '
            )
        FROM dbo.Reservas r
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        WHERE s.NegocioId = @NegocioId
          AND r.ClienteId = @Id
          AND r.Estado IN (1, 2, 3, 4)
          AND (r.Fecha > @Hoy OR (r.Fecha = @Hoy AND r.HoraFin > @HoraActual));

        IF @ReservasActivas IS NOT NULL
            RAISERROR('No se puede inactivar el cliente. Tiene reservas activas futuras: %s. Cancela esas reservas para realizar la accion.', 16, 1, @ReservasActivas);

        UPDATE dbo.Clientes
        SET Activo = 0,
            FechaActualizacion = SYSUTCDATETIME(),
            UsuarioActualizacion = @Usuario
        WHERE Id = @Id
          AND NegocioId = @NegocioId
          AND Activo = 1;

        IF @@ROWCOUNT = 0
            RAISERROR('No se encontro el cliente para inactivar en el negocio.', 16, 1);

        DECLARE @EntidadIdAudit NVARCHAR(80);
        SET @EntidadIdAudit = CONVERT(NVARCHAR(80), @Id);
        EXEC dbo.Sp_Auditoria_Registrar
            @NegocioId = @NegocioId,
            @Modulo = N'CLIENTES',
            @Accion = N'INACTIVATE',
            @Entidad = N'Cliente',
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
