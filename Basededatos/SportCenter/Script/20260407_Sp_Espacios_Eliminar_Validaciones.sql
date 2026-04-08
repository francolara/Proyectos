USE [DbSportCenter]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- SOURCE: 03_Sedes_Espacios.sql (linea 340)
-- Firma: Codex - 07/04/2026 | Inactivacion de espacio bloqueada si la sede tiene reservas activas futuras.
CREATE OR ALTER PROCEDURE [dbo].[Sp_Espacios_Eliminar]
    @NegocioId INT,
    @Id INT,
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        DECLARE @SedeId INT;
        DECLARE @Hoy DATE = CAST(GETDATE() AS DATE);
        DECLARE @HoraActual TIME = CAST(GETDATE() AS TIME);
        DECLARE @ReservasActivas NVARCHAR(MAX);

        SELECT @SedeId = e.SedeId
        FROM dbo.EspaciosDeportivos e
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        WHERE e.Id = @Id
          AND s.NegocioId = @NegocioId;

        IF @SedeId IS NULL
            RAISERROR('No se encontro el espacio deportivo para inactivar.', 16, 1);

        SELECT @ReservasActivas =
            STRING_AGG(
                CONCAT(
                    N'#', CONVERT(NVARCHAR(20), r.Id),
                    N' ', CONVERT(NVARCHAR(10), r.Fecha, 103),
                    N' ', LEFT(CONVERT(NVARCHAR(8), r.HoraInicio, 108), 5),
                    N'-', LEFT(CONVERT(NVARCHAR(8), r.HoraFin, 108), 5),
                    N' ', e2.Nombre
                ),
                N'; '
            )
        FROM dbo.Reservas r
        INNER JOIN dbo.EspaciosDeportivos e2 ON e2.Id = r.EspacioDeportivoId
        WHERE e2.SedeId = @SedeId
          AND r.Estado IN (1, 2, 3, 4)
          AND (r.Fecha > @Hoy OR (r.Fecha = @Hoy AND r.HoraFin > @HoraActual));

        IF @ReservasActivas IS NOT NULL
            RAISERROR('No se puede inactivar el espacio. La sede tiene reservas activas futuras: %s. Cancela esas reservas para realizar la accion.', 16, 1, @ReservasActivas);

        UPDATE e
        SET
            e.Estado = 3,
            e.FechaActualizacion = SYSUTCDATETIME(),
            e.UsuarioActualizacion = @Usuario
        FROM dbo.EspaciosDeportivos e
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        WHERE e.Id = @Id
          AND s.NegocioId = @NegocioId;

        IF @@ROWCOUNT > 0
        BEGIN
            DECLARE @EntidadIdAudit NVARCHAR(80);
            SET @EntidadIdAudit = CONVERT(NVARCHAR(80), @Id);
            EXEC dbo.Sp_Auditoria_Registrar @NegocioId = @NegocioId, @Modulo = N'ESPACIOS', @Accion = N'INACTIVATE', @Entidad = N'EspacioDeportivo', @EntidadId = @EntidadIdAudit, @Usuario = @Usuario, @DetalleJson = NULL;
        END
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO
