USE [DbSportCenter]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- =============================================
-- Author:        FRANCO LARA
-- Create date:   07/04/2026
-- Description:   Validaciones de inactivacion clientes/espacios con reservas activas futuras y ajustes en listado de reservas.
-- Firma:         Codex - 07/04/2026
-- =============================================

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

CREATE OR ALTER PROCEDURE [dbo].[Sp_Reservas_Listar]
    @NegocioId INT,
    @FechaDesde DATE = NULL,
    @FechaHasta DATE = NULL,
    @SedeId INT = NULL,
    @EspacioDeportivoId INT = NULL,
    @Estado INT = NULL,
    @EstadosCsv NVARCHAR(200) = NULL
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        DECLARE @EstadosNormalizados NVARCHAR(200);
        SET @EstadosNormalizados = NULLIF(REPLACE(REPLACE(LTRIM(RTRIM(@EstadosCsv)), N' ', N''), N';', N','), N'');

        SELECT TOP (300)
            r.Id,
            CAST(
                CASE
                    WHEN NULLIF(LTRIM(RTRIM(c.NombreEquipo)), N'') IS NULL THEN c.NombresORazonSocial
                    ELSE CONCAT(c.NombresORazonSocial, N' - Equipo: ', LTRIM(RTRIM(c.NombreEquipo)))
                END
                AS NVARCHAR(250)
            ) AS Cliente,
            e.Nombre AS Espacio,
            s.Nombre AS Sede,
            r.Fecha,
            r.HoraInicio,
            r.HoraFin,
            r.Total,
            r.Adelanto,
            (r.Total - r.Adelanto) AS SaldoPendiente,
            CAST(r.Estado AS NVARCHAR(20)) AS Estado
        FROM dbo.Reservas r
        INNER JOIN dbo.Clientes c ON c.Id = r.ClienteId
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        WHERE s.NegocioId = @NegocioId
          AND (@FechaDesde IS NULL OR r.Fecha >= @FechaDesde)
          AND (@FechaHasta IS NULL OR r.Fecha <= @FechaHasta)
          AND (@SedeId IS NULL OR s.Id = @SedeId)
          AND (@EspacioDeportivoId IS NULL OR e.Id = @EspacioDeportivoId)
          AND
          (
              (@Estado IS NOT NULL AND ((@Estado = 4 AND r.Estado IN (3, 4)) OR (@Estado <> 4 AND r.Estado = @Estado)))
              OR
              (
                  @Estado IS NULL
                  AND
                  (
                      @EstadosNormalizados IS NULL
                      OR EXISTS
                      (
                          SELECT 1
                          FROM STRING_SPLIT(@EstadosNormalizados, N',') estados
                          WHERE (TRY_CAST(estados.value AS INT) = 4 AND r.Estado IN (3, 4)) OR (TRY_CAST(estados.value AS INT) <> 4 AND TRY_CAST(estados.value AS INT) = r.Estado)
                      )
                  )
              )
          )
        ORDER BY r.Fecha ASC, r.HoraInicio ASC;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO
