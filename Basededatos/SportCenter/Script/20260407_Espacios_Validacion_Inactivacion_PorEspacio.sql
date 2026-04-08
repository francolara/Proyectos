/*
Firma: Codex - 07/04/2026
Descripcion: Corrige validacion de inactivacion/mantenimiento de espacios para evaluar solo reservas futuras del espacio objetivo.
*/
USE [DbSportCenter]
GO

CREATE OR ALTER PROCEDURE [dbo].[Sp_Espacios_Actualizar]
    @Id INT,
    @NegocioId INT,
    @SedeId INT,
    @TipoDeporteId INT,
    @TipoSueloId INT,
    @Codigo NVARCHAR(20),
    @Nombre NVARCHAR(150),
    @Capacidad INT,
    @TieneIluminacion BIT,
    @Techada BIT,
    @Estado INT,
    @TarifasJson NVARCHAR(MAX),
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        DECLARE @EstadoActual INT;
        DECLARE @SedeActualId INT;
        DECLARE @Hoy DATE = CAST(GETDATE() AS DATE);
        DECLARE @HoraActual TIME = CAST(GETDATE() AS TIME);
        DECLARE @ReservasActivas NVARCHAR(MAX);

        SELECT
            @EstadoActual = e.Estado,
            @SedeActualId = e.SedeId
        FROM dbo.EspaciosDeportivos e
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        WHERE e.Id = @Id
          AND s.NegocioId = @NegocioId;

        IF @EstadoActual IS NULL
            RAISERROR('No se encontro el espacio deportivo para actualizar.', 16, 1);

        IF @Estado IN (2, 3) AND ISNULL(@EstadoActual, 0) NOT IN (2, 3)
        BEGIN
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
            WHERE r.EspacioDeportivoId = @Id
              AND r.Estado IN (1, 2, 3, 4)
              AND (r.Fecha > @Hoy OR (r.Fecha = @Hoy AND r.HoraFin > @HoraActual));

            IF @ReservasActivas IS NOT NULL
                RAISERROR('No se puede cambiar a mantenimiento/inactivo. El espacio tiene reservas activas futuras: %s. Cancela esas reservas para realizar la accion.', 16, 1, @ReservasActivas);
        END;

        IF ISNULL(LEN(LTRIM(RTRIM(@TarifasJson))), 0) = 0
            RAISERROR('Debes registrar al menos una tarifa.', 16, 1);

        DECLARE @Tarifas TABLE
        (
            Id INT IDENTITY(1,1) NOT NULL,
            DiaSemana INT NOT NULL,
            HoraInicio TIME NOT NULL,
            HoraFin TIME NOT NULL,
            Precio DECIMAL(10,2) NOT NULL
        );

        INSERT INTO @Tarifas (DiaSemana, HoraInicio, HoraFin, Precio)
        SELECT
            j.DiaSemana,
            TRY_CONVERT(TIME, j.HoraInicio),
            TRY_CONVERT(TIME, j.HoraFin),
            j.Precio
        FROM OPENJSON(@TarifasJson)
        WITH
        (
            DiaSemana INT '$.diaSemana',
            HoraInicio NVARCHAR(8) '$.horaInicio',
            HoraFin NVARCHAR(8) '$.horaFin',
            Precio DECIMAL(10,2) '$.precio'
        ) j;

        IF NOT EXISTS (SELECT 1 FROM @Tarifas)
            RAISERROR('Debes registrar al menos una tarifa valida.', 16, 1);

        IF EXISTS (SELECT 1 FROM @Tarifas WHERE DiaSemana NOT BETWEEN 0 AND 6 OR HoraInicio IS NULL OR HoraFin IS NULL OR HoraFin <= HoraInicio OR Precio <= 0)
            RAISERROR('Hay tarifas con dia, horario o precio invalido.', 16, 1);

        IF EXISTS
        (
            SELECT 1
            FROM @Tarifas a
            INNER JOIN @Tarifas b ON a.Id < b.Id
                AND a.DiaSemana = b.DiaSemana
                AND a.HoraInicio < b.HoraFin
                AND a.HoraFin > b.HoraInicio
        )
            RAISERROR('Existen rangos de tarifas superpuestos en el mismo dia.', 16, 1);

        UPDATE e
        SET
            e.SedeId = @SedeId,
            e.TipoDeporteId = @TipoDeporteId,
            e.TipoSueloId = @TipoSueloId,
            e.Codigo = @Codigo,
            e.Nombre = @Nombre,
            e.Capacidad = @Capacidad,
            e.TieneIluminacion = @TieneIluminacion,
            e.Techada = @Techada,
            e.Estado = @Estado,
            e.FechaActualizacion = SYSUTCDATETIME(),
            e.UsuarioActualizacion = @Usuario
        FROM dbo.EspaciosDeportivos e
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        WHERE e.Id = @Id
          AND s.NegocioId = @NegocioId;

        IF @@ROWCOUNT = 0
            RAISERROR('No se encontro el espacio deportivo para actualizar.', 16, 1);

        UPDATE dbo.Tarifas
        SET Activa = 0
        WHERE EspacioDeportivoId = @Id
          AND Activa = 1;

        INSERT INTO dbo.Tarifas
        (
            EspacioDeportivoId, DiaSemana, HoraInicio, HoraFin, Precio, Activa
        )
        SELECT
            @Id,
            t.DiaSemana,
            t.HoraInicio,
            t.HoraFin,
            t.Precio,
            1
        FROM @Tarifas t;

        DECLARE @EntidadIdAudit NVARCHAR(80);
        SET @EntidadIdAudit = CONVERT(NVARCHAR(80), @Id);

        EXEC dbo.Sp_Auditoria_Registrar
            @NegocioId = @NegocioId,
            @Modulo = N'ESPACIOS',
            @Accion = N'EDIT',
            @Entidad = N'EspacioDeportivo',
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
        WHERE r.EspacioDeportivoId = @Id
          AND r.Estado IN (1, 2, 3, 4)
          AND (r.Fecha > @Hoy OR (r.Fecha = @Hoy AND r.HoraFin > @HoraActual));

        IF @ReservasActivas IS NOT NULL
            RAISERROR('No se puede inactivar el espacio. El espacio tiene reservas activas futuras: %s. Cancela esas reservas para realizar la accion.', 16, 1, @ReservasActivas);

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