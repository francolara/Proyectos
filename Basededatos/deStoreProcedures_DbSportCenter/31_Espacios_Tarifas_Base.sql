-- =============================================
-- Author:        FRANCO LARA
-- Create date:   28/03/2026
-- Description:   Tarifas multiples por espacio (dia y rango horario) desde registro de espacios deportivos.
-- Firma:         Codex - 28/03/2026
-- =============================================

IF OBJECT_ID(N'dbo.Tarifas', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.Tarifas
    (
        Id INT IDENTITY(1,1) NOT NULL,
        EspacioDeportivoId INT NOT NULL,
        DiaSemana INT NOT NULL,
        HoraInicio TIME NOT NULL,
        HoraFin TIME NOT NULL,
        Precio DECIMAL(10,2) NOT NULL,
        Activa BIT NOT NULL CONSTRAINT DF_Tarifas_Activa DEFAULT (1),
        FechaCreacion DATETIME2 NOT NULL CONSTRAINT DF_Tarifas_FechaCreacion DEFAULT (SYSUTCDATETIME()),
        UsuarioCreacion NVARCHAR(200) NULL,
        FechaActualizacion DATETIME2 NULL,
        UsuarioActualizacion NVARCHAR(200) NULL,
        CONSTRAINT PK_Tarifas PRIMARY KEY CLUSTERED (Id),
        CONSTRAINT FK_Tarifas_EspaciosDeportivos_EspacioDeportivoId FOREIGN KEY (EspacioDeportivoId) REFERENCES dbo.EspaciosDeportivos(Id)
    );
END;
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Espacios_ObtenerPorId
    @NegocioId INT,
    @Id INT
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT
            e.Id,
            e.SedeId,
            e.TipoDeporteId,
            e.TipoSueloId,
            e.Codigo,
            e.Nombre,
            e.Capacidad,
            e.TieneIluminacion,
            e.Techada,
            e.Estado,
            (
                SELECT
                    t.DiaSemana,
                    CONVERT(NVARCHAR(8), t.HoraInicio, 108) AS HoraInicio,
                    CONVERT(NVARCHAR(8), t.HoraFin, 108) AS HoraFin,
                    t.Precio
                FROM dbo.Tarifas t
                WHERE t.EspacioDeportivoId = e.Id
                  AND t.Activa = 1
                ORDER BY t.DiaSemana, t.HoraInicio
                FOR JSON PATH
            ) AS TarifasJson
        FROM dbo.EspaciosDeportivos e
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        WHERE e.Id = @Id
          AND s.NegocioId = @NegocioId;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Espacios_Crear
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
        IF NOT EXISTS (SELECT 1 FROM dbo.Sedes WHERE Id = @SedeId AND NegocioId = @NegocioId)
            RAISERROR('Sede invalida para el negocio.', 16, 1);

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

        INSERT INTO dbo.EspaciosDeportivos
        (
            SedeId, TipoDeporteId, TipoSueloId, Codigo, Nombre, Capacidad,
            TieneIluminacion, Techada, Estado, FechaCreacion, UsuarioCreacion
        )
        VALUES
        (
            @SedeId, @TipoDeporteId, @TipoSueloId, @Codigo, @Nombre, @Capacidad,
            @TieneIluminacion, @Techada, @Estado, SYSUTCDATETIME(), @Usuario
        );

        DECLARE @Id INT;
        SET @Id = SCOPE_IDENTITY();

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
            @Accion = N'CREATE',
            @Entidad = N'EspacioDeportivo',
            @EntidadId = @EntidadIdAudit,
            @Usuario = @Usuario,
            @DetalleJson = NULL;

        SELECT @Id;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Espacios_Actualizar
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
