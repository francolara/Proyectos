
GO
/****** Object:  StoredProcedure [dbo].[Sp_Espacios_Crear]    Script Date: 3/04/2026 23:18:34 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- SOURCE: 31_Espacios_Tarifas_Base.sql (linea 72)
-- Firma: Codex - 18/04/2026 | Se agrega soporte de AdministracionPrivada para controlar visibilidad del espacio en portal publico.
-- Firma: FRANCO LARA - 26/05/2026 | Agrega configuracion opcional de horario propio por espacio deportivo y su persistencia.
-- Firma: FRANCO LARA - 06/06/2026 | Agrega configuracion bidireccional de espacios compartidos para bloqueo cruzado de horarios.
-- Firma: FRANCO LARA - 08/06/2026 | Separa relaciones operativas entre bloqueo directo y espacios compuestos por componentes, y agrega soporte de fotos para espacios deportivos.
CREATE OR ALTER PROCEDURE [dbo].[Sp_Espacios_Crear]
    @NegocioId INT,
    @SedeId INT,
    @TipoDeporteId INT,
    @TipoSueloId INT,
    @Codigo NVARCHAR(20),
    @Nombre NVARCHAR(150),
    @Capacidad INT,
    @TieneIluminacion BIT,
    @Techada BIT,
    @AdministracionPrivada BIT = 0,
    @FotoPrincipalUrl NVARCHAR(500) = NULL,
    @FotosUrlsCsv NVARCHAR(MAX) = NULL,
    @ConfigurarHorarioPorEspacio BIT = 0,
    @AtiendeLunes BIT = 1,
    @AtiendeMartes BIT = 1,
    @AtiendeMiercoles BIT = 1,
    @AtiendeJueves BIT = 1,
    @AtiendeViernes BIT = 1,
    @AtiendeSabado BIT = 1,
    @AtiendeDomingo BIT = 1,
    @HoraApertura TIME = '08:00',
    @HoraCierre TIME = '23:00',
    @Estado INT,
    @TarifasJson NVARCHAR(MAX),
    @TarifasFeriadoJson NVARCHAR(MAX) = N'[]',
    @TieneEspaciosCompartidos BIT = 0,
    @EspaciosDirectosIdsCsv NVARCHAR(MAX) = NULL,
    @EspaciosComponentesIdsCsv NVARCHAR(MAX) = NULL,
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        DECLARE @FotosAlternativasCount INT = 0;

        SET @FotoPrincipalUrl = NULLIF(LTRIM(RTRIM(@FotoPrincipalUrl)), N'');
        SET @FotosUrlsCsv = NULLIF(LTRIM(RTRIM(@FotosUrlsCsv)), N'');

        IF NOT EXISTS (SELECT 1 FROM dbo.Sedes WHERE Id = @SedeId AND NegocioId = @NegocioId)
            RAISERROR('Sede invalida para el negocio.', 16, 1);

        IF @FotosUrlsCsv IS NOT NULL AND LEN(LTRIM(RTRIM(@FotosUrlsCsv))) > 0
            SELECT @FotosAlternativasCount = COUNT(1)
            FROM STRING_SPLIT(@FotosUrlsCsv, N',')
            WHERE LEN(LTRIM(RTRIM(value))) > 0;

        IF @FotoPrincipalUrl IS NULL AND @FotosAlternativasCount > 0
            RAISERROR('Debes tener una foto principal cuando registres fotos alternativas.', 16, 1);

        IF (CASE WHEN @FotoPrincipalUrl IS NULL OR LEN(LTRIM(RTRIM(@FotoPrincipalUrl))) = 0 THEN 0 ELSE 1 END) + @FotosAlternativasCount > 3
            RAISERROR('Solo se permiten 3 imagenes por espacio deportivo.', 16, 1);

        IF ISNULL(LEN(LTRIM(RTRIM(@TarifasJson))), 0) = 0
            RAISERROR('Debes registrar al menos una tarifa.', 16, 1);

        IF @HoraCierre <= @HoraApertura
            RAISERROR('La hora cierre del espacio debe ser mayor que la hora apertura.', 16, 1);

        DECLARE @Tarifas TABLE
        (
            Id INT IDENTITY(1,1) NOT NULL,
            DiaSemana INT NOT NULL,
            HoraInicio TIME NOT NULL,
            HoraFin TIME NOT NULL,
            Precio DECIMAL(10,2) NOT NULL
        );
        DECLARE @TarifasFeriado TABLE
        (
            Id INT IDENTITY(1,1) NOT NULL,
            HoraInicio TIME NOT NULL,
            HoraFin TIME NOT NULL,
            Precio DECIMAL(10,2) NOT NULL
        );
        DECLARE @EspaciosDirectos TABLE
        (
            EspacioRelacionadoId INT NOT NULL PRIMARY KEY
        );
        DECLARE @EspaciosComponentes TABLE
        (
            EspacioRelacionadoId INT NOT NULL PRIMARY KEY
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

        INSERT INTO @TarifasFeriado (HoraInicio, HoraFin, Precio)
        SELECT
            TRY_CONVERT(TIME, j.HoraInicio),
            TRY_CONVERT(TIME, j.HoraFin),
            j.Precio
        FROM OPENJSON(COALESCE(@TarifasFeriadoJson, N'[]'))
        WITH
        (
            HoraInicio NVARCHAR(8) '$.horaInicio',
            HoraFin NVARCHAR(8) '$.horaFin',
            Precio DECIMAL(10,2) '$.precio'
        ) j;

        IF EXISTS (SELECT 1 FROM @TarifasFeriado WHERE HoraInicio IS NULL OR HoraFin IS NULL OR HoraFin <= HoraInicio OR Precio <= 0)
            RAISERROR('Hay tarifas por feriado con horario o precio invalido.', 16, 1);

        IF EXISTS
        (
            SELECT 1
            FROM @TarifasFeriado a
            INNER JOIN @TarifasFeriado b ON a.Id < b.Id
                AND a.HoraInicio < b.HoraFin
                AND a.HoraFin > b.HoraInicio
        )
            RAISERROR('Existen rangos de tarifas por feriado superpuestos.', 16, 1);

        IF @TieneEspaciosCompartidos = 1 AND ISNULL(LEN(LTRIM(RTRIM(@EspaciosDirectosIdsCsv))), 0) > 0
        BEGIN
            INSERT INTO @EspaciosDirectos (EspacioRelacionadoId)
            SELECT DISTINCT TRY_CONVERT(INT, LTRIM(RTRIM(value)))
            FROM STRING_SPLIT(@EspaciosDirectosIdsCsv, N',')
            WHERE TRY_CONVERT(INT, LTRIM(RTRIM(value))) IS NOT NULL;

            IF EXISTS (SELECT 1 FROM @EspaciosDirectos WHERE EspacioRelacionadoId <= 0)
                RAISERROR('La lista de espacios con bloqueo directo contiene ids invalidos.', 16, 1);

            IF EXISTS
            (
                SELECT 1
                FROM @EspaciosDirectos ec
                INNER JOIN dbo.EspaciosDeportivos er ON er.Id = ec.EspacioRelacionadoId
                WHERE er.SedeId <> @SedeId
                   OR er.Estado <> 1
            )
                RAISERROR('Solo puedes relacionar espacios activos de la misma sede para bloqueo directo.', 16, 1);
        END;

        IF @TieneEspaciosCompartidos = 1 AND ISNULL(LEN(LTRIM(RTRIM(@EspaciosComponentesIdsCsv))), 0) > 0
        BEGIN
            INSERT INTO @EspaciosComponentes (EspacioRelacionadoId)
            SELECT DISTINCT TRY_CONVERT(INT, LTRIM(RTRIM(value)))
            FROM STRING_SPLIT(@EspaciosComponentesIdsCsv, N',')
            WHERE TRY_CONVERT(INT, LTRIM(RTRIM(value))) IS NOT NULL;

            IF EXISTS (SELECT 1 FROM @EspaciosComponentes WHERE EspacioRelacionadoId <= 0)
                RAISERROR('La lista de componentes contiene ids invalidos.', 16, 1);

            IF EXISTS
            (
                SELECT 1
                FROM @EspaciosComponentes ec
                INNER JOIN dbo.EspaciosDeportivos er ON er.Id = ec.EspacioRelacionadoId
                WHERE er.SedeId <> @SedeId
                   OR er.Estado <> 1
            )
                RAISERROR('Solo puedes relacionar espacios activos de la misma sede como componentes.', 16, 1);

            IF EXISTS
            (
                SELECT 1
                FROM @EspaciosDirectos ed
                INNER JOIN @EspaciosComponentes ec ON ec.EspacioRelacionadoId = ed.EspacioRelacionadoId
            )
                RAISERROR('Un espacio no puede registrarse como bloqueo directo y componente al mismo tiempo.', 16, 1);
        END;

        INSERT INTO dbo.EspaciosDeportivos
        (
            SedeId, TipoDeporteId, TipoSueloId, Codigo, Nombre, Capacidad,
            TieneIluminacion, Techada, AdministracionPrivada, FotoPrincipalUrl, FotosUrlsCsv, Estado, FechaCreacion, UsuarioCreacion
        )
        VALUES
        (
            @SedeId, @TipoDeporteId, @TipoSueloId, @Codigo, @Nombre, @Capacidad,
            @TieneIluminacion, @Techada, @AdministracionPrivada, @FotoPrincipalUrl, @FotosUrlsCsv, @Estado, SYSUTCDATETIME(), @Usuario
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

        INSERT INTO dbo.TarifaFeriado
        (
            EspacioDeportivoId, HoraInicio, HoraFin, Precio, Activa
        )
        SELECT
            @Id,
            t.HoraInicio,
            t.HoraFin,
            t.Precio,
            1
        FROM @TarifasFeriado t;

        MERGE dbo.EspacioHorarioAtencion AS target
        USING (
            SELECT
                @Id AS EspacioDeportivoId,
                @ConfigurarHorarioPorEspacio AS ConfigurarHorarioPorEspacio,
                @AtiendeLunes AS AtiendeLunes,
                @AtiendeMartes AS AtiendeMartes,
                @AtiendeMiercoles AS AtiendeMiercoles,
                @AtiendeJueves AS AtiendeJueves,
                @AtiendeViernes AS AtiendeViernes,
                @AtiendeSabado AS AtiendeSabado,
                @AtiendeDomingo AS AtiendeDomingo,
                @HoraApertura AS HoraApertura,
                @HoraCierre AS HoraCierre,
                @Usuario AS Usuario
        ) AS source
        ON target.EspacioDeportivoId = source.EspacioDeportivoId
        WHEN MATCHED THEN
            UPDATE SET
                target.ConfigurarHorarioPorEspacio = source.ConfigurarHorarioPorEspacio,
                target.AtiendeLunes = source.AtiendeLunes,
                target.AtiendeMartes = source.AtiendeMartes,
                target.AtiendeMiercoles = source.AtiendeMiercoles,
                target.AtiendeJueves = source.AtiendeJueves,
                target.AtiendeViernes = source.AtiendeViernes,
                target.AtiendeSabado = source.AtiendeSabado,
                target.AtiendeDomingo = source.AtiendeDomingo,
                target.HoraApertura = source.HoraApertura,
                target.HoraCierre = source.HoraCierre,
                target.FechaActualizacion = SYSUTCDATETIME(),
                target.UsuarioActualizacion = source.Usuario
        WHEN NOT MATCHED THEN
            INSERT
            (
                EspacioDeportivoId, ConfigurarHorarioPorEspacio,
                AtiendeLunes, AtiendeMartes, AtiendeMiercoles, AtiendeJueves,
                AtiendeViernes, AtiendeSabado, AtiendeDomingo,
                HoraApertura, HoraCierre, FechaCreacion, UsuarioCreacion
            )
            VALUES
            (
                source.EspacioDeportivoId, source.ConfigurarHorarioPorEspacio,
                source.AtiendeLunes, source.AtiendeMartes, source.AtiendeMiercoles, source.AtiendeJueves,
                source.AtiendeViernes, source.AtiendeSabado, source.AtiendeDomingo,
                source.HoraApertura, source.HoraCierre, SYSUTCDATETIME(), source.Usuario
            );

        IF @TieneEspaciosCompartidos = 1
        BEGIN
            INSERT INTO dbo.EspaciosDeportivosCompartidos
            (
                EspacioDeportivoId, EspacioRelacionadoId, TipoRelacion, Activo, FechaCreacion, UsuarioCreacion
            )
            SELECT
                @Id,
                ec.EspacioRelacionadoId,
                N'DIRECTO',
                1,
                SYSUTCDATETIME(),
                @Usuario
            FROM @EspaciosDirectos ec;

            INSERT INTO dbo.EspaciosDeportivosCompartidos
            (
                EspacioDeportivoId, EspacioRelacionadoId, TipoRelacion, Activo, FechaCreacion, UsuarioCreacion
            )
            SELECT
                ec.EspacioRelacionadoId,
                @Id,
                N'DIRECTO',
                1,
                SYSUTCDATETIME(),
                @Usuario
            FROM @EspaciosDirectos ec;

            INSERT INTO dbo.EspaciosDeportivosCompartidos
            (
                EspacioDeportivoId, EspacioRelacionadoId, TipoRelacion, Activo, FechaCreacion, UsuarioCreacion
            )
            SELECT
                @Id,
                ec.EspacioRelacionadoId,
                N'COMPUESTO_COMPONENTE',
                1,
                SYSUTCDATETIME(),
                @Usuario
            FROM @EspaciosComponentes ec;
        END;

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
