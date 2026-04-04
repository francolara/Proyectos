-- =============================================
-- Author:        FRANCO LARA
-- Create date:   01/04/2026
-- Description:   Agrega NombreEquipo en clientes y lo propaga a combos/listados/calendario de reservas.
-- Firma:         Codex - 01/04/2026 | Campo NombreEquipo en clientes y visualizacion en reservas (listado, combos, calendario y detalle).
-- =============================================

IF COL_LENGTH('dbo.Clientes', 'NombreEquipo') IS NULL
BEGIN
    ALTER TABLE dbo.Clientes ADD NombreEquipo NVARCHAR(120) NULL;
END;
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Combos_Clientes
    @NegocioId INT
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT
            c.Id,
            CONCAT(
                c.NombresORazonSocial,
                N' (',
                c.NumeroDocumento,
                N')',
                CASE
                    WHEN NULLIF(LTRIM(RTRIM(c.NombreEquipo)), N'') IS NULL THEN N''
                    ELSE CONCAT(N' - Equipo: ', LTRIM(RTRIM(c.NombreEquipo)))
                END
            ) AS NombreCliente
        FROM dbo.Clientes c
        INNER JOIN dbo.NegocioClientes nc ON nc.ClienteId = c.Id
        WHERE nc.NegocioId = @NegocioId
          AND nc.Activo = 1
          AND c.Activo = 1
        ORDER BY c.NombresORazonSocial;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Clientes_Listar
    @NegocioId INT
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT
            c.Id,
            c.NombresORazonSocial,
            c.NombreEquipo,
            c.TipoDocumento,
            c.NumeroDocumento,
            c.Telefono,
            c.Correo,
            c.Activo
        FROM dbo.Clientes c
        INNER JOIN dbo.NegocioClientes nc ON nc.ClienteId = c.Id
        WHERE nc.NegocioId = @NegocioId
          AND nc.Activo = 1
        ORDER BY c.NombresORazonSocial;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Clientes_ObtenerPorId
    @NegocioId INT,
    @Id INT
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT
            c.Id,
            c.NombresORazonSocial,
            c.NombreEquipo,
            c.TipoDocumento,
            c.NumeroDocumento,
            c.Telefono,
            c.Correo,
            c.DireccionFiscal,
            c.Activo
        FROM dbo.Clientes c
        INNER JOIN dbo.NegocioClientes nc ON nc.ClienteId = c.Id
        WHERE nc.NegocioId = @NegocioId
          AND nc.Activo = 1
          AND c.Id = @Id;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Clientes_Crear
    @NegocioId INT,
    @NombresORazonSocial NVARCHAR(200),
    @NombreEquipo NVARCHAR(120) = NULL,
    @TipoDocumento NVARCHAR(20),
    @NumeroDocumento NVARCHAR(20),
    @Telefono NVARCHAR(20) = NULL,
    @Correo NVARCHAR(200) = NULL,
    @DireccionFiscal NVARCHAR(250) = NULL,
    @Activo BIT,
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        DECLARE @NumeroDocumentoNormalizado NVARCHAR(20);
        DECLARE @NombreEquipoNormalizado NVARCHAR(120);
        SET @NumeroDocumentoNormalizado = NULLIF(LTRIM(RTRIM(@NumeroDocumento)), N'');
        SET @NombreEquipoNormalizado = NULLIF(LTRIM(RTRIM(@NombreEquipo)), N'');
        SET @NumeroDocumento = COALESCE(@NumeroDocumentoNormalizado, N'');

        IF @NumeroDocumentoNormalizado IS NOT NULL
           AND EXISTS
           (
               SELECT 1
               FROM dbo.Clientes c
               INNER JOIN dbo.NegocioClientes nc ON nc.ClienteId = c.Id
               WHERE nc.NegocioId = @NegocioId
                 AND nc.Activo = 1
                 AND c.Activo = 1
                 AND LTRIM(RTRIM(c.NumeroDocumento)) = @NumeroDocumentoNormalizado
           )
            RAISERROR('Cliente ya se encuentra registrado.', 16, 1);

        BEGIN TRANSACTION;

        INSERT INTO dbo.Clientes
        (
            NombresORazonSocial, NombreEquipo, TipoDocumento, NumeroDocumento, Telefono,
            Correo, DireccionFiscal, Activo, FechaCreacion, UsuarioCreacion
        )
        VALUES
        (
            @NombresORazonSocial, @NombreEquipoNormalizado, @TipoDocumento, @NumeroDocumento, @Telefono,
            @Correo, @DireccionFiscal, @Activo, SYSUTCDATETIME(), @Usuario
        );

        DECLARE @Id INT;
        SET @Id = SCOPE_IDENTITY();

        INSERT INTO dbo.NegocioClientes (NegocioId, ClienteId, Activo, FechaRegistro, UsuarioCreacion)
        VALUES (@NegocioId, @Id, 1, SYSUTCDATETIME(), @Usuario);

        DECLARE @EntidadIdAudit NVARCHAR(80);
        SET @EntidadIdAudit = CONVERT(NVARCHAR(80), @Id);
        EXEC dbo.Sp_Auditoria_Registrar @NegocioId = @NegocioId, @Modulo = N'CLIENTES', @Accion = N'CREATE', @Entidad = N'Cliente', @EntidadId = @EntidadIdAudit, @Usuario = @Usuario, @DetalleJson = NULL;

        COMMIT TRANSACTION;

        SELECT @Id;
    END TRY
    BEGIN CATCH
        IF XACT_STATE() <> 0
            ROLLBACK TRANSACTION;

        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Clientes_Actualizar
    @Id INT,
    @NegocioId INT,
    @NombresORazonSocial NVARCHAR(200),
    @NombreEquipo NVARCHAR(120) = NULL,
    @TipoDocumento NVARCHAR(20),
    @NumeroDocumento NVARCHAR(20),
    @Telefono NVARCHAR(20) = NULL,
    @Correo NVARCHAR(200) = NULL,
    @DireccionFiscal NVARCHAR(250) = NULL,
    @Activo BIT,
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        DECLARE @NumeroDocumentoNormalizado NVARCHAR(20);
        DECLARE @NombreEquipoNormalizado NVARCHAR(120);
        SET @NumeroDocumentoNormalizado = NULLIF(LTRIM(RTRIM(@NumeroDocumento)), N'');
        SET @NombreEquipoNormalizado = NULLIF(LTRIM(RTRIM(@NombreEquipo)), N'');
        SET @NumeroDocumento = COALESCE(@NumeroDocumentoNormalizado, N'');

        IF @NumeroDocumentoNormalizado IS NOT NULL
           AND EXISTS
           (
               SELECT 1
               FROM dbo.Clientes c
               INNER JOIN dbo.NegocioClientes nc ON nc.ClienteId = c.Id
               WHERE nc.NegocioId = @NegocioId
                 AND nc.Activo = 1
                 AND c.Activo = 1
                 AND c.Id <> @Id
                 AND LTRIM(RTRIM(c.NumeroDocumento)) = @NumeroDocumentoNormalizado
           )
            RAISERROR('Cliente ya se encuentra registrado.', 16, 1);

        UPDATE c
        SET
            c.NombresORazonSocial = @NombresORazonSocial,
            c.NombreEquipo = @NombreEquipoNormalizado,
            c.TipoDocumento = @TipoDocumento,
            c.NumeroDocumento = @NumeroDocumento,
            c.Telefono = @Telefono,
            c.Correo = @Correo,
            c.DireccionFiscal = @DireccionFiscal,
            c.Activo = @Activo,
            c.FechaActualizacion = SYSUTCDATETIME(),
            c.UsuarioActualizacion = @Usuario
        FROM dbo.Clientes c
        INNER JOIN dbo.NegocioClientes nc ON nc.ClienteId = c.Id
        WHERE c.Id = @Id
          AND nc.NegocioId = @NegocioId
          AND nc.Activo = 1;

        IF @@ROWCOUNT = 0
            RAISERROR('No se encontro el cliente para actualizar en el negocio.', 16, 1);

        DECLARE @EntidadIdAudit NVARCHAR(80);
        SET @EntidadIdAudit = CONVERT(NVARCHAR(80), @Id);
        EXEC dbo.Sp_Auditoria_Registrar @NegocioId = @NegocioId, @Modulo = N'CLIENTES', @Accion = N'EDIT', @Entidad = N'Cliente', @EntidadId = @EntidadIdAudit, @Usuario = @Usuario, @DetalleJson = NULL;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Reservas_Listar
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
              (@Estado IS NOT NULL AND r.Estado = @Estado)
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
                          WHERE TRY_CAST(estados.value AS INT) = r.Estado
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

CREATE OR ALTER PROCEDURE dbo.Sp_Reservas_CalendarioEventos
    @NegocioId INT,
    @FechaDesde DATE,
    @FechaHasta DATE,
    @SedeId INT = NULL,
    @EspacioDeportivoId INT = NULL,
    @Estado INT = NULL
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        ;WITH Fechas AS
        (
            SELECT @FechaDesde AS Fecha
            UNION ALL
            SELECT DATEADD(DAY, 1, Fecha) FROM Fechas WHERE Fecha < @FechaHasta
        )
        SELECT
            r.Id,
            CAST(N'RESERVA' AS NVARCHAR(20)) AS TipoEvento,
            CAST(
                CASE
                    WHEN NULLIF(LTRIM(RTRIM(c.NombreEquipo)), N'') IS NULL
                        THEN CONCAT(e.Nombre, N' - ', c.NombresORazonSocial)
                    ELSE CONCAT(e.Nombre, N' - ', LTRIM(RTRIM(c.NombreEquipo)), N' (', c.NombresORazonSocial, N')')
                END
                AS NVARCHAR(300)
            ) AS Titulo,
            r.Fecha,
            r.HoraInicio,
            r.HoraFin,
            r.Estado,
            CAST(
                CASE r.Estado
                    WHEN 1 THEN N'#f59f00'
                    WHEN 2 THEN N'#2f9e44'
                    WHEN 3 THEN N'#1971c2'
                    WHEN 4 THEN N'#495057'
                    WHEN 5 THEN N'#c92a2a'
                    WHEN 6 THEN N'#212529'
                    ELSE N'#6c757d'
                END
                AS NVARCHAR(20)
            ) AS Color,
            e.Id AS EspacioDeportivoId,
            e.Nombre AS Espacio,
            s.Nombre AS Sede,
            CAST(NULL AS NVARCHAR(200)) AS Motivo,
            CAST(
                CASE r.Estado
                    WHEN 1 THEN N'PENDIENTE'
                    WHEN 2 THEN N'CONFIRMADA'
                    WHEN 3 THEN N'EN_USO'
                    WHEN 4 THEN N'FINALIZADA'
                    WHEN 5 THEN N'CANCELADA'
                    WHEN 6 THEN N'NO_SHOW'
                    ELSE N'RESERVADA'
                END
                AS NVARCHAR(40)
            ) AS EstadoCodigo,
            CAST(
                CASE r.Estado
                    WHEN 1 THEN N'Pendiente'
                    WHEN 2 THEN N'Confirmada'
                    WHEN 3 THEN N'En uso'
                    WHEN 4 THEN N'Finalizada'
                    WHEN 5 THEN N'Cancelada'
                    WHEN 6 THEN N'No show'
                    ELSE N'Reservada'
                END
                AS NVARCHAR(80)
            ) AS EstadoTexto
        FROM dbo.Reservas r
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        INNER JOIN dbo.Clientes c ON c.Id = r.ClienteId
        WHERE s.NegocioId = @NegocioId
          AND r.Fecha BETWEEN @FechaDesde AND @FechaHasta
          AND (@SedeId IS NULL OR s.Id = @SedeId)
          AND (@EspacioDeportivoId IS NULL OR e.Id = @EspacioDeportivoId)
          AND (@Estado IS NULL OR r.Estado = @Estado)

        UNION ALL

        SELECT
            b.Id,
            CAST(N'BLOQUEO' AS NVARCHAR(20)) AS TipoEvento,
            CONCAT(N'Bloqueado: ', b.Motivo) AS Titulo,
            b.Fecha,
            b.HoraInicio,
            b.HoraFin,
            NULL AS Estado,
            CAST(N'#64748b' AS NVARCHAR(20)) AS Color,
            e.Id AS EspacioDeportivoId,
            e.Nombre AS Espacio,
            s.Nombre AS Sede,
            b.Motivo AS Motivo,
            CAST(N'BLOQUEADO' AS NVARCHAR(40)) AS EstadoCodigo,
            CAST(N'Bloqueado' AS NVARCHAR(80)) AS EstadoTexto
        FROM dbo.BloqueosHorario b
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = b.EspacioDeportivoId
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        WHERE s.NegocioId = @NegocioId
          AND b.Activo = 1
          AND b.Fecha BETWEEN @FechaDesde AND @FechaHasta
          AND (@SedeId IS NULL OR s.Id = @SedeId)
          AND (@EspacioDeportivoId IS NULL OR e.Id = @EspacioDeportivoId)

        UNION ALL

        SELECT
            (
                110000000
                + (DATEDIFF(DAY, '2020-01-01', sfi.Fecha) * 10000)
                + (e.Id % 10000)
            ),
            CAST(N'NO_ATENCION' AS NVARCHAR(20)),
            CAST(N'Sede sin atencion (fecha inhabilitada)' AS NVARCHAR(200)),
            sfi.Fecha,
            CAST('00:00' AS TIME),
            CAST('23:59' AS TIME),
            NULL,
            CAST(N'#64748b' AS NVARCHAR(20)),
            e.Id,
            e.Nombre,
            s.Nombre,
            CAST(N'Sede sin atencion (fecha inhabilitada)' AS NVARCHAR(200)),
            CAST(N'BLOQUEADO_NO_ATENCION' AS NVARCHAR(40)),
            CAST(N'Bloqueado/No atencion' AS NVARCHAR(80))
        FROM dbo.SedeFechasInhabilitadas sfi
        INNER JOIN dbo.Sedes s ON s.Id = sfi.SedeId
        INNER JOIN dbo.EspaciosDeportivos e ON e.SedeId = s.Id
        WHERE s.NegocioId = @NegocioId
          AND sfi.Activo = 1
          AND sfi.Fecha BETWEEN @FechaDesde AND @FechaHasta
          AND (@SedeId IS NULL OR s.Id = @SedeId)
          AND (@EspacioDeportivoId IS NULL OR e.Id = @EspacioDeportivoId)

        UNION ALL

        SELECT
            (
                120000000
                + (DATEDIFF(DAY, '2020-01-01', f.Fecha) * 10000)
                + (e.Id % 10000)
            ),
            CAST(N'NO_ATENCION' AS NVARCHAR(20)),
            CAST(N'Sede sin atencion (dia no laborable)' AS NVARCHAR(200)),
            f.Fecha,
            CAST('00:00' AS TIME),
            CAST('23:59' AS TIME),
            NULL,
            CAST(N'#64748b' AS NVARCHAR(20)),
            e.Id,
            e.Nombre,
            s.Nombre,
            CAST(N'Sede sin atencion (dia no laborable)' AS NVARCHAR(200)),
            CAST(N'BLOQUEADO_NO_ATENCION' AS NVARCHAR(40)),
            CAST(N'Bloqueado/No atencion' AS NVARCHAR(80))
        FROM Fechas f
        INNER JOIN dbo.Sedes s ON s.NegocioId = @NegocioId
        INNER JOIN dbo.EspaciosDeportivos e ON e.SedeId = s.Id
        LEFT JOIN dbo.SedeHorarioAtencion sha ON sha.SedeId = s.Id
        LEFT JOIN dbo.SedeFechasInhabilitadas sfi ON sfi.SedeId = s.Id AND sfi.Activo = 1 AND sfi.Fecha = f.Fecha
        WHERE (@SedeId IS NULL OR s.Id = @SedeId)
          AND (@EspacioDeportivoId IS NULL OR e.Id = @EspacioDeportivoId)
          AND sfi.SedeId IS NULL
          AND CASE ((DATEDIFF(DAY, '19000101', f.Fecha) % 7) + 1)
                WHEN 1 THEN COALESCE(sha.AtiendeLunes, 1)
                WHEN 2 THEN COALESCE(sha.AtiendeMartes, 1)
                WHEN 3 THEN COALESCE(sha.AtiendeMiercoles, 1)
                WHEN 4 THEN COALESCE(sha.AtiendeJueves, 1)
                WHEN 5 THEN COALESCE(sha.AtiendeViernes, 1)
                WHEN 6 THEN COALESCE(sha.AtiendeSabado, 1)
                WHEN 7 THEN COALESCE(sha.AtiendeDomingo, 1)
              END = 0

        UNION ALL

        SELECT
            (
                130000000
                + (DATEDIFF(DAY, '2020-01-01', f.Fecha) * 10000)
                + (e.Id % 10000)
            ),
            CAST(N'NO_ATENCION' AS NVARCHAR(20)),
            CAST(N'Sede sin atencion (fuera de horario)' AS NVARCHAR(200)),
            f.Fecha,
            CAST('00:00' AS TIME),
            COALESCE(sha.HoraApertura, CAST('08:00' AS TIME)),
            NULL,
            CAST(N'#64748b' AS NVARCHAR(20)),
            e.Id,
            e.Nombre,
            s.Nombre,
            CAST(N'Sede sin atencion (fuera de horario)' AS NVARCHAR(200)),
            CAST(N'BLOQUEADO_NO_ATENCION' AS NVARCHAR(40)),
            CAST(N'Bloqueado/No atencion' AS NVARCHAR(80))
        FROM Fechas f
        INNER JOIN dbo.Sedes s ON s.NegocioId = @NegocioId
        INNER JOIN dbo.EspaciosDeportivos e ON e.SedeId = s.Id
        LEFT JOIN dbo.SedeHorarioAtencion sha ON sha.SedeId = s.Id
        LEFT JOIN dbo.SedeFechasInhabilitadas sfi ON sfi.SedeId = s.Id AND sfi.Activo = 1 AND sfi.Fecha = f.Fecha
        WHERE (@SedeId IS NULL OR s.Id = @SedeId)
          AND (@EspacioDeportivoId IS NULL OR e.Id = @EspacioDeportivoId)
          AND sfi.SedeId IS NULL
          AND
          (
              COALESCE(sha.HoraApertura, CAST('08:00' AS TIME)) > CAST('00:00' AS TIME)
              OR COALESCE(sha.HoraCierre, CAST('23:00' AS TIME)) < CAST('23:59' AS TIME)
          )

        UNION ALL

        SELECT
            (
                140000000
                + (DATEDIFF(DAY, '2020-01-01', f.Fecha) * 10000)
                + (e.Id % 10000)
            ),
            CAST(N'NO_ATENCION' AS NVARCHAR(20)),
            CAST(N'Sede sin atencion (fuera de horario)' AS NVARCHAR(200)),
            f.Fecha,
            COALESCE(sha.HoraCierre, CAST('23:00' AS TIME)),
            CAST('23:59' AS TIME),
            NULL,
            CAST(N'#64748b' AS NVARCHAR(20)),
            e.Id,
            e.Nombre,
            s.Nombre,
            CAST(N'Sede sin atencion (fuera de horario)' AS NVARCHAR(200)),
            CAST(N'BLOQUEADO_NO_ATENCION' AS NVARCHAR(40)),
            CAST(N'Bloqueado/No atencion' AS NVARCHAR(80))
        FROM Fechas f
        INNER JOIN dbo.Sedes s ON s.NegocioId = @NegocioId
        INNER JOIN dbo.EspaciosDeportivos e ON e.SedeId = s.Id
        LEFT JOIN dbo.SedeHorarioAtencion sha ON sha.SedeId = s.Id
        LEFT JOIN dbo.SedeFechasInhabilitadas sfi ON sfi.SedeId = s.Id AND sfi.Activo = 1 AND sfi.Fecha = f.Fecha
        WHERE (@SedeId IS NULL OR s.Id = @SedeId)
          AND (@EspacioDeportivoId IS NULL OR e.Id = @EspacioDeportivoId)
          AND sfi.SedeId IS NULL
          AND
          (
              COALESCE(sha.HoraApertura, CAST('08:00' AS TIME)) > CAST('00:00' AS TIME)
              OR COALESCE(sha.HoraCierre, CAST('23:00' AS TIME)) < CAST('23:59' AS TIME)
          )
        ORDER BY Fecha, HoraInicio
        OPTION (MAXRECURSION 1000);
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Combos_ReservasPorNegocio
    @NegocioId INT,
    @SedeId INT = NULL
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT
            r.Id,
            CONCAT(
                N'#', r.Id, N' - ',
                c.NombresORazonSocial,
                CASE
                    WHEN NULLIF(LTRIM(RTRIM(c.NombreEquipo)), N'') IS NULL THEN N''
                    ELSE CONCAT(N' [', LTRIM(RTRIM(c.NombreEquipo)), N']')
                END
            ) AS ReservaTexto
        FROM dbo.Reservas r
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        INNER JOIN dbo.Clientes c ON c.Id = r.ClienteId
        WHERE s.NegocioId = @NegocioId
          AND (@SedeId IS NULL OR s.Id = @SedeId)
        ORDER BY r.Fecha DESC, r.HoraInicio DESC;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Reservas_ObtenerParaRecordatorio
    @NegocioId INT,
    @ReservaId INT
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT
            r.Id AS ReservaId,
            s.NegocioId,
            CAST(
                CASE
                    WHEN NULLIF(LTRIM(RTRIM(c.NombreEquipo)), N'') IS NULL THEN c.NombresORazonSocial
                    ELSE CONCAT(c.NombresORazonSocial, N' - Equipo: ', LTRIM(RTRIM(c.NombreEquipo)))
                END
                AS NVARCHAR(250)
            ) AS Cliente,
            c.Correo,
            s.Nombre AS Sede,
            e.Nombre AS Espacio,
            r.Fecha,
            r.HoraInicio,
            r.HoraFin,
            s.CorreoNotificacion,
            s.WhatsappContacto
        FROM dbo.Reservas r
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        INNER JOIN dbo.Clientes c ON c.Id = r.ClienteId
        WHERE r.Id = @ReservaId
          AND s.NegocioId = @NegocioId
          AND r.Estado = 1
          AND c.Activo = 1
          AND NULLIF(LTRIM(RTRIM(c.Correo)), N'') IS NOT NULL;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO
