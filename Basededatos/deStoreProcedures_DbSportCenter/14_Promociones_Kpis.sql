-- =============================================
-- Author:        FRANCO LARA
-- Create date:   27/03/2026
-- Description:   Sprint 4 - KPIs avanzados de panel y modulo de promociones horarias.
-- =============================================
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   30/03/2026
-- Description:   Ajusta update/delete de promociones para devolver error controlado cuando no existe registro para el negocio.
-- Firma:         Codex - 30/03/2026 | Elimina pre-chequeos de existencia en C# y centraliza validacion en SP.
-- =============================================

IF OBJECT_ID(N'dbo.PromocionesHorario', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.PromocionesHorario
    (
        Id INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        NegocioId INT NOT NULL,
        SedeId INT NULL,
        EspacioDeportivoId INT NULL,
        Nombre NVARCHAR(150) NOT NULL,
        FechaInicio DATE NOT NULL,
        FechaFin DATE NOT NULL,
        HoraInicio TIME NOT NULL,
        HoraFin TIME NOT NULL,
        PorcentajeDescuento DECIMAL(5,2) NOT NULL,
        Activo BIT NOT NULL CONSTRAINT DF_PromocionesHorario_Activo DEFAULT (1),
        FechaRegistro DATETIME2 NOT NULL CONSTRAINT DF_PromocionesHorario_FechaRegistro DEFAULT (SYSUTCDATETIME()),
        UsuarioCreacion NVARCHAR(200) NULL,
        FechaActualizacion DATETIME2 NULL,
        UsuarioActualizacion NVARCHAR(200) NULL,
        CONSTRAINT FK_PromocionesHorario_Negocios_NegocioId FOREIGN KEY (NegocioId) REFERENCES dbo.Negocios (Id),
        CONSTRAINT FK_PromocionesHorario_Sedes_SedeId FOREIGN KEY (SedeId) REFERENCES dbo.Sedes (Id),
        CONSTRAINT FK_PromocionesHorario_Espacios_EspacioDeportivoId FOREIGN KEY (EspacioDeportivoId) REFERENCES dbo.EspaciosDeportivos (Id)
    );
END;
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Seguridad_SeedModulosPermisosBase
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF NOT EXISTS (SELECT 1 FROM dbo.ModulosSistema WHERE Codigo = N'DASHBOARD')
            INSERT INTO dbo.ModulosSistema (Codigo, Nombre, Activo) VALUES (N'DASHBOARD', N'Dashboard', 1);
        IF NOT EXISTS (SELECT 1 FROM dbo.ModulosSistema WHERE Codigo = N'SEDES')
            INSERT INTO dbo.ModulosSistema (Codigo, Nombre, Activo) VALUES (N'SEDES', N'Sedes', 1);
        IF NOT EXISTS (SELECT 1 FROM dbo.ModulosSistema WHERE Codigo = N'CLIENTES')
            INSERT INTO dbo.ModulosSistema (Codigo, Nombre, Activo) VALUES (N'CLIENTES', N'Clientes', 1);
        IF NOT EXISTS (SELECT 1 FROM dbo.ModulosSistema WHERE Codigo = N'ESPACIOS')
            INSERT INTO dbo.ModulosSistema (Codigo, Nombre, Activo) VALUES (N'ESPACIOS', N'Espacios deportivos', 1);
        IF NOT EXISTS (SELECT 1 FROM dbo.ModulosSistema WHERE Codigo = N'RESERVAS')
            INSERT INTO dbo.ModulosSistema (Codigo, Nombre, Activo) VALUES (N'RESERVAS', N'Reservas', 1);
        IF NOT EXISTS (SELECT 1 FROM dbo.ModulosSistema WHERE Codigo = N'SOLICITUDES')
            INSERT INTO dbo.ModulosSistema (Codigo, Nombre, Activo) VALUES (N'SOLICITUDES', N'Solicitudes publicas', 1);
        IF NOT EXISTS (SELECT 1 FROM dbo.ModulosSistema WHERE Codigo = N'PROMOCIONES')
            INSERT INTO dbo.ModulosSistema (Codigo, Nombre, Activo) VALUES (N'PROMOCIONES', N'Promociones', 1);
        IF NOT EXISTS (SELECT 1 FROM dbo.ModulosSistema WHERE Codigo = N'USUARIOS')
            INSERT INTO dbo.ModulosSistema (Codigo, Nombre, Activo) VALUES (N'USUARIOS', N'Usuarios del negocio', 1);
        IF NOT EXISTS (SELECT 1 FROM dbo.ModulosSistema WHERE Codigo = N'PAGOS')
            INSERT INTO dbo.ModulosSistema (Codigo, Nombre, Activo) VALUES (N'PAGOS', N'Pagos', 1);
        IF NOT EXISTS (SELECT 1 FROM dbo.ModulosSistema WHERE Codigo = N'COMPROBANTES')
            INSERT INTO dbo.ModulosSistema (Codigo, Nombre, Activo) VALUES (N'COMPROBANTES', N'Comprobantes electronicos', 1);
        IF NOT EXISTS (SELECT 1 FROM dbo.ModulosSistema WHERE Codigo = N'REPORTES')
            INSERT INTO dbo.ModulosSistema (Codigo, Nombre, Activo) VALUES (N'REPORTES', N'Reportes', 1);

        ;WITH Roles AS
        (
            SELECT CAST(1 AS INT) AS RolNegocio UNION ALL
            SELECT 2 UNION ALL
            SELECT 3 UNION ALL
            SELECT 4 UNION ALL
            SELECT 5
        )
        INSERT INTO dbo.RolesNegocioPermiso (RolNegocio, ModuloSistemaId, PuedeVer, PuedeCrear, PuedeEditar, PuedeEliminar)
        SELECT
            r.RolNegocio,
            m.Id,
            CAST(CASE WHEN r.RolNegocio = 1 THEN 1
                      WHEN r.RolNegocio = 2 AND m.Codigo IN (N'DASHBOARD', N'RESERVAS', N'SOLICITUDES', N'PAGOS', N'CLIENTES', N'REPORTES', N'PROMOCIONES') THEN 1
                      WHEN r.RolNegocio = 3 AND m.Codigo IN (N'DASHBOARD', N'RESERVAS', N'SOLICITUDES', N'CLIENTES') THEN 1
                      WHEN r.RolNegocio = 4 AND m.Codigo IN (N'DASHBOARD', N'PAGOS', N'COMPROBANTES', N'REPORTES') THEN 1
                      WHEN r.RolNegocio = 5 AND m.Codigo IN (N'DASHBOARD', N'RESERVAS', N'ESPACIOS', N'REPORTES') THEN 1
                      ELSE 0 END AS BIT),
            CAST(CASE WHEN r.RolNegocio = 1 THEN 1
                      WHEN r.RolNegocio = 2 AND m.Codigo IN (N'RESERVAS', N'SOLICITUDES', N'PAGOS', N'CLIENTES', N'PROMOCIONES') THEN 1
                      WHEN r.RolNegocio = 3 AND m.Codigo IN (N'RESERVAS', N'SOLICITUDES', N'CLIENTES') THEN 1
                      WHEN r.RolNegocio = 4 AND m.Codigo IN (N'PAGOS', N'COMPROBANTES') THEN 1
                      ELSE 0 END AS BIT),
            CAST(CASE WHEN r.RolNegocio = 1 THEN 1
                      WHEN r.RolNegocio = 2 AND m.Codigo IN (N'RESERVAS', N'SOLICITUDES', N'PAGOS', N'CLIENTES', N'PROMOCIONES') THEN 1
                      WHEN r.RolNegocio = 3 AND m.Codigo IN (N'RESERVAS', N'SOLICITUDES', N'CLIENTES') THEN 1
                      WHEN r.RolNegocio = 4 AND m.Codigo IN (N'PAGOS', N'COMPROBANTES') THEN 1
                      WHEN r.RolNegocio = 5 AND m.Codigo IN (N'RESERVAS', N'ESPACIOS') THEN 1
                      ELSE 0 END AS BIT),
            CAST(CASE WHEN r.RolNegocio = 1 THEN 1 ELSE 0 END AS BIT)
        FROM Roles r
        CROSS JOIN dbo.ModulosSistema m
        WHERE NOT EXISTS (
            SELECT 1
            FROM dbo.RolesNegocioPermiso rp
            WHERE rp.RolNegocio = r.RolNegocio
              AND rp.ModuloSistemaId = m.Id
        );
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Panel_ObtenerMetricas
    @NegocioId INT,
    @Fecha DATE
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        DECLARE @TotalSedes INT = 0, @TotalEspacios INT = 0, @ReservasHoy INT = 0;
        DECLARE @IngresosHoy DECIMAL(12,2) = 0, @OcupacionHoyPct DECIMAL(5,2) = 0;
        DECLARE @NoShowMes INT = 0, @TicketPromedioMes DECIMAL(12,2) = 0;

        SELECT @TotalSedes = COUNT(1)
        FROM dbo.Sedes s
        WHERE s.NegocioId = @NegocioId;

        SELECT @TotalEspacios = COUNT(1)
        FROM dbo.EspaciosDeportivos e
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        WHERE s.NegocioId = @NegocioId;

        SELECT @ReservasHoy = COUNT(1)
        FROM dbo.Reservas r
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        WHERE s.NegocioId = @NegocioId
          AND r.Fecha = @Fecha;

        SELECT @IngresosHoy = COALESCE(SUM(p.Monto), 0)
        FROM dbo.Pagos p
        INNER JOIN dbo.Reservas r ON r.Id = p.ReservaId
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        WHERE s.NegocioId = @NegocioId
          AND CAST(p.FechaPago AS DATE) = @Fecha;

        IF @TotalEspacios > 0
        BEGIN
            DECLARE @EspaciosOcupadosHoy INT = 0;
            SELECT @EspaciosOcupadosHoy = COUNT(DISTINCT r.EspacioDeportivoId)
            FROM dbo.Reservas r
            INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
            INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
            WHERE s.NegocioId = @NegocioId
              AND r.Fecha = @Fecha
              AND r.Estado NOT IN (5, 6);
            SET @OcupacionHoyPct = CAST((@EspaciosOcupadosHoy * 100.0) / @TotalEspacios AS DECIMAL(5,2));
        END;

        SELECT @NoShowMes = COUNT(1)
        FROM dbo.Reservas r
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        WHERE s.NegocioId = @NegocioId
          AND r.Estado = 6
          AND YEAR(r.Fecha) = YEAR(@Fecha)
          AND MONTH(r.Fecha) = MONTH(@Fecha);

        DECLARE @TotalCobradoMes DECIMAL(12,2) = 0, @ReservasPagadasMes INT = 0;

        SELECT @TotalCobradoMes = COALESCE(SUM(p.Monto), 0)
        FROM dbo.Pagos p
        INNER JOIN dbo.Reservas r ON r.Id = p.ReservaId
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        WHERE s.NegocioId = @NegocioId
          AND YEAR(p.FechaPago) = YEAR(@Fecha)
          AND MONTH(p.FechaPago) = MONTH(@Fecha);

        SELECT @ReservasPagadasMes = COUNT(DISTINCT p.ReservaId)
        FROM dbo.Pagos p
        INNER JOIN dbo.Reservas r ON r.Id = p.ReservaId
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        WHERE s.NegocioId = @NegocioId
          AND YEAR(p.FechaPago) = YEAR(@Fecha)
          AND MONTH(p.FechaPago) = MONTH(@Fecha);

        IF @ReservasPagadasMes > 0
            SET @TicketPromedioMes = CAST(@TotalCobradoMes / @ReservasPagadasMes AS DECIMAL(12,2));

        SELECT
            @TotalSedes AS TotalSedes,
            @TotalEspacios AS TotalEspacios,
            @ReservasHoy AS ReservasHoy,
            @IngresosHoy AS IngresosHoy,
            @OcupacionHoyPct AS OcupacionHoyPct,
            @NoShowMes AS NoShowMes,
            @TicketPromedioMes AS TicketPromedioMes;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Promociones_Listar
    @NegocioId INT
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT
            p.Id,
            p.Nombre,
            COALESCE(s.Nombre, N'Todas') AS Sede,
            COALESCE(e.Nombre, N'Todos') AS Espacio,
            p.FechaInicio,
            p.FechaFin,
            p.HoraInicio,
            p.HoraFin,
            p.PorcentajeDescuento,
            p.Activo
        FROM dbo.PromocionesHorario p
        LEFT JOIN dbo.Sedes s ON s.Id = p.SedeId
        LEFT JOIN dbo.EspaciosDeportivos e ON e.Id = p.EspacioDeportivoId
        WHERE p.NegocioId = @NegocioId
        ORDER BY p.FechaInicio DESC, p.Nombre;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Promociones_ObtenerPorId
    @NegocioId INT,
    @Id INT
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT
            p.Id,
            p.SedeId,
            p.EspacioDeportivoId,
            p.Nombre,
            p.FechaInicio,
            p.FechaFin,
            p.HoraInicio,
            p.HoraFin,
            p.PorcentajeDescuento,
            p.Activo
        FROM dbo.PromocionesHorario p
        WHERE p.NegocioId = @NegocioId
          AND p.Id = @Id;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Promociones_Crear
    @NegocioId INT,
    @SedeId INT = NULL,
    @EspacioDeportivoId INT = NULL,
    @Nombre NVARCHAR(150),
    @FechaInicio DATE,
    @FechaFin DATE,
    @HoraInicio TIME,
    @HoraFin TIME,
    @PorcentajeDescuento DECIMAL(5,2),
    @Activo BIT,
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF @FechaFin < @FechaInicio
            RAISERROR('La fecha fin no puede ser menor a fecha inicio.', 16, 1);
        IF @HoraFin <= @HoraInicio
            RAISERROR('La hora fin debe ser mayor a hora inicio.', 16, 1);
        IF @PorcentajeDescuento < 0 OR @PorcentajeDescuento > 100
            RAISERROR('El descuento debe estar entre 0 y 100.', 16, 1);

        INSERT INTO dbo.PromocionesHorario
        (
            NegocioId, SedeId, EspacioDeportivoId, Nombre, FechaInicio, FechaFin,
            HoraInicio, HoraFin, PorcentajeDescuento, Activo, FechaRegistro, UsuarioCreacion
        )
        VALUES
        (
            @NegocioId, @SedeId, @EspacioDeportivoId, @Nombre, @FechaInicio, @FechaFin,
            @HoraInicio, @HoraFin, @PorcentajeDescuento, @Activo, SYSUTCDATETIME(), @Usuario
        );

        DECLARE @Id INT;
        SET @Id = SCOPE_IDENTITY();
        DECLARE @EntidadIdAudit NVARCHAR(80);
        SET @EntidadIdAudit = CONVERT(NVARCHAR(80), @Id);
        EXEC dbo.Sp_Auditoria_Registrar @NegocioId = @NegocioId, @Modulo = N'PROMOCIONES', @Accion = N'CREATE', @Entidad = N'PromocionHorario', @EntidadId = @EntidadIdAudit, @Usuario = @Usuario, @DetalleJson = NULL;
        SELECT @Id;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Promociones_Actualizar
    @Id INT,
    @NegocioId INT,
    @SedeId INT = NULL,
    @EspacioDeportivoId INT = NULL,
    @Nombre NVARCHAR(150),
    @FechaInicio DATE,
    @FechaFin DATE,
    @HoraInicio TIME,
    @HoraFin TIME,
    @PorcentajeDescuento DECIMAL(5,2),
    @Activo BIT,
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF @FechaFin < @FechaInicio
            RAISERROR('La fecha fin no puede ser menor a fecha inicio.', 16, 1);
        IF @HoraFin <= @HoraInicio
            RAISERROR('La hora fin debe ser mayor a hora inicio.', 16, 1);
        IF @PorcentajeDescuento < 0 OR @PorcentajeDescuento > 100
            RAISERROR('El descuento debe estar entre 0 y 100.', 16, 1);

        UPDATE dbo.PromocionesHorario
        SET SedeId = @SedeId,
            EspacioDeportivoId = @EspacioDeportivoId,
            Nombre = @Nombre,
            FechaInicio = @FechaInicio,
            FechaFin = @FechaFin,
            HoraInicio = @HoraInicio,
            HoraFin = @HoraFin,
            PorcentajeDescuento = @PorcentajeDescuento,
            Activo = @Activo,
            FechaActualizacion = SYSUTCDATETIME(),
            UsuarioActualizacion = @Usuario
        WHERE Id = @Id
          AND NegocioId = @NegocioId;

        IF @@ROWCOUNT = 0
            RAISERROR('No se encontro la promocion para actualizar en el negocio.', 16, 1);

        DECLARE @EntidadIdAudit NVARCHAR(80);
        SET @EntidadIdAudit = CONVERT(NVARCHAR(80), @Id);
        EXEC dbo.Sp_Auditoria_Registrar @NegocioId = @NegocioId, @Modulo = N'PROMOCIONES', @Accion = N'EDIT', @Entidad = N'PromocionHorario', @EntidadId = @EntidadIdAudit, @Usuario = @Usuario, @DetalleJson = NULL;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Promociones_Eliminar
    @NegocioId INT,
    @Id INT,
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        UPDATE dbo.PromocionesHorario
        SET Activo = 0,
            FechaActualizacion = SYSUTCDATETIME(),
            UsuarioActualizacion = @Usuario
        WHERE Id = @Id
          AND NegocioId = @NegocioId;

        IF @@ROWCOUNT = 0
            RAISERROR('No se encontro la promocion para eliminar en el negocio.', 16, 1);

        DECLARE @EntidadIdAudit NVARCHAR(80);
        SET @EntidadIdAudit = CONVERT(NVARCHAR(80), @Id);
        EXEC dbo.Sp_Auditoria_Registrar @NegocioId = @NegocioId, @Modulo = N'PROMOCIONES', @Accion = N'DELETE', @Entidad = N'PromocionHorario', @EntidadId = @EntidadIdAudit, @Usuario = @Usuario, @DetalleJson = NULL;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO
