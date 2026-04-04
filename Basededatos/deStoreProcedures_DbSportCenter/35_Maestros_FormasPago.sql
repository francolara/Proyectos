-- =============================================
-- Author:        FRANCO LARA
-- Create date:   01/04/2026
-- Description:   Modulo Maestros (Monedas, TiposSuelo, TiposDeporte, FormasPago) y pagos enlazados a catalogo de formas de pago.
-- Firma:         Codex - 01/04/2026 | Agrega tabla FormasPago, CRUD de maestros, combo de formas de pago y ajuste de SP de pagos + seed de modulo MAESTROS.
-- =============================================
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   02/04/2026
-- Description:   Completa SP faltantes de Maestros (TiposSuelo, TiposDeporte, FormasPago) para evitar error en la pestana Maestros.
-- Firma:         Codex - 02/04/2026 | Completa CRUD de catalogos Maestros y deja contrato ADO.NET + SP consistente para MaestrosController.
-- =============================================

IF OBJECT_ID(N'dbo.FormasPago', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.FormasPago
    (
        Id INT IDENTITY(1,1) NOT NULL,
        Nombre NVARCHAR(80) NOT NULL,
        Activo BIT NOT NULL CONSTRAINT DF_FormasPago_Activo DEFAULT (1),
        FechaCreacion DATETIME2 NOT NULL CONSTRAINT DF_FormasPago_FechaCreacion DEFAULT (SYSUTCDATETIME()),
        UsuarioCreacion NVARCHAR(200) NULL,
        FechaActualizacion DATETIME2 NULL,
        UsuarioActualizacion NVARCHAR(200) NULL,
        CONSTRAINT PK_FormasPago PRIMARY KEY CLUSTERED (Id),
        CONSTRAINT UQ_FormasPago_Nombre UNIQUE (Nombre)
    );
END;
GO

IF NOT EXISTS (SELECT 1 FROM dbo.FormasPago WHERE UPPER(LTRIM(RTRIM(Nombre))) = N'EFECTIVO')
    INSERT INTO dbo.FormasPago (Nombre, Activo, UsuarioCreacion) VALUES (N'Efectivo', 1, N'seed');
IF NOT EXISTS (SELECT 1 FROM dbo.FormasPago WHERE UPPER(LTRIM(RTRIM(Nombre))) = N'YAPE')
    INSERT INTO dbo.FormasPago (Nombre, Activo, UsuarioCreacion) VALUES (N'Yape', 1, N'seed');
IF NOT EXISTS (SELECT 1 FROM dbo.FormasPago WHERE UPPER(LTRIM(RTRIM(Nombre))) = N'PLIN')
    INSERT INTO dbo.FormasPago (Nombre, Activo, UsuarioCreacion) VALUES (N'Plin', 1, N'seed');
IF NOT EXISTS (SELECT 1 FROM dbo.FormasPago WHERE UPPER(LTRIM(RTRIM(Nombre))) = N'TRANSFERENCIA')
    INSERT INTO dbo.FormasPago (Nombre, Activo, UsuarioCreacion) VALUES (N'Transferencia', 1, N'seed');
IF NOT EXISTS (SELECT 1 FROM dbo.FormasPago WHERE UPPER(LTRIM(RTRIM(Nombre))) = N'TARJETA')
    INSERT INTO dbo.FormasPago (Nombre, Activo, UsuarioCreacion) VALUES (N'Tarjeta', 1, N'seed');
GO

IF EXISTS (SELECT 1 FROM dbo.Pagos p LEFT JOIN dbo.FormasPago fp ON fp.Id = p.FormaPago WHERE fp.Id IS NULL)
BEGIN
    UPDATE p
    SET p.FormaPago = 1
    FROM dbo.Pagos p
    LEFT JOIN dbo.FormasPago fp ON fp.Id = p.FormaPago
    WHERE fp.Id IS NULL;
END;
GO

IF NOT EXISTS (SELECT 1 FROM sys.foreign_keys WHERE name = N'FK_Pagos_FormasPago_FormaPago')
BEGIN
    ALTER TABLE dbo.Pagos
        ADD CONSTRAINT FK_Pagos_FormasPago_FormaPago
            FOREIGN KEY (FormaPago) REFERENCES dbo.FormasPago (Id);
END;
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Seguridad_SeedModulosPermisosBase
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF NOT EXISTS (SELECT 1 FROM dbo.ModulosSistema WHERE Codigo = N'DASHBOARD')
            INSERT INTO dbo.ModulosSistema (Codigo, Nombre, Activo) VALUES (N'DASHBOARD', N'Dashboard', 1);
        IF NOT EXISTS (SELECT 1 FROM dbo.ModulosSistema WHERE Codigo = N'MAESTROS')
            INSERT INTO dbo.ModulosSistema (Codigo, Nombre, Activo) VALUES (N'MAESTROS', N'Maestros', 1);
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

CREATE OR ALTER PROCEDURE dbo.Sp_Combos_FormasPago
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT fp.Id, fp.Nombre
        FROM dbo.FormasPago fp
        WHERE fp.Activo = 1
        ORDER BY fp.Nombre;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Pagos_Listar
    @NegocioId INT,
    @SedeId INT = NULL
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT TOP (100)
            p.Id,
            p.ReservaId,
            p.FechaPago,
            p.Monto,
            fp.Nombre AS FormaPago
        FROM dbo.Pagos p
        INNER JOIN dbo.FormasPago fp ON fp.Id = p.FormaPago
        INNER JOIN dbo.Reservas r ON r.Id = p.ReservaId
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        WHERE s.NegocioId = @NegocioId
          AND (@SedeId IS NULL OR s.Id = @SedeId)
        ORDER BY p.FechaPago DESC;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Pagos_Crear
    @NegocioId INT,
    @ReservaId INT,
    @FechaPago DATETIME2,
    @Monto DECIMAL(10,2),
    @FormaPago INT,
    @NumeroOperacion NVARCHAR(50) = NULL,
    @Observacion NVARCHAR(300) = NULL,
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF @Monto <= 0
            RAISERROR('El monto debe ser mayor que cero.', 16, 1);

        IF NOT EXISTS (SELECT 1 FROM dbo.FormasPago WHERE Id = @FormaPago AND Activo = 1)
            RAISERROR('La forma de pago no es valida.', 16, 1);

        DECLARE @TotalReserva DECIMAL(10,2);
        DECLARE @PagadoActual DECIMAL(10,2);
        DECLARE @NuevoPagado DECIMAL(10,2);

        SELECT @TotalReserva = r.Total
        FROM dbo.Reservas r
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        WHERE r.Id = @ReservaId
          AND s.NegocioId = @NegocioId;

        IF @TotalReserva IS NULL
            RAISERROR('Reserva invalida para el negocio.', 16, 1);

        SELECT @PagadoActual = COALESCE(SUM(p.Monto), 0)
        FROM dbo.Pagos p
        WHERE p.ReservaId = @ReservaId;

        SET @NuevoPagado = @PagadoActual + @Monto;
        IF @NuevoPagado > @TotalReserva
            RAISERROR('El pago excede el total de la reserva.', 16, 1);

        BEGIN TRANSACTION;

        INSERT INTO dbo.Pagos
        (
            ReservaId, FechaPago, Monto, FormaPago, NumeroOperacion, Observacion,
            FechaCreacion, UsuarioCreacion
        )
        VALUES
        (
            @ReservaId, @FechaPago, @Monto, @FormaPago, @NumeroOperacion, @Observacion,
            SYSUTCDATETIME(), @Usuario
        );

        UPDATE r
        SET Adelanto = @NuevoPagado,
            Saldo = (r.Total - @NuevoPagado),
            Estado = CASE WHEN (r.Total - @NuevoPagado) <= 0 AND r.Estado = 1 THEN 2 ELSE r.Estado END,
            FechaActualizacion = SYSUTCDATETIME(),
            UsuarioActualizacion = @Usuario
        FROM dbo.Reservas r
        WHERE r.Id = @ReservaId;

        DECLARE @Id INT;
        SET @Id = SCOPE_IDENTITY();
        DECLARE @EntidadIdAudit NVARCHAR(80);
        SET @EntidadIdAudit = CONVERT(NVARCHAR(80), @Id);
        EXEC dbo.Sp_Auditoria_Registrar @NegocioId = @NegocioId, @Modulo = N'PAGOS', @Accion = N'CREATE', @Entidad = N'Pago', @EntidadId = @EntidadIdAudit, @Usuario = @Usuario, @DetalleJson = NULL;

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

CREATE OR ALTER PROCEDURE dbo.Sp_Pagos_Actualizar
    @Id INT,
    @NegocioId INT,
    @ReservaId INT,
    @FechaPago DATETIME2,
    @Monto DECIMAL(10,2),
    @FormaPago INT,
    @NumeroOperacion NVARCHAR(50) = NULL,
    @Observacion NVARCHAR(300) = NULL,
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF @Monto <= 0
            RAISERROR('El monto debe ser mayor que cero.', 16, 1);

        IF NOT EXISTS (SELECT 1 FROM dbo.FormasPago WHERE Id = @FormaPago AND Activo = 1)
            RAISERROR('La forma de pago no es valida.', 16, 1);

        DECLARE @ReservaAnteriorId INT;
        SELECT @ReservaAnteriorId = p.ReservaId FROM dbo.Pagos p WHERE p.Id = @Id;

        IF @ReservaAnteriorId IS NULL
            RAISERROR('No se encontro el pago para actualizar en el negocio.', 16, 1);

        IF NOT EXISTS (
            SELECT 1
            FROM dbo.Reservas r
            INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
            INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
            WHERE r.Id = @ReservaId
              AND s.NegocioId = @NegocioId
        )
            RAISERROR('Reserva invalida para el negocio.', 16, 1);

        DECLARE @TotalReserva DECIMAL(10,2);
        DECLARE @PagadoSinEste DECIMAL(10,2);
        DECLARE @NuevoPagado DECIMAL(10,2);

        SELECT @TotalReserva = r.Total FROM dbo.Reservas r WHERE r.Id = @ReservaId;

        SELECT @PagadoSinEste = COALESCE(SUM(p.Monto), 0)
        FROM dbo.Pagos p
        WHERE p.ReservaId = @ReservaId
          AND p.Id <> @Id;

        SET @NuevoPagado = @PagadoSinEste + @Monto;
        IF @NuevoPagado > @TotalReserva
            RAISERROR('El pago excede el total de la reserva.', 16, 1);

        BEGIN TRANSACTION;

        UPDATE dbo.Pagos
        SET ReservaId = @ReservaId,
            FechaPago = @FechaPago,
            Monto = @Monto,
            FormaPago = @FormaPago,
            NumeroOperacion = @NumeroOperacion,
            Observacion = @Observacion,
            FechaActualizacion = SYSUTCDATETIME(),
            UsuarioActualizacion = @Usuario
        WHERE Id = @Id;

        DECLARE @ReservaRecalculo TABLE (ReservaId INT PRIMARY KEY);
        INSERT INTO @ReservaRecalculo (ReservaId) VALUES (@ReservaId);
        IF @ReservaAnteriorId <> @ReservaId
            INSERT INTO @ReservaRecalculo (ReservaId) VALUES (@ReservaAnteriorId);

        UPDATE r
        SET Adelanto = x.Pagado,
            Saldo = (r.Total - x.Pagado),
            Estado = CASE WHEN (r.Total - x.Pagado) <= 0 AND r.Estado = 1 THEN 2 ELSE r.Estado END,
            FechaActualizacion = SYSUTCDATETIME(),
            UsuarioActualizacion = @Usuario
        FROM dbo.Reservas r
        INNER JOIN (
            SELECT rr.ReservaId, COALESCE(SUM(p.Monto), 0) AS Pagado
            FROM @ReservaRecalculo rr
            LEFT JOIN dbo.Pagos p ON p.ReservaId = rr.ReservaId
            GROUP BY rr.ReservaId
        ) x ON x.ReservaId = r.Id;

        DECLARE @EntidadIdAudit NVARCHAR(80);
        SET @EntidadIdAudit = CONVERT(NVARCHAR(80), @Id);
        EXEC dbo.Sp_Auditoria_Registrar @NegocioId = @NegocioId, @Modulo = N'PAGOS', @Accion = N'EDIT', @Entidad = N'Pago', @EntidadId = @EntidadIdAudit, @Usuario = @Usuario, @DetalleJson = NULL;

        COMMIT TRANSACTION;
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

CREATE OR ALTER PROCEDURE dbo.Sp_Maestros_Monedas_Listar
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT m.Id, m.Codigo, m.Nombre, m.Simbolo, m.Activo
        FROM dbo.Monedas m
        ORDER BY m.Nombre;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Maestros_Monedas_Crear
    @Codigo NVARCHAR(10),
    @Nombre NVARCHAR(80),
    @Simbolo NVARCHAR(10) = NULL,
    @Activo BIT,
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SET @Codigo = UPPER(LTRIM(RTRIM(@Codigo)));
        SET @Nombre = LTRIM(RTRIM(@Nombre));

        IF @Codigo = N'' OR @Nombre = N''
            RAISERROR('Codigo y nombre son obligatorios.', 16, 1);
        IF EXISTS (SELECT 1 FROM dbo.Monedas WHERE UPPER(LTRIM(RTRIM(Codigo))) = @Codigo)
            RAISERROR('Ya existe una moneda con ese codigo.', 16, 1);
        IF EXISTS (SELECT 1 FROM dbo.Monedas WHERE UPPER(LTRIM(RTRIM(Nombre))) = UPPER(@Nombre))
            RAISERROR('Ya existe una moneda con ese nombre.', 16, 1);

        INSERT INTO dbo.Monedas (Codigo, Nombre, Simbolo, Activo)
        VALUES (@Codigo, @Nombre, NULLIF(LTRIM(RTRIM(@Simbolo)), N''), @Activo);

        DECLARE @Id INT = SCOPE_IDENTITY();
        DECLARE @EntidadIdAudit NVARCHAR(80) = CONVERT(NVARCHAR(80), @Id);
        EXEC dbo.Sp_Auditoria_Registrar @NegocioId = NULL, @Modulo = N'MAESTROS', @Accion = N'CREATE', @Entidad = N'Moneda', @EntidadId = @EntidadIdAudit, @Usuario = @Usuario, @DetalleJson = NULL;
        SELECT @Id;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Maestros_Monedas_Actualizar
    @Id INT,
    @Codigo NVARCHAR(10),
    @Nombre NVARCHAR(80),
    @Simbolo NVARCHAR(10) = NULL,
    @Activo BIT,
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SET @Codigo = UPPER(LTRIM(RTRIM(@Codigo)));
        SET @Nombre = LTRIM(RTRIM(@Nombre));

        IF @Codigo = N'' OR @Nombre = N''
            RAISERROR('Codigo y nombre son obligatorios.', 16, 1);
        IF EXISTS (SELECT 1 FROM dbo.Monedas WHERE UPPER(LTRIM(RTRIM(Codigo))) = @Codigo AND Id <> @Id)
            RAISERROR('Ya existe una moneda con ese codigo.', 16, 1);
        IF EXISTS (SELECT 1 FROM dbo.Monedas WHERE UPPER(LTRIM(RTRIM(Nombre))) = UPPER(@Nombre) AND Id <> @Id)
            RAISERROR('Ya existe una moneda con ese nombre.', 16, 1);
        IF @Activo = 0 AND EXISTS (SELECT 1 FROM dbo.Negocios WHERE MonedaId = @Id AND Activo = 1)
            RAISERROR('No se puede inactivar la moneda porque esta en uso por un negocio activo.', 16, 1);

        UPDATE dbo.Monedas
        SET Codigo = @Codigo,
            Nombre = @Nombre,
            Simbolo = NULLIF(LTRIM(RTRIM(@Simbolo)), N''),
            Activo = @Activo
        WHERE Id = @Id;

        IF @@ROWCOUNT = 0
            RAISERROR('No se encontro la moneda para actualizar.', 16, 1);

        DECLARE @EntidadIdAudit NVARCHAR(80) = CONVERT(NVARCHAR(80), @Id);
        EXEC dbo.Sp_Auditoria_Registrar @NegocioId = NULL, @Modulo = N'MAESTROS', @Accion = N'EDIT', @Entidad = N'Moneda', @EntidadId = @EntidadIdAudit, @Usuario = @Usuario, @DetalleJson = NULL;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Maestros_Monedas_Eliminar
    @Id INT,
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF EXISTS (SELECT 1 FROM dbo.Negocios WHERE MonedaId = @Id AND Activo = 1)
            RAISERROR('No se puede inactivar la moneda porque esta en uso por un negocio activo.', 16, 1);

        UPDATE dbo.Monedas
        SET Activo = 0
        WHERE Id = @Id;

        IF @@ROWCOUNT = 0
            RAISERROR('No se encontro la moneda para inactivar.', 16, 1);

        DECLARE @EntidadIdAudit NVARCHAR(80) = CONVERT(NVARCHAR(80), @Id);
        EXEC dbo.Sp_Auditoria_Registrar @NegocioId = NULL, @Modulo = N'MAESTROS', @Accion = N'DELETE', @Entidad = N'Moneda', @EntidadId = @EntidadIdAudit, @Usuario = @Usuario, @DetalleJson = NULL;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Maestros_TiposSuelo_Listar
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT ts.Id, ts.Nombre, ts.Activo
        FROM dbo.TiposSuelo ts
        ORDER BY ts.Nombre;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Maestros_TiposSuelo_Crear
    @Nombre NVARCHAR(80),
    @Activo BIT,
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SET @Nombre = LTRIM(RTRIM(@Nombre));
        IF @Nombre = N''
            RAISERROR('El nombre es obligatorio.', 16, 1);
        IF EXISTS (SELECT 1 FROM dbo.TiposSuelo WHERE UPPER(LTRIM(RTRIM(Nombre))) = UPPER(@Nombre))
            RAISERROR('Ya existe un tipo de suelo con ese nombre.', 16, 1);

        INSERT INTO dbo.TiposSuelo (Nombre, Activo)
        VALUES (@Nombre, @Activo);

        DECLARE @Id INT = SCOPE_IDENTITY();
        DECLARE @EntidadIdAudit NVARCHAR(80) = CONVERT(NVARCHAR(80), @Id);
        EXEC dbo.Sp_Auditoria_Registrar @NegocioId = NULL, @Modulo = N'MAESTROS', @Accion = N'CREATE', @Entidad = N'TipoSuelo', @EntidadId = @EntidadIdAudit, @Usuario = @Usuario, @DetalleJson = NULL;
        SELECT @Id;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Maestros_TiposSuelo_Actualizar
    @Id INT,
    @Nombre NVARCHAR(80),
    @Activo BIT,
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SET @Nombre = LTRIM(RTRIM(@Nombre));
        IF @Nombre = N''
            RAISERROR('El nombre es obligatorio.', 16, 1);
        IF EXISTS (SELECT 1 FROM dbo.TiposSuelo WHERE UPPER(LTRIM(RTRIM(Nombre))) = UPPER(@Nombre) AND Id <> @Id)
            RAISERROR('Ya existe un tipo de suelo con ese nombre.', 16, 1);
        IF @Activo = 0 AND EXISTS (SELECT 1 FROM dbo.EspaciosDeportivos WHERE TipoSueloId = @Id AND Estado = 1)
            RAISERROR('No se puede inactivar el tipo de suelo porque esta en uso por un espacio activo.', 16, 1);

        UPDATE dbo.TiposSuelo
        SET Nombre = @Nombre,
            Activo = @Activo
        WHERE Id = @Id;

        IF @@ROWCOUNT = 0
            RAISERROR('No se encontro el tipo de suelo para actualizar.', 16, 1);

        DECLARE @EntidadIdAudit NVARCHAR(80) = CONVERT(NVARCHAR(80), @Id);
        EXEC dbo.Sp_Auditoria_Registrar @NegocioId = NULL, @Modulo = N'MAESTROS', @Accion = N'EDIT', @Entidad = N'TipoSuelo', @EntidadId = @EntidadIdAudit, @Usuario = @Usuario, @DetalleJson = NULL;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Maestros_TiposSuelo_Eliminar
    @Id INT,
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF EXISTS (SELECT 1 FROM dbo.EspaciosDeportivos WHERE TipoSueloId = @Id AND Estado = 1)
            RAISERROR('No se puede inactivar el tipo de suelo porque esta en uso por un espacio activo.', 16, 1);

        UPDATE dbo.TiposSuelo
        SET Activo = 0
        WHERE Id = @Id;

        IF @@ROWCOUNT = 0
            RAISERROR('No se encontro el tipo de suelo para inactivar.', 16, 1);

        DECLARE @EntidadIdAudit NVARCHAR(80) = CONVERT(NVARCHAR(80), @Id);
        EXEC dbo.Sp_Auditoria_Registrar @NegocioId = NULL, @Modulo = N'MAESTROS', @Accion = N'DELETE', @Entidad = N'TipoSuelo', @EntidadId = @EntidadIdAudit, @Usuario = @Usuario, @DetalleJson = NULL;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Maestros_TiposDeporte_Listar
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT td.Id, td.Nombre, td.Activo
        FROM dbo.TiposDeporte td
        ORDER BY td.Nombre;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Maestros_TiposDeporte_Crear
    @Nombre NVARCHAR(80),
    @Activo BIT,
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SET @Nombre = LTRIM(RTRIM(@Nombre));
        IF @Nombre = N''
            RAISERROR('El nombre es obligatorio.', 16, 1);
        IF EXISTS (SELECT 1 FROM dbo.TiposDeporte WHERE UPPER(LTRIM(RTRIM(Nombre))) = UPPER(@Nombre))
            RAISERROR('Ya existe un tipo de deporte con ese nombre.', 16, 1);

        INSERT INTO dbo.TiposDeporte (Nombre, Activo)
        VALUES (@Nombre, @Activo);

        DECLARE @Id INT = SCOPE_IDENTITY();
        DECLARE @EntidadIdAudit NVARCHAR(80) = CONVERT(NVARCHAR(80), @Id);
        EXEC dbo.Sp_Auditoria_Registrar @NegocioId = NULL, @Modulo = N'MAESTROS', @Accion = N'CREATE', @Entidad = N'TipoDeporte', @EntidadId = @EntidadIdAudit, @Usuario = @Usuario, @DetalleJson = NULL;
        SELECT @Id;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Maestros_TiposDeporte_Actualizar
    @Id INT,
    @Nombre NVARCHAR(80),
    @Activo BIT,
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SET @Nombre = LTRIM(RTRIM(@Nombre));
        IF @Nombre = N''
            RAISERROR('El nombre es obligatorio.', 16, 1);
        IF EXISTS (SELECT 1 FROM dbo.TiposDeporte WHERE UPPER(LTRIM(RTRIM(Nombre))) = UPPER(@Nombre) AND Id <> @Id)
            RAISERROR('Ya existe un tipo de deporte con ese nombre.', 16, 1);
        IF @Activo = 0 AND EXISTS (SELECT 1 FROM dbo.EspaciosDeportivos WHERE TipoDeporteId = @Id AND Estado = 1)
            RAISERROR('No se puede inactivar el tipo de deporte porque esta en uso por un espacio activo.', 16, 1);

        UPDATE dbo.TiposDeporte
        SET Nombre = @Nombre,
            Activo = @Activo
        WHERE Id = @Id;

        IF @@ROWCOUNT = 0
            RAISERROR('No se encontro el tipo de deporte para actualizar.', 16, 1);

        DECLARE @EntidadIdAudit NVARCHAR(80) = CONVERT(NVARCHAR(80), @Id);
        EXEC dbo.Sp_Auditoria_Registrar @NegocioId = NULL, @Modulo = N'MAESTROS', @Accion = N'EDIT', @Entidad = N'TipoDeporte', @EntidadId = @EntidadIdAudit, @Usuario = @Usuario, @DetalleJson = NULL;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Maestros_TiposDeporte_Eliminar
    @Id INT,
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF EXISTS (SELECT 1 FROM dbo.EspaciosDeportivos WHERE TipoDeporteId = @Id AND Estado = 1)
            RAISERROR('No se puede inactivar el tipo de deporte porque esta en uso por un espacio activo.', 16, 1);

        UPDATE dbo.TiposDeporte
        SET Activo = 0
        WHERE Id = @Id;

        IF @@ROWCOUNT = 0
            RAISERROR('No se encontro el tipo de deporte para inactivar.', 16, 1);

        DECLARE @EntidadIdAudit NVARCHAR(80) = CONVERT(NVARCHAR(80), @Id);
        EXEC dbo.Sp_Auditoria_Registrar @NegocioId = NULL, @Modulo = N'MAESTROS', @Accion = N'DELETE', @Entidad = N'TipoDeporte', @EntidadId = @EntidadIdAudit, @Usuario = @Usuario, @DetalleJson = NULL;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Maestros_FormasPago_Listar
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT fp.Id, fp.Nombre, fp.Activo
        FROM dbo.FormasPago fp
        ORDER BY fp.Nombre;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Maestros_FormasPago_Crear
    @Nombre NVARCHAR(80),
    @Activo BIT,
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SET @Nombre = LTRIM(RTRIM(@Nombre));
        IF @Nombre = N''
            RAISERROR('El nombre es obligatorio.', 16, 1);
        IF EXISTS (SELECT 1 FROM dbo.FormasPago WHERE UPPER(LTRIM(RTRIM(Nombre))) = UPPER(@Nombre))
            RAISERROR('Ya existe una forma de pago con ese nombre.', 16, 1);

        INSERT INTO dbo.FormasPago (Nombre, Activo, UsuarioCreacion)
        VALUES (@Nombre, @Activo, @Usuario);

        DECLARE @Id INT = SCOPE_IDENTITY();
        DECLARE @EntidadIdAudit NVARCHAR(80) = CONVERT(NVARCHAR(80), @Id);
        EXEC dbo.Sp_Auditoria_Registrar @NegocioId = NULL, @Modulo = N'MAESTROS', @Accion = N'CREATE', @Entidad = N'FormaPago', @EntidadId = @EntidadIdAudit, @Usuario = @Usuario, @DetalleJson = NULL;
        SELECT @Id;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Maestros_FormasPago_Actualizar
    @Id INT,
    @Nombre NVARCHAR(80),
    @Activo BIT,
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SET @Nombre = LTRIM(RTRIM(@Nombre));
        IF @Nombre = N''
            RAISERROR('El nombre es obligatorio.', 16, 1);
        IF EXISTS (SELECT 1 FROM dbo.FormasPago WHERE UPPER(LTRIM(RTRIM(Nombre))) = UPPER(@Nombre) AND Id <> @Id)
            RAISERROR('Ya existe una forma de pago con ese nombre.', 16, 1);
        IF @Activo = 0 AND EXISTS (SELECT 1 FROM dbo.Pagos WHERE FormaPago = @Id)
            RAISERROR('No se puede inactivar la forma de pago porque esta en uso por pagos registrados.', 16, 1);

        UPDATE dbo.FormasPago
        SET Nombre = @Nombre,
            Activo = @Activo,
            FechaActualizacion = SYSUTCDATETIME(),
            UsuarioActualizacion = @Usuario
        WHERE Id = @Id;

        IF @@ROWCOUNT = 0
            RAISERROR('No se encontro la forma de pago para actualizar.', 16, 1);

        DECLARE @EntidadIdAudit NVARCHAR(80) = CONVERT(NVARCHAR(80), @Id);
        EXEC dbo.Sp_Auditoria_Registrar @NegocioId = NULL, @Modulo = N'MAESTROS', @Accion = N'EDIT', @Entidad = N'FormaPago', @EntidadId = @EntidadIdAudit, @Usuario = @Usuario, @DetalleJson = NULL;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Maestros_FormasPago_Eliminar
    @Id INT,
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF EXISTS (SELECT 1 FROM dbo.Pagos WHERE FormaPago = @Id)
            RAISERROR('No se puede inactivar la forma de pago porque esta en uso por pagos registrados.', 16, 1);

        UPDATE dbo.FormasPago
        SET Activo = 0,
            FechaActualizacion = SYSUTCDATETIME(),
            UsuarioActualizacion = @Usuario
        WHERE Id = @Id;

        IF @@ROWCOUNT = 0
            RAISERROR('No se encontro la forma de pago para inactivar.', 16, 1);

        DECLARE @EntidadIdAudit NVARCHAR(80) = CONVERT(NVARCHAR(80), @Id);
        EXEC dbo.Sp_Auditoria_Registrar @NegocioId = NULL, @Modulo = N'MAESTROS', @Accion = N'DELETE', @Entidad = N'FormaPago', @EntidadId = @EntidadIdAudit, @Usuario = @Usuario, @DetalleJson = NULL;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO
