-- =============================================
-- Author:        FRANCO LARA
-- Create date:   26/03/2026
-- Description:   CRUD de reservas, pagos y comprobantes electronicos.
-- =============================================
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   27/03/2026
-- Description:   Ajuste de auditoria con parametros nombrados para compatibilidad SQL Server.
-- Firma:         Codex - 30/03/2026 | Ajusta Sp_Reservas_Eliminar y operaciones de pagos/comprobantes para devolver error controlado cuando no existe registro.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.Sp_Combos_EspaciosPorNegocio
    @NegocioId INT
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT e.Id, CONCAT(s.Nombre, N' - ', e.Nombre)
        FROM dbo.EspaciosDeportivos e
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        WHERE s.NegocioId = @NegocioId
          AND e.Estado = 1
        ORDER BY s.Nombre, e.Nombre;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Combos_Clientes
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT c.Id, CONCAT(c.NombresORazonSocial, N' (', c.NumeroDocumento, N')')
        FROM dbo.Clientes c
        WHERE c.Activo = 1
        ORDER BY c.NombresORazonSocial;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Combos_ReservasPorNegocio
    @NegocioId INT
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT r.Id, CONCAT(N'#', r.Id, N' - ', c.NombresORazonSocial)
        FROM dbo.Reservas r
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        INNER JOIN dbo.Clientes c ON c.Id = r.ClienteId
        WHERE s.NegocioId = @NegocioId
        ORDER BY r.Fecha DESC, r.HoraInicio DESC;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Reservas_Listar
    @NegocioId INT
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT TOP (100)
            r.Id,
            c.NombresORazonSocial,
            e.Nombre,
            r.Fecha,
            r.HoraInicio,
            r.HoraFin,
            r.Total,
            CAST(r.Estado AS NVARCHAR(20))
        FROM dbo.Reservas r
        INNER JOIN dbo.Clientes c ON c.Id = r.ClienteId
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        WHERE s.NegocioId = @NegocioId
        ORDER BY r.Fecha DESC, r.HoraInicio DESC;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Reservas_ObtenerPorId
    @NegocioId INT,
    @Id INT
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT r.Id, r.EspacioDeportivoId, r.ClienteId, r.Fecha, r.HoraInicio, r.HoraFin, r.Total, r.Adelanto, r.Estado
        FROM dbo.Reservas r
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        WHERE r.Id = @Id
          AND s.NegocioId = @NegocioId;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Reservas_Crear
    @NegocioId INT,
    @EspacioDeportivoId INT,
    @ClienteId INT,
    @Fecha DATE,
    @HoraInicio TIME,
    @HoraFin TIME,
    @Total DECIMAL(10,2),
    @Adelanto DECIMAL(10,2),
    @Estado INT,
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF EXISTS (
            SELECT 1
            FROM dbo.Reservas r
            WHERE r.EspacioDeportivoId = @EspacioDeportivoId
              AND r.Fecha = @Fecha
              AND r.Estado NOT IN (5, 6)
              AND @HoraInicio < r.HoraFin
              AND @HoraFin > r.HoraInicio
        )
            RAISERROR('Cruce de horario detectado.', 16, 1);

        INSERT INTO dbo.Reservas
        (
            EspacioDeportivoId, ClienteId, Fecha, HoraInicio, HoraFin,
            Estado, Total, Adelanto, Saldo, FechaRegistro, UsuarioCreacion
        )
        VALUES
        (
            @EspacioDeportivoId, @ClienteId, @Fecha, @HoraInicio, @HoraFin,
            @Estado, @Total, @Adelanto, (@Total - @Adelanto), SYSUTCDATETIME(), @Usuario
        );

        DECLARE @Id INT;
        SET @Id = SCOPE_IDENTITY();
        DECLARE @EntidadIdAudit NVARCHAR(80);
            SET @EntidadIdAudit = CONVERT(NVARCHAR(80), @Id);
        EXEC dbo.Sp_Auditoria_Registrar @NegocioId = @NegocioId, @Modulo = N'RESERVAS', @Accion = N'CREATE', @Entidad = N'Reserva', @EntidadId = @EntidadIdAudit, @Usuario = @Usuario, @DetalleJson = NULL;
        SELECT @Id;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Reservas_Actualizar
    @Id INT,
    @NegocioId INT,
    @EspacioDeportivoId INT,
    @ClienteId INT,
    @Fecha DATE,
    @HoraInicio TIME,
    @HoraFin TIME,
    @Total DECIMAL(10,2),
    @Adelanto DECIMAL(10,2),
    @Estado INT,
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        UPDATE dbo.Reservas
        SET EspacioDeportivoId = @EspacioDeportivoId,
            ClienteId = @ClienteId,
            Fecha = @Fecha,
            HoraInicio = @HoraInicio,
            HoraFin = @HoraFin,
            Total = @Total,
            Adelanto = @Adelanto,
            Saldo = (@Total - @Adelanto),
            Estado = @Estado,
            FechaActualizacion = SYSUTCDATETIME(),
            UsuarioActualizacion = @Usuario
        WHERE Id = @Id;

        IF @@ROWCOUNT > 0
        BEGIN
            DECLARE @EntidadIdAudit NVARCHAR(80);
            SET @EntidadIdAudit = CONVERT(NVARCHAR(80), @Id);
            EXEC dbo.Sp_Auditoria_Registrar @NegocioId = @NegocioId, @Modulo = N'RESERVAS', @Accion = N'EDIT', @Entidad = N'Reserva', @EntidadId = @EntidadIdAudit, @Usuario = @Usuario, @DetalleJson = NULL;
        END
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Reservas_Eliminar
    @NegocioId INT,
    @Id INT,
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        UPDATE r
        SET r.Estado = 5,
            r.FechaActualizacion = SYSUTCDATETIME(),
            r.UsuarioActualizacion = @Usuario
        FROM dbo.Reservas r
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        WHERE r.Id = @Id
          AND s.NegocioId = @NegocioId;

        IF @@ROWCOUNT = 0
            RAISERROR('No se encontro la reserva para eliminar.', 16, 1);

        IF @@ROWCOUNT > 0
        BEGIN
            DECLARE @EntidadIdAudit NVARCHAR(80);
            SET @EntidadIdAudit = CONVERT(NVARCHAR(80), @Id);
            EXEC dbo.Sp_Auditoria_Registrar @NegocioId = @NegocioId, @Modulo = N'RESERVAS', @Accion = N'DELETE', @Entidad = N'Reserva', @EntidadId = @EntidadIdAudit, @Usuario = @Usuario, @DetalleJson = NULL;
        END
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Pagos_Listar
    @NegocioId INT
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT TOP (100)
            p.Id,
            p.ReservaId,
            p.FechaPago,
            p.Monto,
            CAST(p.FormaPago AS NVARCHAR(20))
        FROM dbo.Pagos p
        INNER JOIN dbo.Reservas r ON r.Id = p.ReservaId
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        WHERE s.NegocioId = @NegocioId
        ORDER BY p.FechaPago DESC;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Pagos_ObtenerPorId
    @NegocioId INT,
    @Id INT
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT p.Id, p.ReservaId, p.FechaPago, p.Monto, p.FormaPago, p.NumeroOperacion, p.Observacion
        FROM dbo.Pagos p
        INNER JOIN dbo.Reservas r ON r.Id = p.ReservaId
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        WHERE p.Id = @Id
          AND s.NegocioId = @NegocioId;
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
        SET Adelanto = ISNULL((SELECT SUM(p2.Monto) FROM dbo.Pagos p2 WHERE p2.ReservaId = r.Id), 0),
            Saldo = r.Total - ISNULL((SELECT SUM(p2.Monto) FROM dbo.Pagos p2 WHERE p2.ReservaId = r.Id), 0)
        FROM dbo.Reservas r
        WHERE r.Id = @ReservaId;

        DECLARE @Id INT;
        SET @Id = SCOPE_IDENTITY();
        DECLARE @EntidadIdAudit NVARCHAR(80);
            SET @EntidadIdAudit = CONVERT(NVARCHAR(80), @Id);
        EXEC dbo.Sp_Auditoria_Registrar @NegocioId = @NegocioId, @Modulo = N'PAGOS', @Accion = N'CREATE', @Entidad = N'Pago', @EntidadId = @EntidadIdAudit, @Usuario = @Usuario, @DetalleJson = NULL;
        SELECT @Id;
    END TRY
    BEGIN CATCH
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

        IF @@ROWCOUNT = 0
            RAISERROR('No se encontro el pago para actualizar en el negocio.', 16, 1);

        DECLARE @EntidadIdAudit NVARCHAR(80);
        SET @EntidadIdAudit = CONVERT(NVARCHAR(80), @Id);
        EXEC dbo.Sp_Auditoria_Registrar @NegocioId = @NegocioId, @Modulo = N'PAGOS', @Accion = N'EDIT', @Entidad = N'Pago', @EntidadId = @EntidadIdAudit, @Usuario = @Usuario, @DetalleJson = NULL;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Pagos_Eliminar
    @NegocioId INT,
    @Id INT,
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        DELETE FROM dbo.Pagos WHERE Id = @Id;
        IF @@ROWCOUNT = 0
            RAISERROR('No se encontro el pago para eliminar en el negocio.', 16, 1);

        DECLARE @EntidadIdAudit NVARCHAR(80);
        SET @EntidadIdAudit = CONVERT(NVARCHAR(80), @Id);
        EXEC dbo.Sp_Auditoria_Registrar @NegocioId = @NegocioId, @Modulo = N'PAGOS', @Accion = N'DELETE', @Entidad = N'Pago', @EntidadId = @EntidadIdAudit, @Usuario = @Usuario, @DetalleJson = NULL;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Comprobantes_Listar
    @NegocioId INT
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT TOP (100)
            c.Id,
            CAST(c.TipoComprobante AS NVARCHAR(20)),
            CONCAT(c.Serie, N'-', c.Numero),
            c.FechaEmision,
            cl.NombresORazonSocial,
            c.Total,
            CAST(c.Estado AS NVARCHAR(20))
        FROM dbo.ComprobantesElectronicos c
        INNER JOIN dbo.Clientes cl ON cl.Id = c.ClienteId
        WHERE c.NegocioId = @NegocioId
        ORDER BY c.FechaEmision DESC;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Comprobantes_ObtenerPorId
    @NegocioId INT,
    @Id INT
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT c.Id, c.ReservaId, c.TipoComprobante, c.Serie, c.Numero, c.FechaEmision, c.TipoMoneda, c.SubTotal, c.Igv, c.Total, c.Estado
        FROM dbo.ComprobantesElectronicos c
        WHERE c.NegocioId = @NegocioId
          AND c.Id = @Id;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Comprobantes_Crear
    @NegocioId INT,
    @ReservaId INT,
    @TipoComprobante INT,
    @Serie NVARCHAR(4),
    @Numero INT,
    @FechaEmision DATETIME2,
    @TipoMoneda INT,
    @SubTotal DECIMAL(10,2),
    @Igv DECIMAL(10,2),
    @Total DECIMAL(10,2),
    @Estado INT,
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        DECLARE @ClienteId INT;
        SELECT @ClienteId = r.ClienteId FROM dbo.Reservas r WHERE r.Id = @ReservaId;

        INSERT INTO dbo.ComprobantesElectronicos
        (
            NegocioId, ReservaId, ClienteId, TipoComprobante, Serie, Numero,
            FechaEmision, TipoMoneda, CodigoTipoOperacionSunat, CodigoTipoDocumentoClienteSunat,
            SubTotal, Igv, Total, Estado, FechaRegistro, UsuarioCreacion
        )
        VALUES
        (
            @NegocioId, @ReservaId, @ClienteId, @TipoComprobante, @Serie, @Numero,
            @FechaEmision, @TipoMoneda, N'0101', N'1',
            @SubTotal, @Igv, @Total, @Estado, SYSUTCDATETIME(), @Usuario
        );

        DECLARE @Id INT;
        SET @Id = SCOPE_IDENTITY();
        DECLARE @EntidadIdAudit NVARCHAR(80);
            SET @EntidadIdAudit = CONVERT(NVARCHAR(80), @Id);
        EXEC dbo.Sp_Auditoria_Registrar @NegocioId = @NegocioId, @Modulo = N'COMPROBANTES', @Accion = N'CREATE', @Entidad = N'ComprobanteElectronico', @EntidadId = @EntidadIdAudit, @Usuario = @Usuario, @DetalleJson = NULL;
        SELECT @Id;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Comprobantes_Actualizar
    @Id INT,
    @NegocioId INT,
    @ReservaId INT,
    @TipoComprobante INT,
    @Serie NVARCHAR(4),
    @Numero INT,
    @FechaEmision DATETIME2,
    @TipoMoneda INT,
    @SubTotal DECIMAL(10,2),
    @Igv DECIMAL(10,2),
    @Total DECIMAL(10,2),
    @Estado INT,
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        UPDATE dbo.ComprobantesElectronicos
        SET ReservaId = @ReservaId,
            TipoComprobante = @TipoComprobante,
            Serie = @Serie,
            Numero = @Numero,
            FechaEmision = @FechaEmision,
            TipoMoneda = @TipoMoneda,
            SubTotal = @SubTotal,
            Igv = @Igv,
            Total = @Total,
            Estado = @Estado,
            FechaActualizacion = SYSUTCDATETIME(),
            UsuarioActualizacion = @Usuario
        WHERE Id = @Id
          AND NegocioId = @NegocioId;

        IF @@ROWCOUNT = 0
            RAISERROR('No se encontro el comprobante para actualizar en el negocio.', 16, 1);

        DECLARE @EntidadIdAudit NVARCHAR(80);
        SET @EntidadIdAudit = CONVERT(NVARCHAR(80), @Id);
        EXEC dbo.Sp_Auditoria_Registrar @NegocioId = @NegocioId, @Modulo = N'COMPROBANTES', @Accion = N'EDIT', @Entidad = N'ComprobanteElectronico', @EntidadId = @EntidadIdAudit, @Usuario = @Usuario, @DetalleJson = NULL;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Comprobantes_Eliminar
    @NegocioId INT,
    @Id INT,
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        UPDATE dbo.ComprobantesElectronicos
        SET Estado = 5,
            FechaActualizacion = SYSUTCDATETIME(),
            UsuarioActualizacion = @Usuario
        WHERE Id = @Id
          AND NegocioId = @NegocioId;

        IF @@ROWCOUNT = 0
            RAISERROR('No se encontro el comprobante para eliminar en el negocio.', 16, 1);

        DECLARE @EntidadIdAudit NVARCHAR(80);
        SET @EntidadIdAudit = CONVERT(NVARCHAR(80), @Id);
        EXEC dbo.Sp_Auditoria_Registrar @NegocioId = @NegocioId, @Modulo = N'COMPROBANTES', @Accion = N'DELETE', @Entidad = N'ComprobanteElectronico', @EntidadId = @EntidadIdAudit, @Usuario = @Usuario, @DetalleJson = NULL;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO
