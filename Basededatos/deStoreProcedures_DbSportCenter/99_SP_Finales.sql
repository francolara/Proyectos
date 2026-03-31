-- =============================================
-- Author:        FRANCO LARA
-- Create date:   30/03/2026
-- Description:   Consolidado final de stored procedures (ultima version efectiva por nombre) generado automaticamente.
-- Firma:         Codex - 30/03/2026 | Script final para evitar sobreescritura por orden de despliegue; incluye filtro de estados multiples en Sp_Reservas_Listar.
-- =============================================
-- REGLA DE USO:
-- 1) Ejecutar primero los scripts estructurales y funcionales (00..32).
-- 2) Ejecutar este archivo al final.
-- 3) Regenerar este archivo con Generate-99_SP_Finales.ps1 cada vez que cambie un SP.

USE [DbSportCenter];
GO
-- SOURCE: 00_Auditoria.sql (linea 7)
CREATE OR ALTER PROCEDURE dbo.Sp_Auditoria_Registrar
    @NegocioId INT = NULL,
    @Modulo NVARCHAR(50),
    @Accion NVARCHAR(20),
    @Entidad NVARCHAR(80),
    @EntidadId NVARCHAR(80),
    @Usuario NVARCHAR(200),
    @DetalleJson NVARCHAR(4000) = NULL
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        INSERT INTO dbo.BitacoraAuditoria
        (
            NegocioId,
            Modulo,
            Accion,
            Entidad,
            EntidadId,
            UsuarioId,
            UsuarioNombre,
            UsuarioCorreo,
            DetalleJson,
            FechaRegistro
        )
        VALUES
        (
            @NegocioId,
            @Modulo,
            @Accion,
            @Entidad,
            @EntidadId,
            @Usuario,
            @Usuario,
            NULL,
            @DetalleJson,
            SYSUTCDATETIME()
        );
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

-- SOURCE: 01_Seguridad_Panel.sql (linea 149)
CREATE OR ALTER PROCEDURE dbo.Sp_Panel_ListarNegociosUsuario
    @UsuarioId NVARCHAR(450)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT
            un.NegocioId,
            n.NombreComercial,
            CAST(un.RolNegocio AS NVARCHAR(20)) AS Rol
        FROM dbo.UsuariosNegocio un
        INNER JOIN dbo.Negocios n ON n.Id = un.NegocioId
        WHERE un.UsuarioId = @UsuarioId
          AND un.Activo = 1
          AND n.Activo = 1
        ORDER BY n.NombreComercial;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

-- SOURCE: 01_Seguridad_Panel.sql (linea 174)
CREATE OR ALTER PROCEDURE dbo.Sp_Panel_ObtenerRolUsuario
    @UsuarioId NVARCHAR(450),
    @NegocioId INT
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT TOP (1)
            CAST(un.RolNegocio AS NVARCHAR(20))
        FROM dbo.UsuariosNegocio un
        WHERE un.UsuarioId = @UsuarioId
          AND un.NegocioId = @NegocioId
          AND un.Activo = 1;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

-- SOURCE: 02_Home.sql (linea 27)
CREATE OR ALTER PROCEDURE dbo.Sp_Home_ListarTiposDeporte
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT td.Id, td.Nombre
        FROM dbo.TiposDeporte td
        WHERE td.Activo = 1
        ORDER BY td.Nombre;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

-- SOURCE: 03_Sedes_Espacios.sql (linea 158)
CREATE OR ALTER PROCEDURE dbo.Sp_Sedes_Eliminar
    @NegocioId INT,
    @Id INT,
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        UPDATE dbo.Sedes
        SET Activo = 0,
            FechaActualizacion = SYSUTCDATETIME(),
            UsuarioActualizacion = @Usuario
        WHERE Id = @Id
          AND NegocioId = @NegocioId;

        IF @@ROWCOUNT > 0
        BEGIN
            DECLARE @EntidadIdAudit NVARCHAR(80);
            SET @EntidadIdAudit = CONVERT(NVARCHAR(80), @Id);
            EXEC dbo.Sp_Auditoria_Registrar @NegocioId = @NegocioId, @Modulo = N'SEDES', @Accion = N'DELETE', @Entidad = N'Sede', @EntidadId = @EntidadIdAudit, @Usuario = @Usuario, @DetalleJson = NULL;
        END
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

-- SOURCE: 03_Sedes_Espacios.sql (linea 340)
CREATE OR ALTER PROCEDURE dbo.Sp_Espacios_Eliminar
    @NegocioId INT,
    @Id INT,
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
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
            EXEC dbo.Sp_Auditoria_Registrar @NegocioId = @NegocioId, @Modulo = N'ESPACIOS', @Accion = N'DELETE', @Entidad = N'EspacioDeportivo', @EntidadId = @EntidadIdAudit, @Usuario = @Usuario, @DetalleJson = NULL;
        END
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

-- SOURCE: 04_Reservas_Pagos_Comprobantes.sql (linea 104)
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

-- SOURCE: 04_Reservas_Pagos_Comprobantes.sql (linea 223)
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

-- SOURCE: 04_Reservas_Pagos_Comprobantes.sql (linea 286)
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

-- SOURCE: 04_Reservas_Pagos_Comprobantes.sql (linea 445)
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

-- SOURCE: 04_Reservas_Pagos_Comprobantes.sql (linea 465)
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

-- SOURCE: 04_Reservas_Pagos_Comprobantes.sql (linea 513)
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

-- SOURCE: 04_Reservas_Pagos_Comprobantes.sql (linea 562)
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

-- SOURCE: 05_Sedes_Servicios.sql (linea 61)
CREATE OR ALTER PROCEDURE dbo.Sp_Combos_ServiciosSede
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT cs.Id, cs.Nombre
        FROM dbo.CatalogoServiciosSede cs
        WHERE cs.Activo = 1
        ORDER BY cs.Nombre;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

-- SOURCE: 06_Seguridad_Clientes.sql (linea 151)
CREATE OR ALTER PROCEDURE dbo.Sp_Combos_Clientes
    @NegocioId INT
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT c.Id, CONCAT(c.NombresORazonSocial, N' (', c.NumeroDocumento, N')')
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

-- SOURCE: 06_Seguridad_Clientes.sql (linea 173)
CREATE OR ALTER PROCEDURE dbo.Sp_Clientes_Listar
    @NegocioId INT
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT
            c.Id,
            c.NombresORazonSocial,
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

-- SOURCE: 06_Seguridad_Clientes.sql (linea 201)
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

-- SOURCE: 06_Seguridad_Clientes.sql (linea 231)
CREATE OR ALTER PROCEDURE dbo.Sp_Clientes_Crear
    @NegocioId INT,
    @NombresORazonSocial NVARCHAR(200),
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
        SET @NumeroDocumentoNormalizado = NULLIF(LTRIM(RTRIM(@NumeroDocumento)), N'');
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
            NombresORazonSocial, TipoDocumento, NumeroDocumento, Telefono,
            Correo, DireccionFiscal, Activo, FechaCreacion, UsuarioCreacion
        )
        VALUES
        (
            @NombresORazonSocial, @TipoDocumento, @NumeroDocumento, @Telefono,
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

-- SOURCE: 06_Seguridad_Clientes.sql (linea 300)
CREATE OR ALTER PROCEDURE dbo.Sp_Clientes_Actualizar
    @Id INT,
    @NegocioId INT,
    @NombresORazonSocial NVARCHAR(200),
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
        SET @NumeroDocumentoNormalizado = NULLIF(LTRIM(RTRIM(@NumeroDocumento)), N'');
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

-- SOURCE: 06_Seguridad_Clientes.sql (linea 365)
CREATE OR ALTER PROCEDURE dbo.Sp_Clientes_Eliminar
    @NegocioId INT,
    @Id INT,
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        UPDATE nc
        SET nc.Activo = 0
        FROM dbo.NegocioClientes nc
        WHERE nc.NegocioId = @NegocioId
          AND nc.ClienteId = @Id
          AND nc.Activo = 1;

        IF @@ROWCOUNT = 0
            RAISERROR('No se encontro el cliente para eliminar en el negocio.', 16, 1);

        IF NOT EXISTS (SELECT 1 FROM dbo.NegocioClientes WHERE ClienteId = @Id AND Activo = 1)
        BEGIN
            UPDATE dbo.Clientes
            SET Activo = 0,
                FechaActualizacion = SYSUTCDATETIME(),
                UsuarioActualizacion = @Usuario
            WHERE Id = @Id;
        END;

        DECLARE @EntidadIdAudit NVARCHAR(80);
        SET @EntidadIdAudit = CONVERT(NVARCHAR(80), @Id);
        EXEC dbo.Sp_Auditoria_Registrar @NegocioId = @NegocioId, @Modulo = N'CLIENTES', @Accion = N'DELETE', @Entidad = N'Cliente', @EntidadId = @EntidadIdAudit, @Usuario = @Usuario, @DetalleJson = NULL;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

-- SOURCE: 07_Reservas_Pagos_Reglas.sql (linea 79)
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
        IF @HoraFin <= @HoraInicio
            RAISERROR('La hora fin debe ser mayor que la hora inicio.', 16, 1);

        IF @Total < 0 OR @Adelanto < 0
            RAISERROR('Los montos no pueden ser negativos.', 16, 1);

        IF @Adelanto > @Total
            RAISERROR('El adelanto no puede ser mayor que el total.', 16, 1);

        IF NOT EXISTS (
            SELECT 1
            FROM dbo.EspaciosDeportivos e
            INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
            WHERE e.Id = @EspacioDeportivoId
              AND s.NegocioId = @NegocioId
              AND e.Estado = 1
        )
            RAISERROR('El espacio deportivo no esta disponible para este negocio.', 16, 1);

        IF EXISTS (
            SELECT 1
            FROM dbo.Reservas r
            WHERE r.EspacioDeportivoId = @EspacioDeportivoId
              AND r.Fecha = @Fecha
              AND r.Estado NOT IN (5, 6)
              AND r.Id <> @Id
              AND @HoraInicio < r.HoraFin
              AND @HoraFin > r.HoraInicio
        )
            RAISERROR('Cruce de horario detectado.', 16, 1);

        UPDATE r
        SET r.EspacioDeportivoId = @EspacioDeportivoId,
            r.ClienteId = @ClienteId,
            r.Fecha = @Fecha,
            r.HoraInicio = @HoraInicio,
            r.HoraFin = @HoraFin,
            r.Total = @Total,
            r.Adelanto = @Adelanto,
            r.Saldo = (@Total - @Adelanto),
            r.Estado = @Estado,
            r.FechaActualizacion = SYSUTCDATETIME(),
            r.UsuarioActualizacion = @Usuario
        FROM dbo.Reservas r
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        WHERE r.Id = @Id
          AND s.NegocioId = @NegocioId;

        IF @@ROWCOUNT = 0
            RAISERROR('No se encontro la reserva para actualizar.', 16, 1);

        IF @@ROWCOUNT > 0
        BEGIN
            DECLARE @EntidadIdAudit NVARCHAR(80);
            SET @EntidadIdAudit = CONVERT(NVARCHAR(80), @Id);
            EXEC dbo.Sp_Auditoria_Registrar @NegocioId = @NegocioId, @Modulo = N'RESERVAS', @Accion = N'EDIT', @Entidad = N'Reserva', @EntidadId = @EntidadIdAudit, @Usuario = @Usuario, @DetalleJson = NULL;
        END;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

-- SOURCE: 07_Reservas_Pagos_Reglas.sql (linea 162)
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

-- SOURCE: 07_Reservas_Pagos_Reglas.sql (linea 243)
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

-- SOURCE: 07_Reservas_Pagos_Reglas.sql (linea 340)
CREATE OR ALTER PROCEDURE dbo.Sp_Pagos_Eliminar
    @NegocioId INT,
    @Id INT,
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        DECLARE @ReservaId INT;
        SELECT @ReservaId = p.ReservaId
        FROM dbo.Pagos p
        INNER JOIN dbo.Reservas r ON r.Id = p.ReservaId
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        WHERE p.Id = @Id
          AND s.NegocioId = @NegocioId;

        IF @ReservaId IS NULL
            RAISERROR('No se encontro el pago para eliminar en el negocio.', 16, 1);

        BEGIN TRANSACTION;

        DELETE FROM dbo.Pagos WHERE Id = @Id;

        UPDATE r
        SET Adelanto = x.Pagado,
            Saldo = (r.Total - x.Pagado),
            FechaActualizacion = SYSUTCDATETIME(),
            UsuarioActualizacion = @Usuario
        FROM dbo.Reservas r
        INNER JOIN (
            SELECT @ReservaId AS ReservaId, COALESCE(SUM(p.Monto), 0) AS Pagado
            FROM dbo.Pagos p
            WHERE p.ReservaId = @ReservaId
        ) x ON x.ReservaId = r.Id;

        IF @@ROWCOUNT > 0
        BEGIN
            DECLARE @EntidadIdAudit NVARCHAR(80);
            SET @EntidadIdAudit = CONVERT(NVARCHAR(80), @Id);
            EXEC dbo.Sp_Auditoria_Registrar @NegocioId = @NegocioId, @Modulo = N'PAGOS', @Accion = N'DELETE', @Entidad = N'Pago', @EntidadId = @EntidadIdAudit, @Usuario = @Usuario, @DetalleJson = NULL;
        END;

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

-- SOURCE: 08_Reservas_Calendario_Filtros.sql (linea 7)
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
            c.NombresORazonSocial AS Cliente,
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

-- SOURCE: 10_Home_Solicitudes_Publicas.sql (linea 35)
CREATE OR ALTER PROCEDURE dbo.Sp_Home_SolicitarReservaPublica
    @EspacioDeportivoId INT,
    @Fecha DATE,
    @HoraInicio TIME,
    @HoraFin TIME,
    @NombreSolicitante NVARCHAR(200),
    @Telefono NVARCHAR(30),
    @Correo NVARCHAR(200) = NULL,
    @Comentario NVARCHAR(300) = NULL
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF @HoraFin <= @HoraInicio
            RAISERROR('La hora fin debe ser mayor que la hora inicio.', 16, 1);

        IF NOT EXISTS (
            SELECT 1
            FROM dbo.EspaciosDeportivos e
            WHERE e.Id = @EspacioDeportivoId
              AND e.Estado = 1
        )
            RAISERROR('El espacio deportivo no esta disponible.', 16, 1);

        IF EXISTS (
            SELECT 1
            FROM dbo.Reservas r
            WHERE r.EspacioDeportivoId = @EspacioDeportivoId
              AND r.Fecha = @Fecha
              AND r.Estado NOT IN (5, 6)
              AND @HoraInicio < r.HoraFin
              AND @HoraFin > r.HoraInicio
        )
            RAISERROR('El horario seleccionado ya no esta disponible.', 16, 1);

        INSERT INTO dbo.SolicitudesReservaPublica
        (
            EspacioDeportivoId, Fecha, HoraInicio, HoraFin,
            NombreSolicitante, Telefono, Correo, Comentario,
            Estado, NotificadoCliente, FechaRegistro
        )
        VALUES
        (
            @EspacioDeportivoId, @Fecha, @HoraInicio, @HoraFin,
            @NombreSolicitante, @Telefono, @Correo, @Comentario,
            1, 0, SYSUTCDATETIME()
        );

        DECLARE @Id INT;
        SET @Id = SCOPE_IDENTITY();

        DECLARE @CodigoSolicitud NVARCHAR(20);
        SET @CodigoSolicitud = CONCAT(N'SR', FORMAT(GETDATE(), 'yyyyMMdd'), RIGHT(CONCAT(N'00000', CONVERT(NVARCHAR(10), @Id)), 5));

        UPDATE dbo.SolicitudesReservaPublica
        SET CodigoSolicitud = @CodigoSolicitud
        WHERE Id = @Id;

        SELECT @CodigoSolicitud;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

-- SOURCE: 11_Solicitudes_Gestion.sql (linea 116)
CREATE OR ALTER PROCEDURE dbo.Sp_SolicitudesPublicas_Listar
    @NegocioId INT,
    @FechaDesde DATE = NULL,
    @FechaHasta DATE = NULL,
    @Estado INT = NULL
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT
            s.Id,
            s.CodigoSolicitud,
            se.Nombre AS Sede,
            e.Nombre AS Espacio,
            s.Fecha,
            s.HoraInicio,
            s.HoraFin,
            s.NombreSolicitante,
            s.Telefono,
            s.Correo,
            s.Estado,
            s.ReservaId,
            s.FechaRegistro
        FROM dbo.SolicitudesReservaPublica s
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = s.EspacioDeportivoId
        INNER JOIN dbo.Sedes se ON se.Id = e.SedeId
        WHERE se.NegocioId = @NegocioId
          AND (@FechaDesde IS NULL OR s.Fecha >= @FechaDesde)
          AND (@FechaHasta IS NULL OR s.Fecha <= @FechaHasta)
          AND (@Estado IS NULL OR s.Estado = @Estado)
        ORDER BY s.FechaRegistro DESC;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

-- SOURCE: 11_Solicitudes_Gestion.sql (linea 156)
CREATE OR ALTER PROCEDURE dbo.Sp_SolicitudesPublicas_ActualizarEstado
    @NegocioId INT,
    @Id INT,
    @Estado INT,
    @ComentarioGestion NVARCHAR(300) = NULL,
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF @Estado NOT IN (2, 3)
            RAISERROR('Estado invalido. Solo se permite aprobar(2) o rechazar(3).', 16, 1);

        UPDATE s
        SET s.Estado = @Estado,
            s.ComentarioGestion = @ComentarioGestion,
            s.FechaGestion = SYSUTCDATETIME(),
            s.UsuarioGestion = @Usuario
        FROM dbo.SolicitudesReservaPublica s
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = s.EspacioDeportivoId
        INNER JOIN dbo.Sedes se ON se.Id = e.SedeId
        WHERE s.Id = @Id
          AND se.NegocioId = @NegocioId
          AND s.Estado IN (1, 2, 3);

        IF @@ROWCOUNT = 0
            RAISERROR('No se encontro la solicitud para actualizar en el negocio.', 16, 1);

        DECLARE @EntidadIdAudit NVARCHAR(80);
        SET @EntidadIdAudit = CONVERT(NVARCHAR(80), @Id);
        EXEC dbo.Sp_Auditoria_Registrar @NegocioId = @NegocioId, @Modulo = N'SOLICITUDES', @Accion = N'EDIT', @Entidad = N'SolicitudReservaPublica', @EntidadId = @EntidadIdAudit, @Usuario = @Usuario, @DetalleJson = NULL;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

-- SOURCE: 11_Solicitudes_Gestion.sql (linea 196)
CREATE OR ALTER PROCEDURE dbo.Sp_SolicitudesPublicas_ConvertirAReserva
    @NegocioId INT,
    @Id INT,
    @Total DECIMAL(10,2),
    @Adelanto DECIMAL(10,2),
    @EstadoReserva INT,
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF @Total < 0 OR @Adelanto < 0 OR @Adelanto > @Total
            RAISERROR('Montos invalidos para la conversion.', 16, 1);

        DECLARE @EspacioDeportivoId INT, @Fecha DATE, @HoraInicio TIME, @HoraFin TIME, @NombreSolicitante NVARCHAR(200), @Telefono NVARCHAR(30), @Correo NVARCHAR(200);
        DECLARE @ClienteId INT, @ReservaId INT;

        SELECT
            @EspacioDeportivoId = s.EspacioDeportivoId,
            @Fecha = s.Fecha,
            @HoraInicio = s.HoraInicio,
            @HoraFin = s.HoraFin,
            @NombreSolicitante = s.NombreSolicitante,
            @Telefono = s.Telefono,
            @Correo = s.Correo
        FROM dbo.SolicitudesReservaPublica s
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = s.EspacioDeportivoId
        INNER JOIN dbo.Sedes se ON se.Id = e.SedeId
        WHERE s.Id = @Id
          AND se.NegocioId = @NegocioId
          AND s.Estado IN (1, 2);

        IF @EspacioDeportivoId IS NULL
            RAISERROR('Solicitud invalida para el negocio.', 16, 1);

        IF EXISTS (
            SELECT 1
            FROM dbo.Reservas r
            WHERE r.EspacioDeportivoId = @EspacioDeportivoId
              AND r.Fecha = @Fecha
              AND r.Estado NOT IN (5, 6)
              AND @HoraInicio < r.HoraFin
              AND @HoraFin > r.HoraInicio
        )
            RAISERROR('No se puede convertir: el horario ya fue tomado.', 16, 1);

        SELECT TOP (1) @ClienteId = c.Id
        FROM dbo.Clientes c
        INNER JOIN dbo.NegocioClientes nc ON nc.ClienteId = c.Id
        WHERE nc.NegocioId = @NegocioId
          AND nc.Activo = 1
          AND c.Activo = 1
          AND c.NombresORazonSocial = @NombreSolicitante
          AND c.Telefono = @Telefono;

        BEGIN TRANSACTION;

        IF @ClienteId IS NULL
        BEGIN
            INSERT INTO dbo.Clientes
            (
                NombresORazonSocial, TipoDocumento, NumeroDocumento, Telefono, Correo,
                Activo, FechaCreacion, UsuarioCreacion
            )
            VALUES
            (
                @NombreSolicitante, N'OTRO', CONCAT(N'SOL', @Id), @Telefono, @Correo,
                1, SYSUTCDATETIME(), @Usuario
            );

            SET @ClienteId = SCOPE_IDENTITY();

            INSERT INTO dbo.NegocioClientes (NegocioId, ClienteId, Activo, FechaRegistro, UsuarioCreacion)
            VALUES (@NegocioId, @ClienteId, 1, SYSUTCDATETIME(), @Usuario);
        END;

        INSERT INTO dbo.Reservas
        (
            EspacioDeportivoId, ClienteId, Fecha, HoraInicio, HoraFin,
            Estado, Total, Adelanto, Saldo, FechaRegistro, UsuarioCreacion
        )
        VALUES
        (
            @EspacioDeportivoId, @ClienteId, @Fecha, @HoraInicio, @HoraFin,
            @EstadoReserva, @Total, @Adelanto, (@Total - @Adelanto), SYSUTCDATETIME(), @Usuario
        );

        SET @ReservaId = SCOPE_IDENTITY();

        UPDATE dbo.SolicitudesReservaPublica
        SET Estado = 4,
            ReservaId = @ReservaId,
            FechaGestion = SYSUTCDATETIME(),
            UsuarioGestion = @Usuario,
            ComentarioGestion = N'Convertida a reserva'
        WHERE Id = @Id;

        DECLARE @EntidadIdAudit NVARCHAR(80);
        SET @EntidadIdAudit = CONVERT(NVARCHAR(80), @Id);
        EXEC dbo.Sp_Auditoria_Registrar @NegocioId = @NegocioId, @Modulo = N'SOLICITUDES', @Accion = N'EDIT', @Entidad = N'SolicitudReservaPublica', @EntidadId = @EntidadIdAudit, @Usuario = @Usuario, @DetalleJson = NULL;

        SET @EntidadIdAudit = CONVERT(NVARCHAR(80), @ReservaId);
        EXEC dbo.Sp_Auditoria_Registrar @NegocioId = @NegocioId, @Modulo = N'RESERVAS', @Accion = N'CREATE', @Entidad = N'Reserva', @EntidadId = @EntidadIdAudit, @Usuario = @Usuario, @DetalleJson = NULL;

        COMMIT TRANSACTION;

        SELECT @ReservaId;
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

-- SOURCE: 12_Home_Notificaciones_Seguimiento.sql (linea 7)
CREATE OR ALTER PROCEDURE dbo.Sp_Home_ConsultarSolicitudPublica
    @CodigoSolicitud NVARCHAR(20),
    @Telefono NVARCHAR(30)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT TOP (1)
            s.CodigoSolicitud,
            se.Nombre AS Sede,
            e.Nombre AS Espacio,
            s.Fecha,
            s.HoraInicio,
            s.HoraFin,
            s.NombreSolicitante,
            s.Telefono,
            s.Correo,
            s.Estado,
            CASE s.Estado
                WHEN 1 THEN N'Pendiente'
                WHEN 2 THEN N'Aprobada'
                WHEN 3 THEN N'Rechazada'
                WHEN 4 THEN N'Convertida a reserva'
                ELSE N'Desconocido'
            END AS EstadoTexto,
            s.ReservaId,
            s.FechaRegistro
        FROM dbo.SolicitudesReservaPublica s
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = s.EspacioDeportivoId
        INNER JOIN dbo.Sedes se ON se.Id = e.SedeId
        WHERE s.CodigoSolicitud = @CodigoSolicitud
          AND s.Telefono = @Telefono
        ORDER BY s.Id DESC;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

-- SOURCE: 12_Home_Notificaciones_Seguimiento.sql (linea 49)
CREATE OR ALTER PROCEDURE dbo.Sp_Home_ObtenerSolicitudParaNotificacion
    @CodigoSolicitud NVARCHAR(20)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT TOP (1)
            s.CodigoSolicitud,
            s.NombreSolicitante,
            s.Correo,
            se.Nombre AS Sede,
            e.Nombre AS Espacio,
            s.Fecha,
            s.HoraInicio,
            s.HoraFin,
            s.NotificadoCliente
        FROM dbo.SolicitudesReservaPublica s
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = s.EspacioDeportivoId
        INNER JOIN dbo.Sedes se ON se.Id = e.SedeId
        WHERE s.CodigoSolicitud = @CodigoSolicitud;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

-- SOURCE: 12_Home_Notificaciones_Seguimiento.sql (linea 78)
CREATE OR ALTER PROCEDURE dbo.Sp_Home_MarcarSolicitudNotificada
    @CodigoSolicitud NVARCHAR(20)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        UPDATE dbo.SolicitudesReservaPublica
        SET NotificadoCliente = 1
        WHERE CodigoSolicitud = @CodigoSolicitud;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

-- SOURCE: 13_Usuarios_Negocio_Gestion.sql (linea 183)
CREATE OR ALTER PROCEDURE dbo.Sp_UsuariosNegocio_Desactivar
    @NegocioId INT,
    @UsuarioNegocioId INT,
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        UPDATE dbo.UsuariosNegocio
        SET Activo = 0
        WHERE Id = @UsuarioNegocioId
          AND NegocioId = @NegocioId;

        IF @@ROWCOUNT = 0
            RAISERROR('No se encontro el usuario del negocio para desactivar.', 16, 1);

        IF @@ROWCOUNT > 0
        BEGIN
            DECLARE @EntidadIdAudit NVARCHAR(80);
            SET @EntidadIdAudit = CONVERT(NVARCHAR(80), @UsuarioNegocioId);
            EXEC dbo.Sp_Auditoria_Registrar @NegocioId = @NegocioId, @Modulo = N'USUARIOS', @Accion = N'DELETE', @Entidad = N'UsuarioNegocio', @EntidadId = @EntidadIdAudit, @Usuario = @Usuario, @DetalleJson = NULL;
        END;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

-- SOURCE: 13_Usuarios_Negocio_Gestion.sql (linea 214)
CREATE OR ALTER PROCEDURE dbo.Sp_UsuariosNegocio_PermisosListar
    @NegocioId INT,
    @UsuarioNegocioId INT
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF NOT EXISTS (SELECT 1 FROM dbo.UsuariosNegocio WHERE Id = @UsuarioNegocioId AND NegocioId = @NegocioId)
            RAISERROR('UsuarioNegocio invalido para el negocio.', 16, 1);

        DECLARE @RolNegocio INT;
        SELECT @RolNegocio = RolNegocio FROM dbo.UsuariosNegocio WHERE Id = @UsuarioNegocioId;

        SELECT
            m.Id AS ModuloSistemaId,
            m.Codigo,
            m.Nombre,
            COALESCE(up.PuedeVer, rp.PuedeVer) AS PuedeVer,
            COALESCE(up.PuedeCrear, rp.PuedeCrear) AS PuedeCrear,
            COALESCE(up.PuedeEditar, rp.PuedeEditar) AS PuedeEditar,
            COALESCE(up.PuedeEliminar, rp.PuedeEliminar) AS PuedeEliminar
        FROM dbo.ModulosSistema m
        INNER JOIN dbo.RolesNegocioPermiso rp ON rp.ModuloSistemaId = m.Id AND rp.RolNegocio = @RolNegocio
        LEFT JOIN dbo.UsuariosNegocioPermiso up ON up.ModuloSistemaId = m.Id AND up.UsuarioNegocioId = @UsuarioNegocioId
        WHERE m.Activo = 1
        ORDER BY m.Nombre;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

-- SOURCE: 13_Usuarios_Negocio_Gestion.sql (linea 249)
CREATE OR ALTER PROCEDURE dbo.Sp_UsuariosNegocio_PermisoGuardar
    @NegocioId INT,
    @UsuarioNegocioId INT,
    @ModuloSistemaId INT,
    @PuedeVer BIT,
    @PuedeCrear BIT,
    @PuedeEditar BIT,
    @PuedeEliminar BIT,
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF NOT EXISTS (SELECT 1 FROM dbo.UsuariosNegocio WHERE Id = @UsuarioNegocioId AND NegocioId = @NegocioId)
            RAISERROR('UsuarioNegocio invalido para el negocio.', 16, 1);

        IF EXISTS (SELECT 1 FROM dbo.UsuariosNegocioPermiso WHERE UsuarioNegocioId = @UsuarioNegocioId AND ModuloSistemaId = @ModuloSistemaId)
        BEGIN
            UPDATE dbo.UsuariosNegocioPermiso
            SET PuedeVer = @PuedeVer,
                PuedeCrear = @PuedeCrear,
                PuedeEditar = @PuedeEditar,
                PuedeEliminar = @PuedeEliminar
            WHERE UsuarioNegocioId = @UsuarioNegocioId
              AND ModuloSistemaId = @ModuloSistemaId;
        END
        ELSE
        BEGIN
            INSERT INTO dbo.UsuariosNegocioPermiso
            (
                UsuarioNegocioId, ModuloSistemaId, PuedeVer, PuedeCrear, PuedeEditar, PuedeEliminar
            )
            VALUES
            (
                @UsuarioNegocioId, @ModuloSistemaId, @PuedeVer, @PuedeCrear, @PuedeEditar, @PuedeEliminar
            );
        END;

        DECLARE @EntidadIdAudit NVARCHAR(80);
        SET @EntidadIdAudit = CONCAT(CONVERT(NVARCHAR(30), @UsuarioNegocioId), N'-', CONVERT(NVARCHAR(30), @ModuloSistemaId));
        EXEC dbo.Sp_Auditoria_Registrar @NegocioId = @NegocioId, @Modulo = N'USUARIOS', @Accion = N'EDIT', @Entidad = N'UsuarioNegocioPermiso', @EntidadId = @EntidadIdAudit, @Usuario = @Usuario, @DetalleJson = NULL;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

-- SOURCE: 14_Promociones_Kpis.sql (linea 39)
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

-- SOURCE: 14_Promociones_Kpis.sql (linea 242)
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

-- SOURCE: 14_Promociones_Kpis.sql (linea 272)
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

-- SOURCE: 14_Promociones_Kpis.sql (linea 321)
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

-- SOURCE: 14_Promociones_Kpis.sql (linea 375)
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

-- SOURCE: 15_Calendario_Bloqueos.sql (linea 185)
CREATE OR ALTER PROCEDURE dbo.Sp_Bloqueos_Listar
    @NegocioId INT,
    @FechaDesde DATE,
    @FechaHasta DATE,
    @SedeId INT = NULL,
    @EspacioDeportivoId INT = NULL
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT
            b.Id,
            s.Nombre AS Sede,
            e.Nombre AS Espacio,
            b.Fecha,
            b.HoraInicio,
            b.HoraFin,
            b.Motivo,
            b.Activo
        FROM dbo.BloqueosHorario b
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = b.EspacioDeportivoId
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        WHERE s.NegocioId = @NegocioId
          AND b.Activo = 1
          AND b.Fecha BETWEEN @FechaDesde AND @FechaHasta
          AND (@SedeId IS NULL OR s.Id = @SedeId)
          AND (@EspacioDeportivoId IS NULL OR e.Id = @EspacioDeportivoId)
        ORDER BY b.Fecha, b.HoraInicio;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

-- SOURCE: 15_Calendario_Bloqueos.sql (linea 222)
CREATE OR ALTER PROCEDURE dbo.Sp_Bloqueos_Crear
    @NegocioId INT,
    @EspacioDeportivoId INT,
    @Fecha DATE,
    @HoraInicio TIME,
    @HoraFin TIME,
    @Motivo NVARCHAR(250),
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF @HoraFin <= @HoraInicio
            RAISERROR('La hora fin debe ser mayor a la hora inicio.', 16, 1);

        IF NOT EXISTS (
            SELECT 1
            FROM dbo.EspaciosDeportivos e
            INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
            WHERE e.Id = @EspacioDeportivoId
              AND s.NegocioId = @NegocioId
        )
            RAISERROR('Espacio no valido para el negocio.', 16, 1);

        IF EXISTS (
            SELECT 1
            FROM dbo.BloqueosHorario b
            WHERE b.EspacioDeportivoId = @EspacioDeportivoId
              AND b.Fecha = @Fecha
              AND b.Activo = 1
              AND @HoraInicio < b.HoraFin
              AND @HoraFin > b.HoraInicio
        )
            RAISERROR('Ya existe un bloqueo que se cruza con ese horario.', 16, 1);

        IF EXISTS (
            SELECT 1
            FROM dbo.Reservas r
            WHERE r.EspacioDeportivoId = @EspacioDeportivoId
              AND r.Fecha = @Fecha
              AND r.Estado NOT IN (5, 6)
              AND @HoraInicio < r.HoraFin
              AND @HoraFin > r.HoraInicio
        )
            RAISERROR('Existe una reserva en ese horario, no se puede bloquear.', 16, 1);

        INSERT INTO dbo.BloqueosHorario
        (
            EspacioDeportivoId, Fecha, HoraInicio, HoraFin, Motivo,
            Activo, FechaCreacion, UsuarioCreacion
        )
        VALUES
        (
            @EspacioDeportivoId, @Fecha, @HoraInicio, @HoraFin, @Motivo,
            1, SYSUTCDATETIME(), @Usuario
        );

        DECLARE @Id INT;
        SET @Id = SCOPE_IDENTITY();
        DECLARE @EntidadIdAudit NVARCHAR(80);
        SET @EntidadIdAudit = CONVERT(NVARCHAR(80), @Id);
        EXEC dbo.Sp_Auditoria_Registrar
            @NegocioId = @NegocioId,
            @Modulo = N'RESERVAS',
            @Accion = N'BLOCK',
            @Entidad = N'BloqueoHorario',
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

-- SOURCE: 15_Calendario_Bloqueos.sql (linea 302)
CREATE OR ALTER PROCEDURE dbo.Sp_Bloqueos_Eliminar
    @NegocioId INT,
    @Id INT,
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        UPDATE b
        SET b.Activo = 0,
            b.FechaActualizacion = SYSUTCDATETIME(),
            b.UsuarioActualizacion = @Usuario
        FROM dbo.BloqueosHorario b
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = b.EspacioDeportivoId
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        WHERE b.Id = @Id
          AND s.NegocioId = @NegocioId;

        IF @@ROWCOUNT > 0
        BEGIN
            DECLARE @EntidadIdAudit NVARCHAR(80);
            SET @EntidadIdAudit = CONVERT(NVARCHAR(80), @Id);
            EXEC dbo.Sp_Auditoria_Registrar
                @NegocioId = @NegocioId,
                @Modulo = N'RESERVAS',
                @Accion = N'UNBLOCK',
                @Entidad = N'BloqueoHorario',
                @EntidadId = @EntidadIdAudit,
                @Usuario = @Usuario,
                @DetalleJson = NULL;
        END;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

-- SOURCE: 16_Reservas_CheckIn_CheckOut.sql (linea 8)
CREATE OR ALTER PROCEDURE dbo.Sp_Reservas_CambiarEstadoRapido
    @NegocioId INT,
    @Id INT,
    @NuevoEstado INT,
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        DECLARE @EstadoActual INT;

        SELECT @EstadoActual = r.Estado
        FROM dbo.Reservas r
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        WHERE r.Id = @Id
          AND s.NegocioId = @NegocioId;

        IF @EstadoActual IS NULL
            RAISERROR('No se encontro la reserva para cambio de estado.', 16, 1);

        IF @NuevoEstado NOT IN (3, 4, 6)
            RAISERROR('Estado no permitido para cambio rapido.', 16, 1);

        IF @EstadoActual IN (5, 6)
            RAISERROR('La reserva ya esta cancelada o marcada como no asistio.', 16, 1);

        IF @NuevoEstado = 3 AND @EstadoActual NOT IN (1, 2)
            RAISERROR('Check-in solo permitido para reservas pendientes o confirmadas.', 16, 1);

        IF @NuevoEstado = 4 AND @EstadoActual <> 3
            RAISERROR('Check-out solo permitido para reservas en uso.', 16, 1);

        IF @NuevoEstado = 6 AND @EstadoActual NOT IN (1, 2)
            RAISERROR('No-show solo permitido para reservas pendientes o confirmadas.', 16, 1);

        UPDATE dbo.Reservas
        SET Estado = @NuevoEstado,
            FechaActualizacion = SYSUTCDATETIME(),
            UsuarioActualizacion = @Usuario
        WHERE Id = @Id;

        IF @@ROWCOUNT > 0
        BEGIN
            DECLARE @EntidadIdAudit NVARCHAR(80);
            DECLARE @AccionAudit NVARCHAR(30);

            SET @EntidadIdAudit = CONVERT(NVARCHAR(80), @Id);
            SET @AccionAudit =
                CASE @NuevoEstado
                    WHEN 3 THEN N'CHECKIN'
                    WHEN 4 THEN N'CHECKOUT'
                    WHEN 6 THEN N'NOSHOW'
                    ELSE N'EDIT'
                END;

            EXEC dbo.Sp_Auditoria_Registrar
                @NegocioId = @NegocioId,
                @Modulo = N'RESERVAS',
                @Accion = @AccionAudit,
                @Entidad = N'Reserva',
                @EntidadId = @EntidadIdAudit,
                @Usuario = @Usuario,
                @DetalleJson = NULL;
        END;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

-- SOURCE: 17_Automatizacion_Recordatorios_NoShow.sql (linea 63)
CREATE OR ALTER PROCEDURE dbo.Sp_Reservas_MarcarRecordatorioEnviado
    @NegocioId INT,
    @ReservaId INT,
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        UPDATE r
        SET r.RecordatorioEnviado = 1,
            r.FechaRecordatorio = SYSUTCDATETIME(),
            r.FechaActualizacion = SYSUTCDATETIME(),
            r.UsuarioActualizacion = @Usuario
        FROM dbo.Reservas r
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        WHERE r.Id = @ReservaId
          AND s.NegocioId = @NegocioId;

        IF @@ROWCOUNT = 0
            RAISERROR('No se encontro la reserva para marcar recordatorio en el negocio.', 16, 1);
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

-- SOURCE: 18_Sedes_Config_Notificaciones.sql (linea 299)
CREATE OR ALTER PROCEDURE dbo.Sp_Reservas_RecordatoriosPendientes
    @FechaHoraActual DATETIME2
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT
            r.Id AS ReservaId,
            s.NegocioId,
            c.NombresORazonSocial AS Cliente,
            c.Correo,
            s.Nombre AS Sede,
            e.Nombre AS Espacio,
            r.Fecha,
            r.HoraInicio,
            r.HoraFin,
            scn.CorreoNotificacion,
            scn.WhatsappContacto
        FROM dbo.Reservas r
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        INNER JOIN dbo.Clientes c ON c.Id = r.ClienteId
        LEFT JOIN dbo.SedeConfiguracionNotificacion scn ON scn.SedeId = s.Id
        WHERE r.Estado IN (1, 2)
          AND r.RecordatorioEnviado = 0
          AND c.Correo IS NOT NULL
          AND LTRIM(RTRIM(c.Correo)) <> N''
          AND COALESCE(scn.NotificacionesActivas, 1) = 1
          AND @FechaHoraActual >= DATEADD(
                MINUTE,
                -COALESCE(scn.MinutosAnticipacionRecordatorio, 90),
                DATEADD(MINUTE, DATEDIFF(MINUTE, 0, r.HoraInicio), CAST(r.Fecha AS DATETIME2))
          )
          AND @FechaHoraActual <= DATEADD(MINUTE, DATEDIFF(MINUTE, 0, r.HoraInicio), CAST(r.Fecha AS DATETIME2));
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

-- SOURCE: 18_Sedes_Config_Notificaciones.sql (linea 342)
CREATE OR ALTER PROCEDURE dbo.Sp_Reservas_AutoNoShow
    @FechaHoraActual DATETIME2,
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        DECLARE @Actualizadas TABLE
        (
            ReservaId INT NOT NULL,
            NegocioId INT NOT NULL
        );

        UPDATE r
        SET r.Estado = 6,
            r.FechaActualizacion = SYSUTCDATETIME(),
            r.UsuarioActualizacion = @Usuario
        OUTPUT inserted.Id, s.NegocioId
        INTO @Actualizadas (ReservaId, NegocioId)
        FROM dbo.Reservas r
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        LEFT JOIN dbo.SedeConfiguracionNotificacion scn ON scn.SedeId = s.Id
        WHERE r.Estado IN (1, 2)
          AND COALESCE(scn.NotificacionesActivas, 1) = 1
          AND DATEADD(
                MINUTE,
                COALESCE(scn.MinutosToleranciaNoShow, 30),
                DATEADD(MINUTE, DATEDIFF(MINUTE, 0, r.HoraInicio), CAST(r.Fecha AS DATETIME2))
          ) <= @FechaHoraActual;

        DECLARE @ReservaId INT, @NegocioId INT;
        DECLARE c CURSOR LOCAL FAST_FORWARD FOR
            SELECT a.ReservaId, a.NegocioId
            FROM @Actualizadas a;

        OPEN c;
        FETCH NEXT FROM c INTO @ReservaId, @NegocioId;

        WHILE @@FETCH_STATUS = 0
        BEGIN
            DECLARE @EntidadIdAudit NVARCHAR(80);
            SET @EntidadIdAudit = CONVERT(NVARCHAR(80), @ReservaId);
            EXEC dbo.Sp_Auditoria_Registrar
                @NegocioId = @NegocioId,
                @Modulo = N'RESERVAS',
                @Accion = N'AUTO_NOSHOW',
                @Entidad = N'Reserva',
                @EntidadId = @EntidadIdAudit,
                @Usuario = @Usuario,
                @DetalleJson = NULL;

            FETCH NEXT FROM c INTO @ReservaId, @NegocioId;
        END;

        CLOSE c;
        DEALLOCATE c;

        SELECT COUNT(1) AS TotalActualizadas FROM @Actualizadas;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

-- SOURCE: 19_Home_Whatsapp_Publico.sql (linea 7)
CREATE OR ALTER PROCEDURE dbo.Sp_Home_ListarSedesPublicas
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT
            s.Id,
            s.Nombre,
            s.Direccion,
            s.Telefono,
            scn.WhatsappContacto,
            COALESCE(scn.PermiteChatWhatsapp, 0) AS PermiteChatWhatsapp
        FROM dbo.Sedes s
        INNER JOIN dbo.Negocios n ON n.Id = s.NegocioId
        LEFT JOIN dbo.SedeConfiguracionNotificacion scn ON scn.SedeId = s.Id
        WHERE s.Activo = 1
          AND n.Activo = 1
        ORDER BY s.Nombre;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

-- SOURCE: 20_Home_Espacios_Whatsapp.sql (linea 7)
CREATE OR ALTER PROCEDURE dbo.Sp_Home_BuscarEspaciosDisponibles
    @Fecha DATE,
    @HoraInicio TIME,
    @HoraFin TIME,
    @SedeId INT = NULL,
    @TipoDeporteId INT = NULL
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT
            e.Id,
            e.Nombre,
            e.Codigo,
            s.Nombre AS SedeNombre,
            td.Nombre AS TipoDeporte,
            e.TieneIluminacion,
            e.Techada,
            scn.WhatsappContacto,
            COALESCE(scn.PermiteChatWhatsapp, 0) AS PermiteChatWhatsapp
        FROM dbo.EspaciosDeportivos e
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        INNER JOIN dbo.TiposDeporte td ON td.Id = e.TipoDeporteId
        LEFT JOIN dbo.SedeConfiguracionNotificacion scn ON scn.SedeId = s.Id
        WHERE e.Estado = 1
          AND s.Activo = 1
          AND (@SedeId IS NULL OR e.SedeId = @SedeId)
          AND (@TipoDeporteId IS NULL OR e.TipoDeporteId = @TipoDeporteId)
          AND NOT EXISTS
          (
              SELECT 1
              FROM dbo.Reservas r
              WHERE r.EspacioDeportivoId = e.Id
                AND r.Fecha = @Fecha
                AND r.Estado NOT IN (5, 6)
                AND @HoraInicio < r.HoraFin
                AND @HoraFin > r.HoraInicio
          )
        ORDER BY s.Nombre, e.Nombre;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

-- SOURCE: 21_Altas_Clubes.sql (linea 41)
CREATE OR ALTER PROCEDURE dbo.Sp_Home_SolicitarAltaClub
    @NombreContacto NVARCHAR(200),
    @Telefono NVARCHAR(30),
    @Correo NVARCHAR(200),
    @RelacionClub NVARCHAR(80),
    @NombreClub NVARCHAR(200),
    @Pais NVARCHAR(80),
    @ProvinciaEstado NVARCHAR(120),
    @Ciudad NVARCHAR(120),
    @Direccion NVARCHAR(250)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        DECLARE @Secuencia INT;
        DECLARE @Codigo NVARCHAR(30);

        SELECT @Secuencia = COUNT(1) + 1
        FROM dbo.SolicitudesAltaClub
        WHERE CAST(FechaRegistro AS DATE) = CAST(SYSUTCDATETIME() AS DATE);

        SET @Codigo = CONCAT(
            N'CLUB-',
            CONVERT(NVARCHAR(8), CAST(SYSUTCDATETIME() AS DATE), 112),
            N'-',
            RIGHT(CONCAT(N'0000', CONVERT(NVARCHAR(10), @Secuencia)), 4)
        );

        INSERT INTO dbo.SolicitudesAltaClub
        (
            CodigoSolicitud, NombreContacto, Telefono, Correo, RelacionClub, NombreClub,
            Pais, ProvinciaEstado, Ciudad, Direccion, Estado, FechaRegistro
        )
        VALUES
        (
            @Codigo, @NombreContacto, @Telefono, @Correo, @RelacionClub, @NombreClub,
            @Pais, @ProvinciaEstado, @Ciudad, @Direccion, 1, SYSUTCDATETIME()
        );

        SELECT @Codigo;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

-- SOURCE: 21_Altas_Clubes.sql (linea 90)
CREATE OR ALTER PROCEDURE dbo.Sp_AltasClubes_Listar
    @Estado INT = NULL
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT
            ac.Id,
            ac.CodigoSolicitud,
            ac.NombreContacto,
            ac.Telefono,
            ac.Correo,
            ac.RelacionClub,
            ac.NombreClub,
            ac.Pais,
            ac.ProvinciaEstado,
            ac.Ciudad,
            ac.Direccion,
            ac.Estado,
            ac.ComentarioGestion,
            ac.NegocioId,
            ac.SedeId,
            ac.FechaRegistro,
            ac.FechaGestion
        FROM dbo.SolicitudesAltaClub ac
        WHERE (@Estado IS NULL OR ac.Estado = @Estado)
        ORDER BY ac.FechaRegistro DESC;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

-- SOURCE: 21_Altas_Clubes.sql (linea 126)
CREATE OR ALTER PROCEDURE dbo.Sp_AltasClubes_Aprobar
    @Id INT,
    @Usuario NVARCHAR(200),
    @ComentarioGestion NVARCHAR(300) = NULL
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        DECLARE @Correo NVARCHAR(200), @NombreClub NVARCHAR(200), @Telefono NVARCHAR(30), @Direccion NVARCHAR(250), @Ciudad NVARCHAR(120), @EstadoActual INT;
        DECLARE @NegocioId INT, @SedeId INT;

        SELECT
            @Correo = ac.Correo,
            @NombreClub = ac.NombreClub,
            @Telefono = ac.Telefono,
            @Direccion = ac.Direccion,
            @Ciudad = ac.Ciudad,
            @EstadoActual = ac.Estado
        FROM dbo.SolicitudesAltaClub ac
        WHERE ac.Id = @Id;

        IF @EstadoActual IS NULL
            RAISERROR('Solicitud no encontrada.', 16, 1);

        IF @EstadoActual <> 1
            RAISERROR('Solo se pueden aprobar solicitudes pendientes.', 16, 1);

        BEGIN TRANSACTION;

        INSERT INTO dbo.Negocios (NombreComercial, RazonSocial, DocumentoFiscal, Activo, FechaRegistro)
        VALUES (@NombreClub, NULL, NULL, 1, SYSUTCDATETIME());
        SET @NegocioId = SCOPE_IDENTITY();

        INSERT INTO dbo.Sedes (NegocioId, Nombre, Direccion, Telefono, Activo, FechaCreacion, UsuarioCreacion)
        VALUES (@NegocioId, CONCAT(@NombreClub, N' - Principal'), CONCAT(@Ciudad, N' - ', @Direccion), @Telefono, 1, SYSUTCDATETIME(), @Usuario);
        SET @SedeId = SCOPE_IDENTITY();

        IF OBJECT_ID(N'dbo.SedeConfiguracionNotificacion', N'U') IS NOT NULL
        BEGIN
            IF NOT EXISTS (SELECT 1 FROM dbo.SedeConfiguracionNotificacion WHERE SedeId = @SedeId)
            BEGIN
                INSERT INTO dbo.SedeConfiguracionNotificacion
                (
                    SedeId, NotificacionesActivas, MinutosAnticipacionRecordatorio, MinutosToleranciaNoShow,
                    CorreoNotificacion, WhatsappContacto, PermiteChatWhatsapp, FechaCreacion, UsuarioCreacion
                )
                VALUES
                (
                    @SedeId, 1, 90, 30, @Correo, NULL, 0, SYSUTCDATETIME(), @Usuario
                );
            END;
        END;

        DECLARE @UsuarioId NVARCHAR(450);
        SELECT TOP (1) @UsuarioId = u.Id
        FROM dbo.AspNetUsers u
        WHERE u.NormalizedEmail = UPPER(@Correo);

        IF @UsuarioId IS NOT NULL
        BEGIN
            IF NOT EXISTS (SELECT 1 FROM dbo.UsuariosNegocio WHERE UsuarioId = @UsuarioId AND NegocioId = @NegocioId)
            BEGIN
                INSERT INTO dbo.UsuariosNegocio (UsuarioId, NegocioId, RolNegocio, Activo)
                VALUES (@UsuarioId, @NegocioId, 1, 1);
            END;
        END;

        UPDATE dbo.SolicitudesAltaClub
        SET Estado = 2,
            ComentarioGestion = @ComentarioGestion,
            NegocioId = @NegocioId,
            SedeId = @SedeId,
            FechaGestion = SYSUTCDATETIME(),
            UsuarioGestion = @Usuario
        WHERE Id = @Id;

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

-- SOURCE: 21_Altas_Clubes.sql (linea 215)
CREATE OR ALTER PROCEDURE dbo.Sp_AltasClubes_Rechazar
    @Id INT,
    @Usuario NVARCHAR(200),
    @ComentarioGestion NVARCHAR(300) = NULL
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        UPDATE dbo.SolicitudesAltaClub
        SET Estado = 3,
            ComentarioGestion = @ComentarioGestion,
            FechaGestion = SYSUTCDATETIME(),
            UsuarioGestion = @Usuario
        WHERE Id = @Id
          AND Estado = 1;

        IF @@ROWCOUNT = 0
            RAISERROR('No se encontro la solicitud pendiente para rechazar.', 16, 1);
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

-- SOURCE: 22_Registro_Club_Prueba.sql (linea 30)
CREATE OR ALTER PROCEDURE dbo.Sp_Home_RegistrarClubConPrueba
    @UsuarioId NVARCHAR(450),
    @NombreContacto NVARCHAR(200),
    @Telefono NVARCHAR(30),
    @Correo NVARCHAR(200),
    @RelacionClub NVARCHAR(80),
    @NombreClub NVARCHAR(200),
    @Pais NVARCHAR(80),
    @ProvinciaEstado NVARCHAR(120),
    @Ciudad NVARCHAR(120),
    @Direccion NVARCHAR(250)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        DECLARE @NegocioId INT;
        DECLARE @SedeId INT;
        DECLARE @CodigoSolicitud NVARCHAR(30);
        DECLARE @Secuencia INT;

        IF NOT EXISTS (SELECT 1 FROM dbo.AspNetUsers u WHERE u.Id = @UsuarioId)
            RAISERROR('Usuario invalido.', 16, 1);

        IF EXISTS (SELECT 1 FROM dbo.UsuariosNegocio un WHERE un.UsuarioId = @UsuarioId AND un.Activo = 1)
            RAISERROR('El usuario ya tiene un negocio asociado. Solo se permite el alta inicial.', 16, 1);

        BEGIN TRANSACTION;

        INSERT INTO dbo.Negocios (NombreComercial, RazonSocial, DocumentoFiscal, Activo, FechaRegistro)
        VALUES (@NombreClub, NULL, NULL, 1, SYSUTCDATETIME());
        SET @NegocioId = SCOPE_IDENTITY();

        INSERT INTO dbo.Sedes (NegocioId, Nombre, Direccion, Telefono, Activo, FechaCreacion, UsuarioCreacion)
        VALUES
        (
            @NegocioId,
            CONCAT(@NombreClub, N' - Principal'),
            CONCAT(@Pais, N', ', @ProvinciaEstado, N', ', @Ciudad, N' - ', @Direccion),
            @Telefono,
            1,
            SYSUTCDATETIME(),
            @Correo
        );
        SET @SedeId = SCOPE_IDENTITY();

        INSERT INTO dbo.UsuariosNegocio (UsuarioId, NegocioId, RolNegocio, Activo)
        VALUES (@UsuarioId, @NegocioId, 1, 1);

        IF NOT EXISTS (SELECT 1 FROM dbo.NegociosSuscripcion WHERE NegocioId = @NegocioId)
        BEGIN
            INSERT INTO dbo.NegociosSuscripcion
            (
                NegocioId, EstadoSuscripcion, EsPrueba,
                FechaInicioPrueba, FechaFinPrueba,
                FechaInicioPlan, FechaFinPlan,
                FechaCreacion, UsuarioCreacion
            )
            VALUES
            (
                @NegocioId, 1, 1,
                CAST(SYSUTCDATETIME() AS DATE),
                DATEADD(DAY, 30, CAST(SYSUTCDATETIME() AS DATE)),
                NULL, NULL,
                SYSUTCDATETIME(),
                @Correo
            );
        END;

        IF OBJECT_ID(N'dbo.SedeConfiguracionNotificacion', N'U') IS NOT NULL
        BEGIN
            IF NOT EXISTS (SELECT 1 FROM dbo.SedeConfiguracionNotificacion WHERE SedeId = @SedeId)
            BEGIN
                INSERT INTO dbo.SedeConfiguracionNotificacion
                (
                    SedeId, NotificacionesActivas, MinutosAnticipacionRecordatorio, MinutosToleranciaNoShow,
                    CorreoNotificacion, WhatsappContacto, PermiteChatWhatsapp, FechaCreacion, UsuarioCreacion
                )
                VALUES
                (
                    @SedeId, 1, 90, 30, @Correo, NULL, 0, SYSUTCDATETIME(), @Correo
                );
            END;
        END;

        IF OBJECT_ID(N'dbo.SolicitudesAltaClub', N'U') IS NOT NULL
        BEGIN
            SELECT @Secuencia = COUNT(1) + 1
            FROM dbo.SolicitudesAltaClub
            WHERE CAST(FechaRegistro AS DATE) = CAST(SYSUTCDATETIME() AS DATE);

            SET @CodigoSolicitud = CONCAT(
                N'CLUB-',
                CONVERT(NVARCHAR(8), CAST(SYSUTCDATETIME() AS DATE), 112),
                N'-',
                RIGHT(CONCAT(N'0000', CONVERT(NVARCHAR(10), @Secuencia)), 4)
            );

            INSERT INTO dbo.SolicitudesAltaClub
            (
                CodigoSolicitud, NombreContacto, Telefono, Correo, RelacionClub, NombreClub,
                Pais, ProvinciaEstado, Ciudad, Direccion, Estado, ComentarioGestion,
                NegocioId, SedeId, FechaRegistro, FechaGestion, UsuarioGestion
            )
            VALUES
            (
                @CodigoSolicitud, @NombreContacto, @Telefono, @Correo, @RelacionClub, @NombreClub,
                @Pais, @ProvinciaEstado, @Ciudad, @Direccion, 2, N'Autoaprobada por registro directo.',
                @NegocioId, @SedeId, SYSUTCDATETIME(), SYSUTCDATETIME(), @Correo
            );
        END;
        ELSE
        BEGIN
            SET @CodigoSolicitud = CONCAT(N'ALTA-', CONVERT(NVARCHAR(8), CAST(SYSUTCDATETIME() AS DATE), 112), N'-', CONVERT(NVARCHAR(20), @NegocioId));
        END;

        IF OBJECT_ID(N'dbo.Sp_Auditoria_Registrar', N'P') IS NOT NULL
        BEGIN
            DECLARE @EntidadIdAudit NVARCHAR(80) = CONVERT(NVARCHAR(80), @NegocioId);
            EXEC dbo.Sp_Auditoria_Registrar
                @NegocioId = @NegocioId,
                @Modulo = N'ALTAS_CLUBES',
                @Accion = N'CREATE',
                @Entidad = N'Negocio',
                @EntidadId = @EntidadIdAudit,
                @Usuario = @Correo,
                @DetalleJson = NULL;
        END;

        COMMIT TRANSACTION;
        SELECT @CodigoSolicitud AS CodigoRegistro;
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

-- SOURCE: 23_Suscripcion_Bloqueo_Operacion.sql (linea 136)
CREATE OR ALTER PROCEDURE dbo.Sp_Panel_ListarModulosPermitidos
    @UsuarioId NVARCHAR(450),
    @NegocioId INT
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        DECLARE @UsuarioNegocioId INT, @RolNegocio INT;
        DECLARE @EstadoSuscripcion INT, @EsPrueba BIT, @FechaFinPrueba DATE, @FechaFinPlan DATE;
        DECLARE @Hoy DATE = CAST(SYSUTCDATETIME() AS DATE);

        SELECT @UsuarioNegocioId = un.Id, @RolNegocio = un.RolNegocio
        FROM dbo.UsuariosNegocio un
        WHERE un.UsuarioId = @UsuarioId
          AND un.NegocioId = @NegocioId
          AND un.Activo = 1;

        IF OBJECT_ID(N'dbo.NegociosSuscripcion', N'U') IS NOT NULL
        BEGIN
            SELECT
                @EstadoSuscripcion = ns.EstadoSuscripcion,
                @EsPrueba = ns.EsPrueba,
                @FechaFinPrueba = ns.FechaFinPrueba,
                @FechaFinPlan = ns.FechaFinPlan
            FROM dbo.NegociosSuscripcion ns
            WHERE ns.NegocioId = @NegocioId;

            IF @EstadoSuscripcion = 1 AND @EsPrueba = 1 AND @FechaFinPrueba IS NOT NULL AND @FechaFinPrueba < @Hoy
            BEGIN
                UPDATE dbo.NegociosSuscripcion
                SET EstadoSuscripcion = 3,
                    EsPrueba = 0,
                    FechaActualizacion = SYSUTCDATETIME(),
                    UsuarioActualizacion = @UsuarioId
                WHERE NegocioId = @NegocioId;
                SET @EstadoSuscripcion = 3;
            END;

            IF @EstadoSuscripcion = 2 AND @FechaFinPlan IS NOT NULL AND @FechaFinPlan < @Hoy
            BEGIN
                UPDATE dbo.NegociosSuscripcion
                SET EstadoSuscripcion = 3,
                    FechaActualizacion = SYSUTCDATETIME(),
                    UsuarioActualizacion = @UsuarioId
                WHERE NegocioId = @NegocioId;
                SET @EstadoSuscripcion = 3;
            END;

            IF @EstadoSuscripcion IN (3, 4)
            BEGIN
                SELECT
                    m.Id,
                    m.Codigo,
                    m.Nombre,
                    CAST(0 AS BIT) AS PuedeVer,
                    CAST(0 AS BIT) AS PuedeCrear,
                    CAST(0 AS BIT) AS PuedeEditar,
                    CAST(0 AS BIT) AS PuedeEliminar
                FROM dbo.ModulosSistema m
                WHERE 1 = 0;
                RETURN;
            END;
        END;

        SELECT
            m.Id,
            m.Codigo,
            m.Nombre,
            COALESCE(up.PuedeVer, rp.PuedeVer) AS PuedeVer,
            COALESCE(up.PuedeCrear, rp.PuedeCrear) AS PuedeCrear,
            COALESCE(up.PuedeEditar, rp.PuedeEditar) AS PuedeEditar,
            COALESCE(up.PuedeEliminar, rp.PuedeEliminar) AS PuedeEliminar
        FROM dbo.ModulosSistema m
        INNER JOIN dbo.RolesNegocioPermiso rp ON rp.ModuloSistemaId = m.Id AND rp.RolNegocio = @RolNegocio
        LEFT JOIN dbo.UsuariosNegocioPermiso up ON up.ModuloSistemaId = m.Id AND up.UsuarioNegocioId = @UsuarioNegocioId
        WHERE m.Activo = 1
        ORDER BY m.Nombre;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

-- SOURCE: 23_Suscripcion_Bloqueo_Operacion.sql (linea 222)
CREATE OR ALTER PROCEDURE dbo.Sp_NegociosSuscripcion_ActivarPlan
    @NegocioId INT,
    @DiasVigencia INT = 30,
    @Usuario NVARCHAR(200) = NULL
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF @DiasVigencia IS NULL OR @DiasVigencia <= 0
            SET @DiasVigencia = 30;

        IF NOT EXISTS (SELECT 1 FROM dbo.Negocios WHERE Id = @NegocioId)
            RAISERROR('Negocio no encontrado.', 16, 1);

        IF OBJECT_ID(N'dbo.NegociosSuscripcion', N'U') IS NULL
            RAISERROR('No existe la tabla NegociosSuscripcion. Ejecuta primero 22_Registro_Club_Prueba.sql.', 16, 1);

        IF EXISTS (SELECT 1 FROM dbo.NegociosSuscripcion WHERE NegocioId = @NegocioId)
        BEGIN
            UPDATE dbo.NegociosSuscripcion
            SET EstadoSuscripcion = 2,
                EsPrueba = 0,
                FechaInicioPlan = CAST(SYSUTCDATETIME() AS DATE),
                FechaFinPlan = DATEADD(DAY, @DiasVigencia, CAST(SYSUTCDATETIME() AS DATE)),
                FechaActualizacion = SYSUTCDATETIME(),
                UsuarioActualizacion = COALESCE(@Usuario, N'sistema')
            WHERE NegocioId = @NegocioId;
        END
        ELSE
        BEGIN
            INSERT INTO dbo.NegociosSuscripcion
            (
                NegocioId, EstadoSuscripcion, EsPrueba, FechaInicioPrueba, FechaFinPrueba,
                FechaInicioPlan, FechaFinPlan, FechaCreacion, UsuarioCreacion
            )
            VALUES
            (
                @NegocioId, 2, 0, NULL, NULL,
                CAST(SYSUTCDATETIME() AS DATE),
                DATEADD(DAY, @DiasVigencia, CAST(SYSUTCDATETIME() AS DATE)),
                SYSUTCDATETIME(), COALESCE(@Usuario, N'sistema')
            );
        END;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

-- SOURCE: 24_Sedes_Horario_NoLaborable.sql (linea 100)
CREATE OR ALTER PROCEDURE dbo.Sp_Sedes_ObtenerPorId
    @NegocioId INT,
    @Id INT
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT
            s.Id, s.NegocioId, s.Nombre, s.Direccion, s.Telefono, s.Activo,
            STUFF((SELECT N',' + CONVERT(NVARCHAR(20), ss.ServicioId) FROM dbo.SedeServicios ss WHERE ss.SedeId = s.Id ORDER BY ss.ServicioId FOR XML PATH(''), TYPE).value('.', 'NVARCHAR(MAX)'), 1, 1, N'') AS ServiciosIdsCsv,
            COALESCE(scn.NotificacionesActivas, 1) AS NotificacionesActivas,
            COALESCE(scn.MinutosAnticipacionRecordatorio, 90) AS MinutosAnticipacionRecordatorio,
            COALESCE(scn.MinutosToleranciaNoShow, 30) AS MinutosToleranciaNoShow,
            scn.CorreoNotificacion,
            scn.WhatsappContacto,
            COALESCE(scn.PermiteChatWhatsapp, 0) AS PermiteChatWhatsapp,
            COALESCE(sha.AtiendeLunes, 1) AS AtiendeLunes,
            COALESCE(sha.AtiendeMartes, 1) AS AtiendeMartes,
            COALESCE(sha.AtiendeMiercoles, 1) AS AtiendeMiercoles,
            COALESCE(sha.AtiendeJueves, 1) AS AtiendeJueves,
            COALESCE(sha.AtiendeViernes, 1) AS AtiendeViernes,
            COALESCE(sha.AtiendeSabado, 1) AS AtiendeSabado,
            COALESCE(sha.AtiendeDomingo, 1) AS AtiendeDomingo,
            COALESCE(sha.HoraApertura, CAST('08:00' AS TIME)) AS HoraApertura,
            COALESCE(sha.HoraCierre, CAST('23:00' AS TIME)) AS HoraCierre,
            STUFF((SELECT N',' + CONVERT(NVARCHAR(10), sfi.Fecha, 23) FROM dbo.SedeFechasInhabilitadas sfi WHERE sfi.SedeId = s.Id AND sfi.Activo = 1 ORDER BY sfi.Fecha FOR XML PATH(''), TYPE).value('.', 'NVARCHAR(MAX)'), 1, 1, N'') AS FechasInhabilitadasCsv
        FROM dbo.Sedes s
        LEFT JOIN dbo.SedeConfiguracionNotificacion scn ON scn.SedeId = s.Id
        LEFT JOIN dbo.SedeHorarioAtencion sha ON sha.SedeId = s.Id
        WHERE s.NegocioId = @NegocioId
          AND s.Id = @Id;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

-- SOURCE: 25_Sedes_Horario_Crear_Actualizar.sql (linea 8)
CREATE OR ALTER PROCEDURE dbo.Sp_Sedes_Crear
    @NegocioId INT,
    @Nombre NVARCHAR(150),
    @Direccion NVARCHAR(250),
    @Telefono NVARCHAR(20) = NULL,
    @Activo BIT,
    @ServiciosIdsCsv NVARCHAR(MAX) = NULL,
    @NotificacionesActivas BIT = 1,
    @MinutosAnticipacionRecordatorio INT = 90,
    @MinutosToleranciaNoShow INT = 30,
    @CorreoNotificacion NVARCHAR(200) = NULL,
    @WhatsappContacto NVARCHAR(20) = NULL,
    @PermiteChatWhatsapp BIT = 0,
    @AtiendeLunes BIT = 1,
    @AtiendeMartes BIT = 1,
    @AtiendeMiercoles BIT = 1,
    @AtiendeJueves BIT = 1,
    @AtiendeViernes BIT = 1,
    @AtiendeSabado BIT = 1,
    @AtiendeDomingo BIT = 1,
    @HoraApertura TIME = '08:00',
    @HoraCierre TIME = '23:00',
    @FechasInhabilitadasCsv NVARCHAR(MAX) = NULL,
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF @HoraCierre <= @HoraApertura
            RAISERROR('La hora de cierre debe ser mayor a la hora de apertura.', 16, 1);
        IF COALESCE(@AtiendeLunes, 0) + COALESCE(@AtiendeMartes, 0) + COALESCE(@AtiendeMiercoles, 0) + COALESCE(@AtiendeJueves, 0) + COALESCE(@AtiendeViernes, 0) + COALESCE(@AtiendeSabado, 0) + COALESCE(@AtiendeDomingo, 0) = 0
            RAISERROR('Debes seleccionar al menos un dia de atencion.', 16, 1);

        BEGIN TRANSACTION;

        INSERT INTO dbo.Sedes (NegocioId, Nombre, Direccion, Telefono, Activo, FechaCreacion, UsuarioCreacion)
        VALUES (@NegocioId, @Nombre, @Direccion, @Telefono, @Activo, SYSUTCDATETIME(), @Usuario);

        DECLARE @Id INT = SCOPE_IDENTITY();

        INSERT INTO dbo.SedeConfiguracionNotificacion
        (
            SedeId, NotificacionesActivas, MinutosAnticipacionRecordatorio, MinutosToleranciaNoShow,
            CorreoNotificacion, WhatsappContacto, PermiteChatWhatsapp, FechaCreacion, UsuarioCreacion
        )
        VALUES
        (
            @Id, @NotificacionesActivas, @MinutosAnticipacionRecordatorio, @MinutosToleranciaNoShow,
            @CorreoNotificacion, @WhatsappContacto, @PermiteChatWhatsapp, SYSUTCDATETIME(), @Usuario
        );

        MERGE dbo.SedeHorarioAtencion AS tgt
        USING (SELECT @Id AS SedeId) AS src
            ON tgt.SedeId = src.SedeId
        WHEN MATCHED THEN
            UPDATE SET
                AtiendeLunes = @AtiendeLunes,
                AtiendeMartes = @AtiendeMartes,
                AtiendeMiercoles = @AtiendeMiercoles,
                AtiendeJueves = @AtiendeJueves,
                AtiendeViernes = @AtiendeViernes,
                AtiendeSabado = @AtiendeSabado,
                AtiendeDomingo = @AtiendeDomingo,
                HoraApertura = @HoraApertura,
                HoraCierre = @HoraCierre,
                FechaActualizacion = SYSUTCDATETIME(),
                UsuarioActualizacion = @Usuario
        WHEN NOT MATCHED THEN
            INSERT
            (
                SedeId, AtiendeLunes, AtiendeMartes, AtiendeMiercoles, AtiendeJueves, AtiendeViernes, AtiendeSabado, AtiendeDomingo,
                HoraApertura, HoraCierre, FechaCreacion, UsuarioCreacion
            )
            VALUES
            (
                @Id, @AtiendeLunes, @AtiendeMartes, @AtiendeMiercoles, @AtiendeJueves, @AtiendeViernes, @AtiendeSabado, @AtiendeDomingo,
                @HoraApertura, @HoraCierre, SYSUTCDATETIME(), @Usuario
            );

        IF @ServiciosIdsCsv IS NOT NULL AND LEN(LTRIM(RTRIM(@ServiciosIdsCsv))) > 0
        BEGIN
            ;WITH Servicios AS
            (
                SELECT DISTINCT TRY_CONVERT(INT, LTRIM(RTRIM(value))) AS ServicioId
                FROM STRING_SPLIT(@ServiciosIdsCsv, N',')
                WHERE TRY_CONVERT(INT, LTRIM(RTRIM(value))) IS NOT NULL
            )
            INSERT INTO dbo.SedeServicios (SedeId, ServicioId, FechaRegistro, UsuarioCreacion)
            SELECT @Id, s.ServicioId, SYSUTCDATETIME(), @Usuario
            FROM Servicios s
            INNER JOIN dbo.CatalogoServiciosSede cs ON cs.Id = s.ServicioId
            WHERE cs.Activo = 1;
        END;

        IF @FechasInhabilitadasCsv IS NOT NULL AND LEN(LTRIM(RTRIM(@FechasInhabilitadasCsv))) > 0
        BEGIN
            ;WITH Fechas AS
            (
                SELECT DISTINCT TRY_CONVERT(DATE, LTRIM(RTRIM(value))) AS Fecha
                FROM STRING_SPLIT(@FechasInhabilitadasCsv, N',')
                WHERE TRY_CONVERT(DATE, LTRIM(RTRIM(value))) IS NOT NULL
            )
            INSERT INTO dbo.SedeFechasInhabilitadas (SedeId, Fecha, Activo, FechaCreacion, UsuarioCreacion)
            SELECT @Id, f.Fecha, 1, SYSUTCDATETIME(), @Usuario
            FROM Fechas f;
        END;

        DECLARE @EntidadIdAudit NVARCHAR(80);
        SET @EntidadIdAudit = CONVERT(NVARCHAR(80), @Id);
        EXEC dbo.Sp_Auditoria_Registrar @NegocioId = @NegocioId, @Modulo = N'SEDES', @Accion = N'CREATE', @Entidad = N'Sede', @EntidadId = @EntidadIdAudit, @Usuario = @Usuario, @DetalleJson = NULL;

        COMMIT TRANSACTION;
        SELECT @Id;
    END TRY
    BEGIN CATCH
        IF XACT_STATE() <> 0 ROLLBACK TRANSACTION;
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

-- SOURCE: 25_Sedes_Horario_Crear_Actualizar.sql (linea 131)
CREATE OR ALTER PROCEDURE dbo.Sp_Sedes_Actualizar
    @Id INT,
    @NegocioId INT,
    @Nombre NVARCHAR(150),
    @Direccion NVARCHAR(250),
    @Telefono NVARCHAR(20) = NULL,
    @Activo BIT,
    @ServiciosIdsCsv NVARCHAR(MAX) = NULL,
    @NotificacionesActivas BIT = 1,
    @MinutosAnticipacionRecordatorio INT = 90,
    @MinutosToleranciaNoShow INT = 30,
    @CorreoNotificacion NVARCHAR(200) = NULL,
    @WhatsappContacto NVARCHAR(20) = NULL,
    @PermiteChatWhatsapp BIT = 0,
    @AtiendeLunes BIT = 1,
    @AtiendeMartes BIT = 1,
    @AtiendeMiercoles BIT = 1,
    @AtiendeJueves BIT = 1,
    @AtiendeViernes BIT = 1,
    @AtiendeSabado BIT = 1,
    @AtiendeDomingo BIT = 1,
    @HoraApertura TIME = '08:00',
    @HoraCierre TIME = '23:00',
    @FechasInhabilitadasCsv NVARCHAR(MAX) = NULL,
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF @HoraCierre <= @HoraApertura
            RAISERROR('La hora de cierre debe ser mayor a la hora de apertura.', 16, 1);
        IF COALESCE(@AtiendeLunes, 0) + COALESCE(@AtiendeMartes, 0) + COALESCE(@AtiendeMiercoles, 0) + COALESCE(@AtiendeJueves, 0) + COALESCE(@AtiendeViernes, 0) + COALESCE(@AtiendeSabado, 0) + COALESCE(@AtiendeDomingo, 0) = 0
            RAISERROR('Debes seleccionar al menos un dia de atencion.', 16, 1);

        BEGIN TRANSACTION;

        UPDATE dbo.Sedes
        SET Nombre = @Nombre,
            Direccion = @Direccion,
            Telefono = @Telefono,
            Activo = @Activo,
            FechaActualizacion = SYSUTCDATETIME(),
            UsuarioActualizacion = @Usuario
        WHERE Id = @Id
          AND NegocioId = @NegocioId;

        IF @@ROWCOUNT = 0
        BEGIN
            ROLLBACK TRANSACTION;
            RETURN;
        END;

        MERGE dbo.SedeConfiguracionNotificacion AS tgt
        USING (SELECT @Id AS SedeId) AS src ON tgt.SedeId = src.SedeId
        WHEN MATCHED THEN UPDATE SET
            NotificacionesActivas = @NotificacionesActivas,
            MinutosAnticipacionRecordatorio = @MinutosAnticipacionRecordatorio,
            MinutosToleranciaNoShow = @MinutosToleranciaNoShow,
            CorreoNotificacion = @CorreoNotificacion,
            WhatsappContacto = @WhatsappContacto,
            PermiteChatWhatsapp = @PermiteChatWhatsapp,
            FechaActualizacion = SYSUTCDATETIME(),
            UsuarioActualizacion = @Usuario
        WHEN NOT MATCHED THEN
            INSERT (SedeId, NotificacionesActivas, MinutosAnticipacionRecordatorio, MinutosToleranciaNoShow, CorreoNotificacion, WhatsappContacto, PermiteChatWhatsapp, FechaCreacion, UsuarioCreacion)
            VALUES (@Id, @NotificacionesActivas, @MinutosAnticipacionRecordatorio, @MinutosToleranciaNoShow, @CorreoNotificacion, @WhatsappContacto, @PermiteChatWhatsapp, SYSUTCDATETIME(), @Usuario);

        MERGE dbo.SedeHorarioAtencion AS tgt
        USING (SELECT @Id AS SedeId) AS src ON tgt.SedeId = src.SedeId
        WHEN MATCHED THEN UPDATE SET
            AtiendeLunes = @AtiendeLunes, AtiendeMartes = @AtiendeMartes, AtiendeMiercoles = @AtiendeMiercoles, AtiendeJueves = @AtiendeJueves,
            AtiendeViernes = @AtiendeViernes, AtiendeSabado = @AtiendeSabado, AtiendeDomingo = @AtiendeDomingo,
            HoraApertura = @HoraApertura, HoraCierre = @HoraCierre, FechaActualizacion = SYSUTCDATETIME(), UsuarioActualizacion = @Usuario
        WHEN NOT MATCHED THEN
            INSERT (SedeId, AtiendeLunes, AtiendeMartes, AtiendeMiercoles, AtiendeJueves, AtiendeViernes, AtiendeSabado, AtiendeDomingo, HoraApertura, HoraCierre, FechaCreacion, UsuarioCreacion)
            VALUES (@Id, @AtiendeLunes, @AtiendeMartes, @AtiendeMiercoles, @AtiendeJueves, @AtiendeViernes, @AtiendeSabado, @AtiendeDomingo, @HoraApertura, @HoraCierre, SYSUTCDATETIME(), @Usuario);

        DELETE FROM dbo.SedeServicios WHERE SedeId = @Id;
        IF @ServiciosIdsCsv IS NOT NULL AND LEN(LTRIM(RTRIM(@ServiciosIdsCsv))) > 0
        BEGIN
            ;WITH Servicios AS
            (
                SELECT DISTINCT TRY_CONVERT(INT, LTRIM(RTRIM(value))) AS ServicioId
                FROM STRING_SPLIT(@ServiciosIdsCsv, N',')
                WHERE TRY_CONVERT(INT, LTRIM(RTRIM(value))) IS NOT NULL
            )
            INSERT INTO dbo.SedeServicios (SedeId, ServicioId, FechaRegistro, UsuarioCreacion)
            SELECT @Id, s.ServicioId, SYSUTCDATETIME(), @Usuario
            FROM Servicios s
            INNER JOIN dbo.CatalogoServiciosSede cs ON cs.Id = s.ServicioId
            WHERE cs.Activo = 1;
        END;

        DELETE FROM dbo.SedeFechasInhabilitadas WHERE SedeId = @Id;
        IF @FechasInhabilitadasCsv IS NOT NULL AND LEN(LTRIM(RTRIM(@FechasInhabilitadasCsv))) > 0
        BEGIN
            ;WITH Fechas AS
            (
                SELECT DISTINCT TRY_CONVERT(DATE, LTRIM(RTRIM(value))) AS Fecha
                FROM STRING_SPLIT(@FechasInhabilitadasCsv, N',')
                WHERE TRY_CONVERT(DATE, LTRIM(RTRIM(value))) IS NOT NULL
            )
            INSERT INTO dbo.SedeFechasInhabilitadas (SedeId, Fecha, Activo, FechaCreacion, UsuarioCreacion)
            SELECT @Id, f.Fecha, 1, SYSUTCDATETIME(), @Usuario
            FROM Fechas f;
        END;

        DECLARE @EntidadIdAudit NVARCHAR(80);
        SET @EntidadIdAudit = CONVERT(NVARCHAR(80), @Id);
        EXEC dbo.Sp_Auditoria_Registrar @NegocioId = @NegocioId, @Modulo = N'SEDES', @Accion = N'EDIT', @Entidad = N'Sede', @EntidadId = @EntidadIdAudit, @Usuario = @Usuario, @DetalleJson = NULL;

        COMMIT TRANSACTION;
    END TRY
    BEGIN CATCH
        IF XACT_STATE() <> 0 ROLLBACK TRANSACTION;
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

-- SOURCE: 26_Reservas_Validacion_Horario_Sede.sql (linea 9)
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
        IF @HoraFin <= @HoraInicio
            RAISERROR('La hora fin debe ser mayor que la hora inicio.', 16, 1);

        DECLARE @SedeId INT, @HoraApertura TIME, @HoraCierre TIME;
        DECLARE @AtiendeLunes BIT, @AtiendeMartes BIT, @AtiendeMiercoles BIT, @AtiendeJueves BIT, @AtiendeViernes BIT, @AtiendeSabado BIT, @AtiendeDomingo BIT;
        DECLARE @DiaSemana INT, @DiaHabilitado BIT;

        SELECT
            @SedeId = s.Id,
            @HoraApertura = COALESCE(sha.HoraApertura, CAST('08:00' AS TIME)),
            @HoraCierre = COALESCE(sha.HoraCierre, CAST('23:00' AS TIME)),
            @AtiendeLunes = COALESCE(sha.AtiendeLunes, 1),
            @AtiendeMartes = COALESCE(sha.AtiendeMartes, 1),
            @AtiendeMiercoles = COALESCE(sha.AtiendeMiercoles, 1),
            @AtiendeJueves = COALESCE(sha.AtiendeJueves, 1),
            @AtiendeViernes = COALESCE(sha.AtiendeViernes, 1),
            @AtiendeSabado = COALESCE(sha.AtiendeSabado, 1),
            @AtiendeDomingo = COALESCE(sha.AtiendeDomingo, 1)
        FROM dbo.EspaciosDeportivos e
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        LEFT JOIN dbo.SedeHorarioAtencion sha ON sha.SedeId = s.Id
        WHERE e.Id = @EspacioDeportivoId
          AND s.NegocioId = @NegocioId
          AND e.Estado = 1;

        IF @SedeId IS NULL
            RAISERROR('El espacio deportivo no esta disponible para este negocio.', 16, 1);

        IF EXISTS (SELECT 1 FROM dbo.SedeFechasInhabilitadas sfi WHERE sfi.SedeId = @SedeId AND sfi.Fecha = @Fecha AND sfi.Activo = 1)
            RAISERROR('La sede no atiende en la fecha seleccionada.', 16, 1);

        SET @DiaSemana = (DATEDIFF(DAY, '19000101', @Fecha) % 7) + 1;
        SET @DiaHabilitado = CASE @DiaSemana
            WHEN 1 THEN @AtiendeLunes
            WHEN 2 THEN @AtiendeMartes
            WHEN 3 THEN @AtiendeMiercoles
            WHEN 4 THEN @AtiendeJueves
            WHEN 5 THEN @AtiendeViernes
            WHEN 6 THEN @AtiendeSabado
            WHEN 7 THEN @AtiendeDomingo
            ELSE 0 END;

        IF COALESCE(@DiaHabilitado, 0) = 0
            RAISERROR('La sede no atiende el dia seleccionado.', 16, 1);
        IF @HoraInicio < @HoraApertura OR @HoraFin > @HoraCierre
            RAISERROR('El horario de reserva esta fuera del horario de atencion de la sede.', 16, 1);

        IF EXISTS (SELECT 1 FROM dbo.Reservas r WHERE r.EspacioDeportivoId = @EspacioDeportivoId AND r.Fecha = @Fecha AND r.Estado NOT IN (5, 6) AND @HoraInicio < r.HoraFin AND @HoraFin > r.HoraInicio)
            RAISERROR('Cruce de horario detectado.', 16, 1);
        IF EXISTS (SELECT 1 FROM dbo.BloqueosHorario b WHERE b.EspacioDeportivoId = @EspacioDeportivoId AND b.Fecha = @Fecha AND b.Activo = 1 AND @HoraInicio < b.HoraFin AND @HoraFin > b.HoraInicio)
            RAISERROR('El horario esta bloqueado para ese espacio.', 16, 1);

        INSERT INTO dbo.Reservas (EspacioDeportivoId, ClienteId, Fecha, HoraInicio, HoraFin, Estado, Total, Adelanto, Saldo, FechaRegistro, UsuarioCreacion)
        VALUES (@EspacioDeportivoId, @ClienteId, @Fecha, @HoraInicio, @HoraFin, @Estado, @Total, @Adelanto, (@Total - @Adelanto), SYSUTCDATETIME(), @Usuario);

        DECLARE @Id INT = SCOPE_IDENTITY();
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

-- SOURCE: 26_Reservas_Validacion_Horario_Sede.sql (linea 93)
CREATE OR ALTER PROCEDURE dbo.Sp_Reservas_Mover
    @NegocioId INT,
    @Id INT,
    @Fecha DATE,
    @HoraInicio TIME,
    @HoraFin TIME,
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF @HoraFin <= @HoraInicio
            RAISERROR('La hora fin debe ser mayor a la hora inicio.', 16, 1);

        DECLARE @EspacioDeportivoId INT, @SedeId INT, @HoraApertura TIME, @HoraCierre TIME;
        DECLARE @AtiendeLunes BIT, @AtiendeMartes BIT, @AtiendeMiercoles BIT, @AtiendeJueves BIT, @AtiendeViernes BIT, @AtiendeSabado BIT, @AtiendeDomingo BIT;
        DECLARE @DiaSemana INT, @DiaHabilitado BIT;

        SELECT
            @EspacioDeportivoId = r.EspacioDeportivoId,
            @SedeId = s.Id,
            @HoraApertura = COALESCE(sha.HoraApertura, CAST('08:00' AS TIME)),
            @HoraCierre = COALESCE(sha.HoraCierre, CAST('23:00' AS TIME)),
            @AtiendeLunes = COALESCE(sha.AtiendeLunes, 1),
            @AtiendeMartes = COALESCE(sha.AtiendeMartes, 1),
            @AtiendeMiercoles = COALESCE(sha.AtiendeMiercoles, 1),
            @AtiendeJueves = COALESCE(sha.AtiendeJueves, 1),
            @AtiendeViernes = COALESCE(sha.AtiendeViernes, 1),
            @AtiendeSabado = COALESCE(sha.AtiendeSabado, 1),
            @AtiendeDomingo = COALESCE(sha.AtiendeDomingo, 1)
        FROM dbo.Reservas r
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        LEFT JOIN dbo.SedeHorarioAtencion sha ON sha.SedeId = s.Id
        WHERE r.Id = @Id
          AND s.NegocioId = @NegocioId
          AND r.Estado NOT IN (5, 6);

        IF @EspacioDeportivoId IS NULL
            RAISERROR('No se encontro la reserva para mover.', 16, 1);

        IF EXISTS (SELECT 1 FROM dbo.SedeFechasInhabilitadas sfi WHERE sfi.SedeId = @SedeId AND sfi.Fecha = @Fecha AND sfi.Activo = 1)
            RAISERROR('La sede no atiende en la fecha seleccionada.', 16, 1);

        SET @DiaSemana = (DATEDIFF(DAY, '19000101', @Fecha) % 7) + 1;
        SET @DiaHabilitado = CASE @DiaSemana
            WHEN 1 THEN @AtiendeLunes
            WHEN 2 THEN @AtiendeMartes
            WHEN 3 THEN @AtiendeMiercoles
            WHEN 4 THEN @AtiendeJueves
            WHEN 5 THEN @AtiendeViernes
            WHEN 6 THEN @AtiendeSabado
            WHEN 7 THEN @AtiendeDomingo
            ELSE 0 END;

        IF COALESCE(@DiaHabilitado, 0) = 0
            RAISERROR('La sede no atiende el dia seleccionado.', 16, 1);
        IF @HoraInicio < @HoraApertura OR @HoraFin > @HoraCierre
            RAISERROR('El horario de reserva esta fuera del horario de atencion de la sede.', 16, 1);

        IF EXISTS (SELECT 1 FROM dbo.Reservas r WHERE r.EspacioDeportivoId = @EspacioDeportivoId AND r.Fecha = @Fecha AND r.Estado NOT IN (5, 6) AND r.Id <> @Id AND @HoraInicio < r.HoraFin AND @HoraFin > r.HoraInicio)
            RAISERROR('Cruce de horario con otra reserva.', 16, 1);
        IF EXISTS (SELECT 1 FROM dbo.BloqueosHorario b WHERE b.EspacioDeportivoId = @EspacioDeportivoId AND b.Fecha = @Fecha AND b.Activo = 1 AND @HoraInicio < b.HoraFin AND @HoraFin > b.HoraInicio)
            RAISERROR('El horario esta bloqueado para ese espacio.', 16, 1);

        UPDATE dbo.Reservas
        SET Fecha = @Fecha, HoraInicio = @HoraInicio, HoraFin = @HoraFin, FechaActualizacion = SYSUTCDATETIME(), UsuarioActualizacion = @Usuario
        WHERE Id = @Id;

        IF @@ROWCOUNT = 0
            RAISERROR('No se encontro la reserva para mover.', 16, 1);

        IF @@ROWCOUNT > 0
        BEGIN
            DECLARE @EntidadIdAudit NVARCHAR(80);
            SET @EntidadIdAudit = CONVERT(NVARCHAR(80), @Id);
            EXEC dbo.Sp_Auditoria_Registrar @NegocioId = @NegocioId, @Modulo = N'RESERVAS', @Accion = N'MOVE', @Entidad = N'Reserva', @EntidadId = @EntidadIdAudit, @Usuario = @Usuario, @DetalleJson = NULL;
        END;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

-- SOURCE: 27_Calendario_No_Atencion_Sede.sql (linea 10)
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
            CONCAT(e.Nombre, N' - ', c.NombresORazonSocial) AS Titulo,
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
          AND CASE ((DATEDIFF(DAY, '19000101', f.Fecha) % 7) + 1)
                WHEN 1 THEN COALESCE(sha.AtiendeLunes, 1)
                WHEN 2 THEN COALESCE(sha.AtiendeMartes, 1)
                WHEN 3 THEN COALESCE(sha.AtiendeMiercoles, 1)
                WHEN 4 THEN COALESCE(sha.AtiendeJueves, 1)
                WHEN 5 THEN COALESCE(sha.AtiendeViernes, 1)
                WHEN 6 THEN COALESCE(sha.AtiendeSabado, 1)
                WHEN 7 THEN COALESCE(sha.AtiendeDomingo, 1)
              END = 1
          AND COALESCE(sha.HoraApertura, CAST('08:00' AS TIME)) > CAST('00:00' AS TIME)

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
          AND CASE ((DATEDIFF(DAY, '19000101', f.Fecha) % 7) + 1)
                WHEN 1 THEN COALESCE(sha.AtiendeLunes, 1)
                WHEN 2 THEN COALESCE(sha.AtiendeMartes, 1)
                WHEN 3 THEN COALESCE(sha.AtiendeMiercoles, 1)
                WHEN 4 THEN COALESCE(sha.AtiendeJueves, 1)
                WHEN 5 THEN COALESCE(sha.AtiendeViernes, 1)
                WHEN 6 THEN COALESCE(sha.AtiendeSabado, 1)
                WHEN 7 THEN COALESCE(sha.AtiendeDomingo, 1)
              END = 1
          AND COALESCE(sha.HoraCierre, CAST('23:00' AS TIME)) < CAST('23:59' AS TIME)

        ORDER BY Fecha, HoraInicio
        OPTION (MAXRECURSION 400);
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

-- SOURCE: 28_Espacios_Deporte_Suelo_Catalogos.sql (linea 89)
CREATE OR ALTER PROCEDURE dbo.Sp_Combos_TiposDeporte
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT td.Id, td.Nombre
        FROM dbo.TiposDeporte td
        WHERE td.Activo = 1
        ORDER BY td.Nombre;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

-- SOURCE: 28_Espacios_Deporte_Suelo_Catalogos.sql (linea 107)
CREATE OR ALTER PROCEDURE dbo.Sp_Combos_TiposSuelo
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT ts.Id, ts.Nombre
        FROM dbo.TiposSuelo ts
        WHERE ts.Activo = 1
        ORDER BY ts.Nombre;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

-- SOURCE: 29_Reservas_ValidarDisponibilidad_Modal.sql (linea 8)
CREATE OR ALTER PROCEDURE dbo.Sp_Reservas_ValidarDisponibilidad
    @NegocioId INT,
    @ReservaId INT = NULL,
    @EspacioDeportivoId INT,
    @Fecha DATE,
    @HoraInicio TIME,
    @HoraFin TIME
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF @HoraFin <= @HoraInicio
        BEGIN
            SELECT CAST(0 AS BIT) AS Disponible, CAST(N'La hora fin debe ser mayor que la hora inicio.' AS NVARCHAR(300)) AS Mensaje, CAST(NULL AS NVARCHAR(20)) AS ConflictoTipo, CAST(NULL AS INT) AS ConflictoId;
            RETURN;
        END;

        DECLARE @SedeId INT, @HoraApertura TIME, @HoraCierre TIME;
        DECLARE @AtiendeLunes BIT, @AtiendeMartes BIT, @AtiendeMiercoles BIT, @AtiendeJueves BIT, @AtiendeViernes BIT, @AtiendeSabado BIT, @AtiendeDomingo BIT;
        DECLARE @DiaSemana INT, @DiaHabilitado BIT;

        SELECT
            @SedeId = s.Id,
            @HoraApertura = COALESCE(sha.HoraApertura, CAST('08:00' AS TIME)),
            @HoraCierre = COALESCE(sha.HoraCierre, CAST('23:00' AS TIME)),
            @AtiendeLunes = COALESCE(sha.AtiendeLunes, 1),
            @AtiendeMartes = COALESCE(sha.AtiendeMartes, 1),
            @AtiendeMiercoles = COALESCE(sha.AtiendeMiercoles, 1),
            @AtiendeJueves = COALESCE(sha.AtiendeJueves, 1),
            @AtiendeViernes = COALESCE(sha.AtiendeViernes, 1),
            @AtiendeSabado = COALESCE(sha.AtiendeSabado, 1),
            @AtiendeDomingo = COALESCE(sha.AtiendeDomingo, 1)
        FROM dbo.EspaciosDeportivos e
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        LEFT JOIN dbo.SedeHorarioAtencion sha ON sha.SedeId = s.Id
        WHERE e.Id = @EspacioDeportivoId
          AND s.NegocioId = @NegocioId
          AND e.Estado = 1;

        IF @SedeId IS NULL
        BEGIN
            SELECT CAST(0 AS BIT) AS Disponible, CAST(N'El espacio deportivo no esta disponible para este negocio.' AS NVARCHAR(300)) AS Mensaje, CAST(NULL AS NVARCHAR(20)) AS ConflictoTipo, CAST(NULL AS INT) AS ConflictoId;
            RETURN;
        END;

        IF EXISTS (SELECT 1 FROM dbo.SedeFechasInhabilitadas sfi WHERE sfi.SedeId = @SedeId AND sfi.Fecha = @Fecha AND sfi.Activo = 1)
        BEGIN
            SELECT CAST(0 AS BIT) AS Disponible, CAST(N'La sede no atiende en la fecha seleccionada.' AS NVARCHAR(300)) AS Mensaje, CAST(NULL AS NVARCHAR(20)) AS ConflictoTipo, CAST(NULL AS INT) AS ConflictoId;
            RETURN;
        END;

        SET @DiaSemana = (DATEDIFF(DAY, '19000101', @Fecha) % 7) + 1;
        SET @DiaHabilitado = CASE @DiaSemana
            WHEN 1 THEN @AtiendeLunes
            WHEN 2 THEN @AtiendeMartes
            WHEN 3 THEN @AtiendeMiercoles
            WHEN 4 THEN @AtiendeJueves
            WHEN 5 THEN @AtiendeViernes
            WHEN 6 THEN @AtiendeSabado
            WHEN 7 THEN @AtiendeDomingo
            ELSE 0 END;

        IF COALESCE(@DiaHabilitado, 0) = 0
        BEGIN
            SELECT CAST(0 AS BIT) AS Disponible, CAST(N'La sede no atiende el dia seleccionado.' AS NVARCHAR(300)) AS Mensaje, CAST(NULL AS NVARCHAR(20)) AS ConflictoTipo, CAST(NULL AS INT) AS ConflictoId;
            RETURN;
        END;

        IF @HoraInicio < @HoraApertura OR @HoraFin > @HoraCierre
        BEGIN
            SELECT
                CAST(0 AS BIT) AS Disponible,
                CAST(
                    CONCAT(
                        N'Horario fuera de atencion. La sede atiende de ',
                        CONVERT(NVARCHAR(5), @HoraApertura, 108),
                        N' a ',
                        CONVERT(NVARCHAR(5), @HoraCierre, 108),
                        N'.'
                    )
                    AS NVARCHAR(300)
                ) AS Mensaje,
                CAST(NULL AS NVARCHAR(20)) AS ConflictoTipo,
                CAST(NULL AS INT) AS ConflictoId;
            RETURN;
        END;

        DECLARE @ReservaCruceId INT = NULL, @ReservaCruceInicio TIME = NULL, @ReservaCruceFin TIME = NULL;
        SELECT TOP 1
            @ReservaCruceId = r.Id,
            @ReservaCruceInicio = r.HoraInicio,
            @ReservaCruceFin = r.HoraFin
        FROM dbo.Reservas r
        WHERE r.EspacioDeportivoId = @EspacioDeportivoId
          AND r.Fecha = @Fecha
          AND r.Estado NOT IN (5, 6)
          AND (@ReservaId IS NULL OR r.Id <> @ReservaId)
          AND @HoraInicio < r.HoraFin
          AND @HoraFin > r.HoraInicio
        ORDER BY r.HoraInicio;

        IF @ReservaCruceId IS NOT NULL
        BEGIN
            SELECT
                CAST(0 AS BIT) AS Disponible,
                CAST(
                    CONCAT(
                        N'Cruce con reserva #',
                        @ReservaCruceId,
                        N' (',
                        CONVERT(NVARCHAR(5), @ReservaCruceInicio, 108),
                        N' - ',
                        CONVERT(NVARCHAR(5), @ReservaCruceFin, 108),
                        N').'
                    )
                    AS NVARCHAR(300)
                ) AS Mensaje,
                CAST(N'RESERVA' AS NVARCHAR(20)) AS ConflictoTipo,
                @ReservaCruceId AS ConflictoId;
            RETURN;
        END;

        DECLARE @BloqueoInicio TIME = NULL, @BloqueoFin TIME = NULL, @BloqueoMotivo NVARCHAR(250) = NULL;
        SELECT TOP 1
            @BloqueoInicio = b.HoraInicio,
            @BloqueoFin = b.HoraFin,
            @BloqueoMotivo = b.Motivo
        FROM dbo.BloqueosHorario b
        WHERE b.EspacioDeportivoId = @EspacioDeportivoId
          AND b.Fecha = @Fecha
          AND b.Activo = 1
          AND @HoraInicio < b.HoraFin
          AND @HoraFin > b.HoraInicio
        ORDER BY b.HoraInicio;

        IF @BloqueoInicio IS NOT NULL
        BEGIN
            SELECT
                CAST(0 AS BIT) AS Disponible,
                CAST(
                    CONCAT(
                        N'Horario bloqueado (',
                        CONVERT(NVARCHAR(5), @BloqueoInicio, 108),
                        N' - ',
                        CONVERT(NVARCHAR(5), @BloqueoFin, 108),
                        N'). Motivo: ',
                        COALESCE(@BloqueoMotivo, N'Sin detalle')
                    )
                    AS NVARCHAR(300)
                ) AS Mensaje,
                CAST(N'BLOQUEO' AS NVARCHAR(20)) AS ConflictoTipo,
                CAST(NULL AS INT) AS ConflictoId;
            RETURN;
        END;

        SELECT CAST(1 AS BIT) AS Disponible, CAST(N'Horario disponible.' AS NVARCHAR(300)) AS Mensaje, CAST(NULL AS NVARCHAR(20)) AS ConflictoTipo, CAST(NULL AS INT) AS ConflictoId;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

-- SOURCE: 30_Configuracion_Club_Monedas.sql (linea 119)
CREATE OR ALTER PROCEDURE dbo.Sp_Combos_Monedas
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        SELECT
            m.Id,
            CONCAT(m.Nombre, N' (', m.Codigo, N')') AS Nombre
        FROM dbo.Monedas m
        WHERE m.Activo = 1
        ORDER BY m.Nombre;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

-- SOURCE: 30_Configuracion_Club_Monedas.sql (linea 140)
CREATE OR ALTER PROCEDURE dbo.Sp_ConfiguracionClub_Obtener
    @NegocioId INT
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        SELECT
            n.Id,
            n.NombreComercial,
            n.RazonSocial,
            COALESCE(NULLIF(n.TipoDocumentoFiscal, N''), N'DNI') AS TipoDocumentoFiscal,
            COALESCE(NULLIF(n.NumeroDocumentoFiscal, N''), n.DocumentoFiscal) AS NumeroDocumentoFiscal,
            n.DireccionFiscal,
            COALESCE(n.MonedaId, 1) AS MonedaId
        FROM dbo.Negocios n
        WHERE n.Id = @NegocioId
          AND n.Activo = 1;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

-- SOURCE: 30_Configuracion_Club_Monedas.sql (linea 167)
CREATE OR ALTER PROCEDURE dbo.Sp_ConfiguracionClub_Actualizar
    @NegocioId INT,
    @NombreComercial NVARCHAR(200),
    @RazonSocial NVARCHAR(200) = NULL,
    @TipoDocumentoFiscal NVARCHAR(20) = NULL,
    @NumeroDocumentoFiscal NVARCHAR(20) = NULL,
    @DireccionFiscal NVARCHAR(250) = NULL,
    @MonedaId INT,
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        IF NOT EXISTS (SELECT 1 FROM dbo.Monedas WHERE Id = @MonedaId AND Activo = 1)
            RAISERROR('La moneda seleccionada no es valida.', 16, 1);

        UPDATE n
        SET
            n.NombreComercial = @NombreComercial,
            n.RazonSocial = NULLIF(@RazonSocial, N''),
            n.TipoDocumentoFiscal = NULLIF(@TipoDocumentoFiscal, N''),
            n.NumeroDocumentoFiscal = NULLIF(@NumeroDocumentoFiscal, N''),
            n.DireccionFiscal = NULLIF(@DireccionFiscal, N''),
            n.DocumentoFiscal = NULLIF(@NumeroDocumentoFiscal, N''),
            n.MonedaId = @MonedaId
        FROM dbo.Negocios n
        WHERE n.Id = @NegocioId
          AND n.Activo = 1;

        IF @@ROWCOUNT = 0
            RAISERROR('No se encontro el club para actualizar.', 16, 1);

        DECLARE @EntidadIdAuditoria NVARCHAR(80);
        SET @EntidadIdAuditoria = CONVERT(NVARCHAR(80), @NegocioId);

        EXEC dbo.Sp_Auditoria_Registrar
            @NegocioId = @NegocioId,
            @Modulo = N'CONFIGURACION',
            @Accion = N'EDIT',
            @Entidad = N'Negocio',
            @EntidadId = @EntidadIdAuditoria,
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

-- SOURCE: 31_Espacios_Tarifas_Base.sql (linea 29)
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

-- SOURCE: 31_Espacios_Tarifas_Base.sql (linea 72)
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

-- SOURCE: 31_Espacios_Tarifas_Base.sql (linea 185)
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

-- SOURCE: 32_Usuarios_Sede_Restriccion_Filtros.sql (linea 48)
CREATE OR ALTER PROCEDURE dbo.Sp_Combos_Sedes
    @NegocioId INT,
    @SedeId INT = NULL
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT s.Id, s.Nombre
        FROM dbo.Sedes s
        WHERE s.NegocioId = @NegocioId
          AND (@SedeId IS NULL OR s.Id = @SedeId)
          AND s.Activo = 1
        ORDER BY s.Nombre;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

-- SOURCE: 32_Usuarios_Sede_Restriccion_Filtros.sql (linea 70)
CREATE OR ALTER PROCEDURE dbo.Sp_Sedes_Listar
    @NegocioId INT,
    @SedeId INT = NULL
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT
            s.Id,
            s.Nombre,
            s.Direccion,
            STUFF((
                SELECT N', ' + cs.Nombre
                FROM dbo.SedeServicios ss
                INNER JOIN dbo.CatalogoServiciosSede cs ON cs.Id = ss.ServicioId
                WHERE ss.SedeId = s.Id
                  AND cs.Activo = 1
                ORDER BY cs.Nombre
                FOR XML PATH(''), TYPE
            ).value('.', 'NVARCHAR(MAX)'), 1, 2, N'') AS Servicios,
            COALESCE(scn.NotificacionesActivas, 1) AS NotificacionesActivas,
            scn.CorreoNotificacion,
            scn.WhatsappContacto,
            COALESCE(scn.PermiteChatWhatsapp, 0) AS PermiteChatWhatsapp,
            COALESCE(scn.MinutosAnticipacionRecordatorio, 90) AS MinutosAnticipacionRecordatorio,
            COALESCE(scn.MinutosToleranciaNoShow, 30) AS MinutosToleranciaNoShow,
            CONCAT(
                CASE WHEN COALESCE(sha.AtiendeLunes, 1) = 1 THEN N'Lun ' ELSE N'' END,
                CASE WHEN COALESCE(sha.AtiendeMartes, 1) = 1 THEN N'Mar ' ELSE N'' END,
                CASE WHEN COALESCE(sha.AtiendeMiercoles, 1) = 1 THEN N'Mie ' ELSE N'' END,
                CASE WHEN COALESCE(sha.AtiendeJueves, 1) = 1 THEN N'Jue ' ELSE N'' END,
                CASE WHEN COALESCE(sha.AtiendeViernes, 1) = 1 THEN N'Vie ' ELSE N'' END,
                CASE WHEN COALESCE(sha.AtiendeSabado, 1) = 1 THEN N'Sab ' ELSE N'' END,
                CASE WHEN COALESCE(sha.AtiendeDomingo, 1) = 1 THEN N'Dom' ELSE N'' END
            ) AS DiasAtencion,
            CONCAT(CONVERT(NVARCHAR(5), COALESCE(sha.HoraApertura, CAST('08:00' AS TIME)), 108), N' - ', CONVERT(NVARCHAR(5), COALESCE(sha.HoraCierre, CAST('23:00' AS TIME)), 108)) AS HorarioAtencion,
            (SELECT COUNT(1) FROM dbo.SedeFechasInhabilitadas sfi WHERE sfi.SedeId = s.Id AND sfi.Activo = 1) AS FechasNoLaborablesCount,
            s.Activo
        FROM dbo.Sedes s
        LEFT JOIN dbo.SedeConfiguracionNotificacion scn ON scn.SedeId = s.Id
        LEFT JOIN dbo.SedeHorarioAtencion sha ON sha.SedeId = s.Id
        WHERE s.NegocioId = @NegocioId
          AND (@SedeId IS NULL OR s.Id = @SedeId)
        ORDER BY s.Nombre;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

-- SOURCE: 32_Usuarios_Sede_Restriccion_Filtros.sql (linea 123)
CREATE OR ALTER PROCEDURE dbo.Sp_Espacios_Listar
    @NegocioId INT,
    @SedeId INT = NULL
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        DECLARE @SimboloMoneda NVARCHAR(10);
        SET @SimboloMoneda = N'S/';

        SELECT TOP (1) @SimboloMoneda = COALESCE(m.Simbolo, N'S/')
        FROM dbo.Negocios n
        LEFT JOIN dbo.Monedas m ON m.Id = n.MonedaId
        WHERE n.Id = @NegocioId;

        SELECT
            e.Id,
            e.Codigo,
            e.Nombre,
            s.Nombre AS Sede,
            td.Nombre AS TipoDeporte,
            ts.Nombre AS TipoSuelo,
            CASE e.Estado WHEN 1 THEN N'Activo' WHEN 2 THEN N'EnMantenimiento' ELSE N'Inactivo' END AS Estado,
            COALESCE
            (
                NULLIF
                (
                    STUFF
                    (
                        (
                            SELECT N' | '
                                + CASE t.DiaSemana
                                    WHEN 1 THEN N'Lun'
                                    WHEN 2 THEN N'Mar'
                                    WHEN 3 THEN N'Mie'
                                    WHEN 4 THEN N'Jue'
                                    WHEN 5 THEN N'Vie'
                                    WHEN 6 THEN N'Sab'
                                    WHEN 0 THEN N'Dom'
                                    ELSE N'Dia'
                                  END
                                + N' '
                                + CONVERT(NVARCHAR(5), t.HoraInicio, 108)
                                + N'-'
                                + CONVERT(NVARCHAR(5), t.HoraFin, 108)
                                + N' '
                                + @SimboloMoneda
                                + CONVERT(NVARCHAR(20), CAST(t.Precio AS DECIMAL(10,2)))
                            FROM dbo.Tarifas t
                            WHERE t.EspacioDeportivoId = e.Id
                              AND t.Activa = 1
                            ORDER BY
                                CASE t.DiaSemana
                                    WHEN 1 THEN 1
                                    WHEN 2 THEN 2
                                    WHEN 3 THEN 3
                                    WHEN 4 THEN 4
                                    WHEN 5 THEN 5
                                    WHEN 6 THEN 6
                                    WHEN 0 THEN 7
                                    ELSE 8
                                END,
                                t.HoraInicio,
                                t.HoraFin
                            FOR XML PATH(''), TYPE
                        ).value('.', 'NVARCHAR(MAX)'),
                        1, 3, N''
                    ),
                    N''
                ),
                N'Sin tarifa configurada'
            ) AS TarifaResumen
        FROM dbo.EspaciosDeportivos e
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        INNER JOIN dbo.TiposDeporte td ON td.Id = e.TipoDeporteId
        INNER JOIN dbo.TiposSuelo ts ON ts.Id = e.TipoSueloId
        WHERE s.NegocioId = @NegocioId
          AND (@SedeId IS NULL OR s.Id = @SedeId)
        ORDER BY s.Nombre, e.Nombre;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

-- SOURCE: 32_Usuarios_Sede_Restriccion_Filtros.sql (linea 211)
CREATE OR ALTER PROCEDURE dbo.Sp_Combos_EspaciosPorNegocio
    @NegocioId INT,
    @SedeId INT = NULL
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT
            e.Id,
            CONCAT(
                COALESCE(NULLIF(LTRIM(RTRIM(e.Codigo)), N''), N'S/C'),
                N' - ',
                e.Nombre,
                N' (',
                COALESCE(ts.Nombre, N'Sin suelo'),
                N')'
            )
        FROM dbo.EspaciosDeportivos e
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        LEFT JOIN dbo.TiposSuelo ts ON ts.Id = e.TipoSueloId
        WHERE s.NegocioId = @NegocioId
          AND (@SedeId IS NULL OR s.Id = @SedeId)
          AND e.Estado = 1
        ORDER BY e.Codigo, e.Nombre;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

-- SOURCE: 32_Usuarios_Sede_Restriccion_Filtros.sql (linea 244)
CREATE OR ALTER PROCEDURE dbo.Sp_Combos_ReservasPorNegocio
    @NegocioId INT,
    @SedeId INT = NULL
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

-- SOURCE: 32_Usuarios_Sede_Restriccion_Filtros.sql (linea 268)
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
            CAST(p.FormaPago AS NVARCHAR(20))
        FROM dbo.Pagos p
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

-- SOURCE: 32_Usuarios_Sede_Restriccion_Filtros.sql (linea 297)
CREATE OR ALTER PROCEDURE dbo.Sp_Comprobantes_Listar
    @NegocioId INT,
    @SedeId INT = NULL
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
        LEFT JOIN dbo.Reservas r ON r.Id = c.ReservaId
        LEFT JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
        LEFT JOIN dbo.Sedes s ON s.Id = e.SedeId
        WHERE c.NegocioId = @NegocioId
          AND (@SedeId IS NULL OR s.Id = @SedeId)
        ORDER BY c.FechaEmision DESC;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

-- SOURCE: 32_Usuarios_Sede_Restriccion_Filtros.sql (linea 329)
CREATE OR ALTER PROCEDURE dbo.Sp_Reportes_OcupacionPorEspacio
    @NegocioId INT,
    @FechaDesde DATE,
    @FechaHasta DATE,
    @SedeId INT = NULL
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT
            s.Nombre AS Sede,
            e.Nombre AS Espacio,
            COUNT(1) AS CantidadReservas,
            CAST(SUM(DATEDIFF(MINUTE, r.HoraInicio, r.HoraFin)) / 60.0 AS DECIMAL(10,2)) AS HorasReservadas,
            SUM(r.Total) AS MontoReservado,
            COALESCE(SUM(p.Monto), 0) AS MontoCobrado
        FROM dbo.Reservas r
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        LEFT JOIN dbo.Pagos p ON p.ReservaId = r.Id
        WHERE s.NegocioId = @NegocioId
          AND (@SedeId IS NULL OR s.Id = @SedeId)
          AND r.Fecha >= @FechaDesde
          AND r.Fecha <= @FechaHasta
          AND r.Estado NOT IN (5, 6)
        GROUP BY s.Nombre, e.Nombre
        ORDER BY HorasReservadas DESC, CantidadReservas DESC;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

-- SOURCE: 32_Usuarios_Sede_Restriccion_Filtros.sql (linea 365)
CREATE OR ALTER PROCEDURE dbo.Sp_Reportes_IngresosPorDia
    @NegocioId INT,
    @FechaDesde DATE,
    @FechaHasta DATE,
    @SedeId INT = NULL
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT
            r.Fecha,
            COUNT(DISTINCT r.Id) AS CantidadReservas,
            COALESCE(SUM(p.Monto), 0) AS Ingresos
        FROM dbo.Reservas r
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        LEFT JOIN dbo.Pagos p ON p.ReservaId = r.Id
        WHERE s.NegocioId = @NegocioId
          AND (@SedeId IS NULL OR s.Id = @SedeId)
          AND r.Fecha >= @FechaDesde
          AND r.Fecha <= @FechaHasta
        GROUP BY r.Fecha
        ORDER BY r.Fecha ASC;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

-- SOURCE: 32_Usuarios_Sede_Restriccion_Filtros.sql (linea 397)
CREATE OR ALTER PROCEDURE dbo.Sp_Panel_ObtenerMetricas
    @NegocioId INT,
    @Fecha DATE,
    @SedeId INT = NULL
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        DECLARE @TotalSedes INT = 0, @TotalEspacios INT = 0, @ReservasHoy INT = 0;
        DECLARE @IngresosHoy DECIMAL(12,2) = 0, @OcupacionHoyPct DECIMAL(5,2) = 0;
        DECLARE @NoShowMes INT = 0, @TicketPromedioMes DECIMAL(12,2) = 0;

        SELECT @TotalSedes = COUNT(1)
        FROM dbo.Sedes s
        WHERE s.NegocioId = @NegocioId
          AND (@SedeId IS NULL OR s.Id = @SedeId);

        SELECT @TotalEspacios = COUNT(1)
        FROM dbo.EspaciosDeportivos e
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        WHERE s.NegocioId = @NegocioId
          AND (@SedeId IS NULL OR s.Id = @SedeId);

        SELECT @ReservasHoy = COUNT(1)
        FROM dbo.Reservas r
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        WHERE s.NegocioId = @NegocioId
          AND (@SedeId IS NULL OR s.Id = @SedeId)
          AND r.Fecha = @Fecha;

        SELECT @IngresosHoy = COALESCE(SUM(p.Monto), 0)
        FROM dbo.Pagos p
        INNER JOIN dbo.Reservas r ON r.Id = p.ReservaId
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        WHERE s.NegocioId = @NegocioId
          AND (@SedeId IS NULL OR s.Id = @SedeId)
          AND CAST(p.FechaPago AS DATE) = @Fecha;

        IF @TotalEspacios > 0
        BEGIN
            DECLARE @EspaciosOcupadosHoy INT = 0;
            SELECT @EspaciosOcupadosHoy = COUNT(DISTINCT r.EspacioDeportivoId)
            FROM dbo.Reservas r
            INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
            INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
            WHERE s.NegocioId = @NegocioId
              AND (@SedeId IS NULL OR s.Id = @SedeId)
              AND r.Fecha = @Fecha
              AND r.Estado NOT IN (5, 6);
            SET @OcupacionHoyPct = CAST((@EspaciosOcupadosHoy * 100.0) / @TotalEspacios AS DECIMAL(5,2));
        END;

        SELECT @NoShowMes = COUNT(1)
        FROM dbo.Reservas r
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        WHERE s.NegocioId = @NegocioId
          AND (@SedeId IS NULL OR s.Id = @SedeId)
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
          AND (@SedeId IS NULL OR s.Id = @SedeId)
          AND YEAR(p.FechaPago) = YEAR(@Fecha)
          AND MONTH(p.FechaPago) = MONTH(@Fecha);

        SELECT @ReservasPagadasMes = COUNT(DISTINCT p.ReservaId)
        FROM dbo.Pagos p
        INNER JOIN dbo.Reservas r ON r.Id = p.ReservaId
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        WHERE s.NegocioId = @NegocioId
          AND (@SedeId IS NULL OR s.Id = @SedeId)
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

-- SOURCE: 32_Usuarios_Sede_Restriccion_Filtros.sql (linea 503)
CREATE OR ALTER PROCEDURE dbo.Sp_Promociones_Listar
    @NegocioId INT,
    @SedeId INT = NULL
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
          AND (@SedeId IS NULL OR p.SedeId = @SedeId OR (p.SedeId IS NULL AND p.EspacioDeportivoId IS NULL))
        ORDER BY p.FechaInicio DESC, p.Nombre;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

-- SOURCE: 32_Usuarios_Sede_Restriccion_Filtros.sql (linea 536)
CREATE OR ALTER PROCEDURE dbo.Sp_UsuariosNegocio_Listar
    @NegocioId INT,
    @SedeId INT = NULL
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT
            un.Id AS UsuarioNegocioId,
            un.UsuarioId,
            COALESCE(u.Nombres, N'') AS Nombres,
            COALESCE(u.Apellidos, N'') AS Apellidos,
            COALESCE(u.Email, N'') AS Correo,
            un.RolNegocio,
            un.Activo,
            un.SedeId,
            COALESCE(s.Nombre, N'') AS SedeNombre
        FROM dbo.UsuariosNegocio un
        INNER JOIN dbo.AspNetUsers u ON u.Id = un.UsuarioId
        LEFT JOIN dbo.Sedes s ON s.Id = un.SedeId
        WHERE un.NegocioId = @NegocioId
          AND (@SedeId IS NULL OR un.SedeId = @SedeId)
        ORDER BY un.Activo DESC, u.Email;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

-- SOURCE: 32_Usuarios_Sede_Restriccion_Filtros.sql (linea 568)
CREATE OR ALTER PROCEDURE dbo.Sp_UsuariosNegocio_AsignarPorCorreo
    @NegocioId INT,
    @Correo NVARCHAR(256),
    @RolNegocio INT,
    @SedeId INT = NULL,
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        DECLARE @UsuarioId NVARCHAR(450);
        SELECT TOP (1) @UsuarioId = u.Id
        FROM dbo.AspNetUsers u
        WHERE u.NormalizedEmail = UPPER(@Correo);

        IF @UsuarioId IS NULL
            RAISERROR('No existe usuario con ese correo en el sistema.', 16, 1);

        IF @RolNegocio = 1
            SET @SedeId = NULL;

        IF @RolNegocio <> 1 AND @SedeId IS NULL
            RAISERROR('La sede es obligatoria para usuarios no administradores.', 16, 1);

        IF @SedeId IS NOT NULL
        BEGIN
            IF NOT EXISTS (SELECT 1 FROM dbo.Sedes s WHERE s.Id = @SedeId AND s.NegocioId = @NegocioId)
                RAISERROR('La sede no pertenece al negocio seleccionado.', 16, 1);
        END;

        IF EXISTS (SELECT 1 FROM dbo.UsuariosNegocio WHERE NegocioId = @NegocioId AND UsuarioId = @UsuarioId)
        BEGIN
            UPDATE dbo.UsuariosNegocio
            SET RolNegocio = @RolNegocio,
                SedeId = @SedeId,
                Activo = 1
            WHERE NegocioId = @NegocioId
              AND UsuarioId = @UsuarioId;
        END
        ELSE
        BEGIN
            INSERT INTO dbo.UsuariosNegocio (UsuarioId, NegocioId, RolNegocio, SedeId, Activo)
            VALUES (@UsuarioId, @NegocioId, @RolNegocio, @SedeId, 1);
        END;

        DECLARE @UsuarioNegocioId INT;
        SELECT TOP (1) @UsuarioNegocioId = Id FROM dbo.UsuariosNegocio WHERE NegocioId = @NegocioId AND UsuarioId = @UsuarioId;

        DECLARE @EntidadIdAudit NVARCHAR(80);
        SET @EntidadIdAudit = CONVERT(NVARCHAR(80), @UsuarioNegocioId);
        EXEC dbo.Sp_Auditoria_Registrar @NegocioId = @NegocioId, @Modulo = N'USUARIOS', @Accion = N'CREATE', @Entidad = N'UsuarioNegocio', @EntidadId = @EntidadIdAudit, @Usuario = @Usuario, @DetalleJson = NULL;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

-- SOURCE: 32_Usuarios_Sede_Restriccion_Filtros.sql (linea 628)
CREATE OR ALTER PROCEDURE dbo.Sp_UsuariosNegocio_ActualizarRol
    @NegocioId INT,
    @UsuarioNegocioId INT,
    @RolNegocio INT,
    @SedeId INT = NULL,
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF @RolNegocio = 1
            SET @SedeId = NULL;

        IF @RolNegocio <> 1 AND @SedeId IS NULL
            RAISERROR('La sede es obligatoria para usuarios no administradores.', 16, 1);

        IF @SedeId IS NOT NULL
        BEGIN
            IF NOT EXISTS (SELECT 1 FROM dbo.Sedes s WHERE s.Id = @SedeId AND s.NegocioId = @NegocioId)
                RAISERROR('La sede no pertenece al negocio seleccionado.', 16, 1);
        END;

        UPDATE dbo.UsuariosNegocio
        SET RolNegocio = @RolNegocio,
            SedeId = @SedeId
        WHERE Id = @UsuarioNegocioId
          AND NegocioId = @NegocioId;

        IF @@ROWCOUNT = 0
            RAISERROR('No se encontro el usuario del negocio para actualizar rol.', 16, 1);

        IF @@ROWCOUNT > 0
        BEGIN
            DECLARE @EntidadIdAudit NVARCHAR(80);
            SET @EntidadIdAudit = CONVERT(NVARCHAR(80), @UsuarioNegocioId);
            EXEC dbo.Sp_Auditoria_Registrar @NegocioId = @NegocioId, @Modulo = N'USUARIOS', @Accion = N'EDIT', @Entidad = N'UsuarioNegocio', @EntidadId = @EntidadIdAudit, @Usuario = @Usuario, @DetalleJson = NULL;
        END;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

-- SOURCE: 32_Usuarios_Sede_Restriccion_Filtros.sql (linea 674)
CREATE OR ALTER PROCEDURE dbo.Sp_Seguridad_ObtenerContextoModulo
    @UsuarioId NVARCHAR(450),
    @NegocioId INT,
    @ModuloCodigo NVARCHAR(50)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        DECLARE @UsuarioNegocioId INT, @RolNegocio INT, @NegocioNombre NVARCHAR(200), @ModuloId INT, @ModuloNombre NVARCHAR(120);
        DECLARE @PuedeVer BIT = 0, @PuedeCrear BIT = 0, @PuedeEditar BIT = 0, @PuedeEliminar BIT = 0;
        DECLARE @EstadoSuscripcion INT, @EsPrueba BIT, @FechaFinPrueba DATE, @FechaFinPlan DATE;
        DECLARE @Hoy DATE = CAST(SYSUTCDATETIME() AS DATE);
        DECLARE @SedeIdAsignada INT = NULL, @EsAdministrador BIT = 0;

        SELECT
            @UsuarioNegocioId = un.Id,
            @RolNegocio = un.RolNegocio,
            @SedeIdAsignada = un.SedeId,
            @NegocioNombre = n.NombreComercial
        FROM dbo.UsuariosNegocio un
        INNER JOIN dbo.Negocios n ON n.Id = un.NegocioId
        WHERE un.UsuarioId = @UsuarioId
          AND un.NegocioId = @NegocioId
          AND un.Activo = 1
          AND n.Activo = 1;

        SET @EsAdministrador = CASE WHEN @RolNegocio = 1 THEN 1 ELSE 0 END;
        IF @EsAdministrador = 1
            SET @SedeIdAsignada = NULL;

        IF @UsuarioNegocioId IS NULL
        BEGIN
            SELECT CAST(0 AS BIT), @NegocioId, N'', @ModuloCodigo, N'', N'', CAST(0 AS BIT), CAST(0 AS BIT), CAST(0 AS BIT), CAST(0 AS BIT), N'Usuario sin acceso al negocio', CAST(NULL AS INT), CAST(0 AS BIT);
            RETURN;
        END;

        IF OBJECT_ID(N'dbo.NegociosSuscripcion', N'U') IS NOT NULL
        BEGIN
            SELECT
                @EstadoSuscripcion = ns.EstadoSuscripcion,
                @EsPrueba = ns.EsPrueba,
                @FechaFinPrueba = ns.FechaFinPrueba,
                @FechaFinPlan = ns.FechaFinPlan
            FROM dbo.NegociosSuscripcion ns
            WHERE ns.NegocioId = @NegocioId;

            IF @EstadoSuscripcion = 1 AND @EsPrueba = 1 AND @FechaFinPrueba IS NOT NULL AND @FechaFinPrueba < @Hoy
            BEGIN
                UPDATE dbo.NegociosSuscripcion
                SET EstadoSuscripcion = 3,
                    EsPrueba = 0,
                    FechaActualizacion = SYSUTCDATETIME(),
                    UsuarioActualizacion = @UsuarioId
                WHERE NegocioId = @NegocioId;
                SET @EstadoSuscripcion = 3;
            END;

            IF @EstadoSuscripcion = 2 AND @FechaFinPlan IS NOT NULL AND @FechaFinPlan < @Hoy
            BEGIN
                UPDATE dbo.NegociosSuscripcion
                SET EstadoSuscripcion = 3,
                    FechaActualizacion = SYSUTCDATETIME(),
                    UsuarioActualizacion = @UsuarioId
                WHERE NegocioId = @NegocioId;
                SET @EstadoSuscripcion = 3;
            END;

            IF @EstadoSuscripcion IN (3, 4)
            BEGIN
                SELECT
                    CAST(0 AS BIT),
                    @NegocioId,
                    @NegocioNombre,
                    @ModuloCodigo,
                    N'',
                    CAST(@RolNegocio AS NVARCHAR(20)),
                    CAST(0 AS BIT),
                    CAST(0 AS BIT),
                    CAST(0 AS BIT),
                    CAST(0 AS BIT),
                    N'La suscripcion del negocio esta vencida o suspendida. Activa un plan para continuar operando.',
                    @SedeIdAsignada,
                    @EsAdministrador;
                RETURN;
            END;
        END;

        SELECT @ModuloId = m.Id, @ModuloNombre = m.Nombre
        FROM dbo.ModulosSistema m
        WHERE m.Codigo = @ModuloCodigo AND m.Activo = 1;

        IF @ModuloId IS NULL
        BEGIN
            SELECT CAST(0 AS BIT), @NegocioId, @NegocioNombre, @ModuloCodigo, N'', CAST(@RolNegocio AS NVARCHAR(20)), CAST(0 AS BIT), CAST(0 AS BIT), CAST(0 AS BIT), CAST(0 AS BIT), N'Modulo no configurado', @SedeIdAsignada, @EsAdministrador;
            RETURN;
        END;

        SELECT
            @PuedeVer = rp.PuedeVer,
            @PuedeCrear = rp.PuedeCrear,
            @PuedeEditar = rp.PuedeEditar,
            @PuedeEliminar = rp.PuedeEliminar
        FROM dbo.RolesNegocioPermiso rp
        WHERE rp.RolNegocio = @RolNegocio
          AND rp.ModuloSistemaId = @ModuloId;

        SELECT
            @PuedeVer = COALESCE(up.PuedeVer, @PuedeVer),
            @PuedeCrear = COALESCE(up.PuedeCrear, @PuedeCrear),
            @PuedeEditar = COALESCE(up.PuedeEditar, @PuedeEditar),
            @PuedeEliminar = COALESCE(up.PuedeEliminar, @PuedeEliminar)
        FROM dbo.UsuariosNegocioPermiso up
        WHERE up.UsuarioNegocioId = @UsuarioNegocioId
          AND up.ModuloSistemaId = @ModuloId;

        SELECT
            CAST(CASE WHEN @PuedeVer = 1 THEN 1 ELSE 0 END AS BIT) AS Autorizado,
            @NegocioId,
            @NegocioNombre,
            @ModuloCodigo,
            @ModuloNombre,
            CAST(@RolNegocio AS NVARCHAR(20)) AS RolActual,
            @PuedeVer,
            @PuedeCrear,
            @PuedeEditar,
            @PuedeEliminar,
            CAST(NULL AS NVARCHAR(200)) AS Mensaje,
            @SedeIdAsignada AS SedeIdAsignada,
            @EsAdministrador AS EsAdministrador;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO
